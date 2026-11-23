# -*- coding: utf-8 -*-
"""
식품안전나라(식품의약품안전처) OpenAPI 수집 클라이언트 및 SQLite 로컬 캐시(이어받기) 모듈

연동 대상 실제 OpenAPI 서비스:
1) I2500 : 영업 중인 인허가 업소 정보 (INDUTY_CD_NM=식육가공업, 식육포장처리업, 축산물가공업 합집합)
2) I1320 : 축산물 식육포장처리업 영업허가대장 (LCNS_NO 기준 누락 업소 보완 병합, 폐업·취소 제외)
3) C006  : 축산물 품목제조보고(원재료) (LCNS_NO 개별 조회 또는 전체 페이지네이션 후 매핑)
"""

from __future__ import annotations

import contextlib
import functools
import logging
import sqlite3
import threading
import time
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Set, Tuple

import requests
from requests.adapters import HTTPAdapter

try:
    from tenacity import (
        retry,
        retry_if_exception_type,
        stop_after_attempt,
        wait_exponential,
    )

    _HAS_TENACITY = True
except ImportError:  # pragma: no cover - tenacity 미설치 환경에서도 동일 지수 백오프 보장
    _HAS_TENACITY = False

    def _fallback_retry(exc_type: type, max_attempts: int = 4):
        def decorator(func):
            @functools.wraps(func)
            def wrapper(*args, **kwargs):
                delay = 1.2
                for attempt in range(1, max_attempts + 1):
                    try:
                        return func(*args, **kwargs)
                    except exc_type:
                        if attempt >= max_attempts:
                            raise
                        time.sleep(min(delay, 10.0))
                        delay *= 2.0

            return wrapper

        return decorator

logger = logging.getLogger(__name__)

BASE_API_URL = "http://openapi.foodsafetykorea.go.kr/api"
MAX_PAGE_SIZE = 1000

I2500_TARGET_INDUSTRIES: Tuple[str, ...] = (
    "식육가공업",
    "식육포장처리업",
    "축산물가공업",
)

# I1320 응답 내 폐업·취소 여부 판별을 위한 상태 컬럼 후보 및 제외 키워드
CLOSED_STATUS_FIELDS: Tuple[str, ...] = (
    "CLSBIZ_DVS_CD_NM",
    "BSN_STATE_NM",
    "TRD_STATE_NM",
    "STATE_NM",
    "UNITY_SIGN_NM",
    "BIZ_STAT_NM",
)
CLOSED_DATE_FIELDS: Tuple[str, ...] = (
    "CLSBIZ_DT",
    "CANCL_DT",
    "CLOSURE_DT",
    "RTRCN_DT",
)
CLOSED_KEYWORDS: Tuple[str, ...] = (
    "폐업",
    "취소",
    "말소",
    "폐쇄",
    "직권폐업",
    "허가취소",
    "영업폐쇄",
)

ProgressCallback = Optional[Callable[[str, float, str], None]]


class FoodSafetyAPIError(Exception):
    """식품안전나라 OpenAPI 기본 예외"""


class InvalidApiKeyError(FoodSafetyAPIError):
    """유효하지 않은 인증키 예외 (재시도 불가)"""


class QuotaExceededError(FoodSafetyAPIError):
    """일일/시간당 호출 한도 초과 예외 (재시도 불가, 캐시 저장 후 중단)"""


class TransientAPIError(FoodSafetyAPIError):
    """일시적 네트워크·게이트웨이 오류 (tenacity 자동 재시도 대상)"""


def is_active_license_row(row: Dict[str, Any]) -> bool:
    """I1320 등 인허가 대장 행에서 폐업·취소 상태 레코드를 제외한다."""
    for date_col in CLOSED_DATE_FIELDS:
        val = str(row.get(date_col, "") or "").strip()
        if val and val not in ("0", "00000000", "-", "null", "None"):
            return False

    for status_col in CLOSED_STATUS_FIELDS:
        status_val = str(row.get(status_col, "") or "").strip()
        if status_val and any(kw in status_val for kw in CLOSED_KEYWORDS):
            return False

    return True


class SQLiteCacheStore:
    """수집 진행 상태(페이지/인허가번호 단위) 및 원본·가공 데이터를 영속화하는 SQLite 캐시"""

    def __init__(self, db_path: Path) -> None:
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        self._init_schema()

    @contextlib.contextmanager
    def _connect(self):
        """Windows 환경에서 파일 잠금(WinError 32)이 남지 않도록 트랜잭션 커밋 후 즉시 close 보장"""
        conn = sqlite3.connect(str(self.db_path), timeout=30.0)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA synchronous=NORMAL;")
        try:
            with conn:
                yield conn
        finally:
            conn.close()

    def _init_schema(self) -> None:
        with self._lock, self._connect() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS collection_progress (
                    task_key TEXT PRIMARY KEY,
                    total_count INTEGER NOT NULL DEFAULT 0,
                    last_end_idx INTEGER NOT NULL DEFAULT 0,
                    is_completed INTEGER NOT NULL DEFAULT 0,
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS raw_businesses (
                    lcns_no TEXT PRIMARY KEY,
                    bssh_nm TEXT NOT NULL DEFAULT '',
                    induty_cd_nm TEXT NOT NULL DEFAULT '',
                    prsdnt_nm TEXT NOT NULL DEFAULT '',
                    telno TEXT NOT NULL DEFAULT '',
                    addr TEXT NOT NULL DEFAULT '',
                    prms_dt TEXT NOT NULL DEFAULT '',
                    source TEXT NOT NULL DEFAULT '',
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS c006_checked_licenses (
                    lcns_no TEXT PRIMARY KEY,
                    item_count INTEGER NOT NULL DEFAULT 0,
                    checked_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS raw_c006_items (
                    lcns_no TEXT NOT NULL,
                    prdlst_report_no TEXT NOT NULL,
                    prdlst_nm TEXT NOT NULL,
                    bssh_nm TEXT NOT NULL DEFAULT '',
                    prdlst_dcnm TEXT NOT NULL DEFAULT '',
                    rawmtrl_nm TEXT NOT NULL DEFAULT '',
                    updated_at TEXT NOT NULL,
                    PRIMARY KEY (lcns_no, prdlst_report_no, prdlst_nm)
                );

                CREATE INDEX IF NOT EXISTS idx_c006_lcns_no
                    ON raw_c006_items (lcns_no);
                """
            )

    def get_task_progress(self, task_key: str) -> Tuple[int, int, bool]:
        with self._lock, self._connect() as conn:
            row = conn.execute(
                "SELECT total_count, last_end_idx, is_completed FROM collection_progress WHERE task_key = ?",
                (task_key,),
            ).fetchone()
            if not row:
                return 0, 0, False
            return int(row["total_count"]), int(row["last_end_idx"]), bool(row["is_completed"])

    def update_task_progress(
        self, task_key: str, total_count: int, last_end_idx: int, is_completed: bool
    ) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        with self._lock, self._connect() as conn:
            conn.execute(
                """
                INSERT INTO collection_progress (task_key, total_count, last_end_idx, is_completed, updated_at)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(task_key) DO UPDATE SET
                    total_count = excluded.total_count,
                    last_end_idx = excluded.last_end_idx,
                    is_completed = excluded.is_completed,
                    updated_at = excluded.updated_at
                """,
                (task_key, int(total_count), int(last_end_idx), 1 if is_completed else 0, now),
            )

    def upsert_i2500_rows(self, rows: Sequence[Dict[str, Any]], query_industry: str) -> int:
        if not rows:
            return 0
        now = datetime.now().isoformat(timespec="seconds")
        saved = 0
        with self._lock, self._connect() as conn:
            for r in rows:
                lcns_no = str(r.get("LCNS_NO", "") or "").strip()
                if not lcns_no:
                    continue
                bssh_nm = str(r.get("BSSH_NM", "") or "").strip()
                induty = str(r.get("INDUTY_CD_NM", "") or query_industry).strip()
                prsdnt = str(r.get("PRSDNT_NM", "") or "").strip()
                telno = str(r.get("TELNO", "") or "").strip()
                addr = str(r.get("ADDR", "") or r.get("LOCP_ADDR", "") or "").strip()
                prms_dt = str(r.get("PRMS_DT", "") or "").strip()

                existing = conn.execute(
                    "SELECT bssh_nm, induty_cd_nm, prsdnt_nm, telno, addr, prms_dt, source FROM raw_businesses WHERE lcns_no = ?",
                    (lcns_no,),
                ).fetchone()

                if existing is None:
                    conn.execute(
                        """
                        INSERT INTO raw_businesses
                            (lcns_no, bssh_nm, induty_cd_nm, prsdnt_nm, telno, addr, prms_dt, source, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, 'I2500', ?)
                        """,
                        (lcns_no, bssh_nm, induty, prsdnt, telno, addr, prms_dt, now),
                    )
                else:
                    # I2500은 메인 소스이므로 기존 값 보완·우선 반영
                    merged_induty = induty or existing["induty_cd_nm"]
                    if existing["induty_cd_nm"] and induty and induty not in existing["induty_cd_nm"]:
                        merged_induty = f"{induty}, {existing['induty_cd_nm']}"
                    conn.execute(
                        """
                        UPDATE raw_businesses
                        SET bssh_nm = ?,
                            induty_cd_nm = ?,
                            prsdnt_nm = ?,
                            telno = ?,
                            addr = ?,
                            prms_dt = ?,
                            source = 'I2500',
                            updated_at = ?
                        WHERE lcns_no = ?
                        """,
                        (
                            bssh_nm or existing["bssh_nm"],
                            merged_induty,
                            prsdnt or existing["prsdnt_nm"],
                            telno or existing["telno"],
                            addr or existing["addr"],
                            prms_dt or existing["prms_dt"],
                            now,
                            lcns_no,
                        ),
                    )
                saved += 1
        return saved

    def merge_i1320_rows(self, rows: Sequence[Dict[str, Any]]) -> Tuple[int, int]:
        """
        I1320(축산물 식육포장처리업 영업허가대장) 레코드를 병합한다.
        - 폐업·취소 상태 레코드 제외
        - I2500에 없는 LCNS_NO는 신규 추가
        - I2500에 이미 존재하는 LCNS_NO는 누락된 전화번호/대표자/주소만 보조 보완
        Returns: (신규 보완 건수, 기존 필드 보완 건수)
        """
        if not rows:
            return 0, 0
        now = datetime.now().isoformat(timespec="seconds")
        inserted = 0
        supplemented = 0

        with self._lock, self._connect() as conn:
            for r in rows:
                if not is_active_license_row(r):
                    continue
                lcns_no = str(r.get("LCNS_NO", "") or "").strip()
                if not lcns_no:
                    continue

                bssh_nm = str(r.get("BSSH_NM", "") or "").strip()
                induty = str(
                    r.get("INDUTY_CD_NM", "") or r.get("INDUTY_NM", "") or "식육포장처리업"
                ).strip()
                prsdnt = str(r.get("PRSDNT_NM", "") or "").strip()
                telno = str(r.get("TELNO", "") or "").strip()
                # I1320 실제 응답 주소 컬럼명은 LOCP_ADDR (혹시 ADDR로 내려오는 경우도 대비)
                addr = str(r.get("LOCP_ADDR", "") or r.get("ADDR", "") or "").strip()
                prms_dt = str(r.get("PRMS_DT", "") or "").strip()

                existing = conn.execute(
                    "SELECT bssh_nm, induty_cd_nm, prsdnt_nm, telno, addr, prms_dt, source FROM raw_businesses WHERE lcns_no = ?",
                    (lcns_no,),
                ).fetchone()

                if existing is None:
                    conn.execute(
                        """
                        INSERT INTO raw_businesses
                            (lcns_no, bssh_nm, induty_cd_nm, prsdnt_nm, telno, addr, prms_dt, source, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, 'I1320', ?)
                        """,
                        (lcns_no, bssh_nm, induty, prsdnt, telno, addr, prms_dt, now),
                    )
                    inserted += 1
                else:
                    new_tel = existing["telno"] or telno
                    new_prsdnt = existing["prsdnt_nm"] or prsdnt
                    new_addr = existing["addr"] or addr
                    new_prms = existing["prms_dt"] or prms_dt
                    new_bssh = existing["bssh_nm"] or bssh_nm
                    cur_induty = existing["induty_cd_nm"] or ""
                    if "식육포장처리업" not in cur_induty and "식육포장처리업" in induty:
                        new_induty = f"{cur_induty}, 식육포장처리업".strip(", ")
                    else:
                        new_induty = cur_induty or induty

                    if (
                        new_tel != existing["telno"]
                        or new_prsdnt != existing["prsdnt_nm"]
                        or new_addr != existing["addr"]
                        or new_induty != existing["induty_cd_nm"]
                    ):
                        src = (
                            "I2500+I1320"
                            if "I2500" in (existing["source"] or "")
                            else existing["source"]
                        )
                        conn.execute(
                            """
                            UPDATE raw_businesses
                            SET bssh_nm = ?,
                                induty_cd_nm = ?,
                                prsdnt_nm = ?,
                                telno = ?,
                                addr = ?,
                                prms_dt = ?,
                                source = ?,
                                updated_at = ?
                            WHERE lcns_no = ?
                            """,
                            (
                                new_bssh,
                                new_induty,
                                new_prsdnt,
                                new_tel,
                                new_addr,
                                new_prms,
                                src,
                                now,
                                lcns_no,
                            ),
                        )
                        supplemented += 1

        return inserted, supplemented

    def upsert_c006_rows(
        self, rows: Sequence[Dict[str, Any]], allowed_licenses: Optional[Set[str]] = None
    ) -> int:
        if not rows:
            return 0
        now = datetime.now().isoformat(timespec="seconds")
        saved = 0
        with self._lock, self._connect() as conn:
            for r in rows:
                lcns_no = str(r.get("LCNS_NO", "") or "").strip()
                if not lcns_no:
                    continue
                if allowed_licenses is not None and lcns_no not in allowed_licenses:
                    continue

                prdlst_nm = str(r.get("PRDLST_NM", "") or "").strip()
                report_no = str(r.get("PRDLST_REPORT_NO", "") or "").strip()
                if not prdlst_nm and not report_no:
                    continue

                bssh_nm = str(r.get("BSSH_NM", "") or "").strip()
                dcnm = str(r.get("PRDLST_DCNM", "") or "").strip()
                rawmtrl = str(r.get("RAWMTRL_NM", "") or "").strip()

                existing = conn.execute(
                    """
                    SELECT rawmtrl_nm, prdlst_dcnm
                    FROM raw_c006_items
                    WHERE lcns_no = ? AND prdlst_report_no = ? AND prdlst_nm = ?
                    """,
                    (lcns_no, report_no, prdlst_nm),
                ).fetchone()

                if existing is None:
                    conn.execute(
                        """
                        INSERT INTO raw_c006_items
                            (lcns_no, prdlst_report_no, prdlst_nm, bssh_nm, prdlst_dcnm, rawmtrl_nm, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                        """,
                        (lcns_no, report_no, prdlst_nm, bssh_nm, dcnm, rawmtrl, now),
                    )
                else:
                    # 동일 품목제조보고번호의 원재료 행이 여러 줄로 나뉜 경우 원재료 병합
                    merged_raw = existing["rawmtrl_nm"] or ""
                    if rawmtrl and rawmtrl not in merged_raw:
                        merged_raw = f"{merged_raw}, {rawmtrl}".strip(", ")
                    conn.execute(
                        """
                        UPDATE raw_c006_items
                        SET prdlst_dcnm = ?,
                            rawmtrl_nm = ?,
                            updated_at = ?
                        WHERE lcns_no = ? AND prdlst_report_no = ? AND prdlst_nm = ?
                        """,
                        (
                            dcnm or existing["prdlst_dcnm"],
                            merged_raw,
                            now,
                            lcns_no,
                            report_no,
                            prdlst_nm,
                        ),
                    )
                saved += 1
        return saved

    def mark_c006_license_checked(self, lcns_no: str, item_count: int) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        with self._lock, self._connect() as conn:
            conn.execute(
                """
                INSERT INTO c006_checked_licenses (lcns_no, item_count, checked_at)
                VALUES (?, ?, ?)
                ON CONFLICT(lcns_no) DO UPDATE SET
                    item_count = excluded.item_count,
                    checked_at = excluded.checked_at
                """,
                (lcns_no, int(item_count), now),
            )

    def get_checked_c006_licenses(self) -> Set[str]:
        with self._lock, self._connect() as conn:
            rows = conn.execute("SELECT lcns_no FROM c006_checked_licenses").fetchall()
            return {str(r["lcns_no"]) for r in rows}

    def load_all_businesses(self) -> List[Dict[str, Any]]:
        with self._lock, self._connect() as conn:
            rows = conn.execute(
                """
                SELECT
                    lcns_no AS LCNS_NO,
                    bssh_nm AS BSSH_NM,
                    induty_cd_nm AS INDUTY_CD_NM,
                    prsdnt_nm AS PRSDNT_NM,
                    telno AS TELNO,
                    addr AS ADDR,
                    prms_dt AS PRMS_DT,
                    source AS SOURCE
                FROM raw_businesses
                """
            ).fetchall()
            return [dict(r) for r in rows]

    def load_c006_items_for_licenses(self, lcns_nos: Iterable[str]) -> Dict[str, List[Dict[str, Any]]]:
        lcns_list = [str(x).strip() for x in lcns_nos if str(x).strip()]
        if not lcns_list:
            return {}

        grouped: Dict[str, List[Dict[str, Any]]] = {}
        chunk_size = 900  # SQLite parameter limit safety
        with self._lock, self._connect() as conn:
            for idx in range(0, len(lcns_list), chunk_size):
                chunk = lcns_list[idx : idx + chunk_size]
                placeholders = ",".join(["?"] * len(chunk))
                rows = conn.execute(
                    f"""
                    SELECT
                        lcns_no AS LCNS_NO,
                        prdlst_report_no AS PRDLST_REPORT_NO,
                        prdlst_nm AS PRDLST_NM,
                        bssh_nm AS BSSH_NM,
                        prdlst_dcnm AS PRDLST_DCNM,
                        rawmtrl_nm AS RAWMTRL_NM
                    FROM raw_c006_items
                    WHERE lcns_no IN ({placeholders})
                    """,
                    chunk,
                ).fetchall()
                for r in rows:
                    d = dict(r)
                    grouped.setdefault(d["LCNS_NO"], []).append(d)
        return grouped

    def get_cache_summary(self) -> Dict[str, Any]:
        with self._lock, self._connect() as conn:
            biz_cnt = conn.execute("SELECT COUNT(*) AS cnt FROM raw_businesses").fetchone()["cnt"]
            c006_cnt = conn.execute("SELECT COUNT(*) AS cnt FROM raw_c006_items").fetchone()["cnt"]
            checked_cnt = conn.execute(
                "SELECT COUNT(*) AS cnt FROM c006_checked_licenses"
            ).fetchone()["cnt"]
            tasks = conn.execute(
                "SELECT task_key, total_count, last_end_idx, is_completed, updated_at FROM collection_progress"
            ).fetchall()
            return {
                "business_count": int(biz_cnt),
                "c006_item_count": int(c006_cnt),
                "c006_checked_licenses": int(checked_cnt),
                "tasks": [dict(t) for t in tasks],
            }

    def clear_all(self) -> None:
        with self._lock, self._connect() as conn:
            conn.executescript(
                """
                DELETE FROM collection_progress;
                DELETE FROM raw_businesses;
                DELETE FROM c006_checked_licenses;
                DELETE FROM raw_c006_items;
                """
            )


class FoodSafetyAPIClient:
    """식품안전나라 OpenAPI 수집기 (requests.Session + tenacity 재시도 적용)"""

    def __init__(
        self,
        api_key: str,
        cache_store: SQLiteCacheStore,
        timeout: float = 25.0,
        max_workers: int = 4,
    ) -> None:
        raw_key = (api_key or "").strip()
        # 웹 페이지 표에서 복사 시 탭(\t)·공백 중복 또는 20자 키 2회 연속 붙여넣기(40자) 자동 정제
        token = raw_key.split()[0].strip() if raw_key else ""
        if len(token) == 40 and token[:20] == token[20:]:
            token = token[:20]
        self.api_key = token
        if not self.api_key:
            raise InvalidApiKeyError("식품안전나라 OpenAPI 인증키가 비어 있습니다.")
        self.cache = cache_store
        self.timeout = timeout
        self.max_workers = max(1, int(max_workers))
        self._thread_local = threading.local()

    def _get_session(self) -> requests.Session:
        session: Optional[requests.Session] = getattr(self._thread_local, "session", None)
        if session is None:
            session = requests.Session()
            adapter = HTTPAdapter(
                pool_connections=self.max_workers + 2,
                pool_maxsize=self.max_workers + 4,
            )
            session.mount("http://", adapter)
            session.mount("https://", adapter)
            session.headers.update(
                {
                    "Accept": "application/json",
                    "User-Agent": "FoodSafetyKorea-MeatCollector/1.0",
                }
            )
            self._thread_local.session = session
        return session

    def build_url(
        self,
        service_id: str,
        start_idx: int,
        end_idx: int,
        params: Optional[Dict[str, str]] = None,
    ) -> str:
        """
        규격 URL 생성:
        http://openapi.foodsafetykorea.go.kr/api/{API_KEY}/{SERVICE_ID}/json/{START_IDX}/{END_IDX}/{추가파라미터}
        """
        encoded_key = urllib.parse.quote(self.api_key, safe="")
        base = f"{BASE_API_URL}/{encoded_key}/{service_id}/json/{start_idx}/{end_idx}"
        if not params:
            return base

        parts = []
        for k, v in params.items():
            if v is None or str(v).strip() == "":
                continue
            ek = urllib.parse.quote(str(k).strip(), safe="")
            ev = urllib.parse.quote(str(v).strip(), safe="")
            parts.append(f"{ek}={ev}")
        if not parts:
            return base
        return f"{base}/{'&'.join(parts)}"

    @(
        retry(
            reraise=True,
            stop=stop_after_attempt(4),
            wait=wait_exponential(multiplier=1.2, min=1, max=10),
            retry=retry_if_exception_type(TransientAPIError),
        )
        if _HAS_TENACITY
        else _fallback_retry(TransientAPIError, max_attempts=4)
    )
    def fetch_page(
        self,
        service_id: str,
        start_idx: int,
        end_idx: int,
        params: Optional[Dict[str, str]] = None,
    ) -> Tuple[int, List[Dict[str, Any]], str, str]:
        """
        단일 페이지 호출 및 RESULT.CODE 검증.
        Returns: (total_count, rows, result_code, result_msg)
        """
        url = self.build_url(service_id, start_idx, end_idx, params)
        session = self._get_session()

        try:
            resp = session.get(url, timeout=self.timeout)
        except requests.RequestException as exc:
            raise TransientAPIError(f"[{service_id}] 네트워크 연결 오류: {exc}") from exc

        if resp.status_code >= 500:
            raise TransientAPIError(
                f"[{service_id}] 서버 일시 오류 (HTTP {resp.status_code})"
            )
        if resp.status_code != 200:
            raise FoodSafetyAPIError(
                f"[{service_id}] 비정상 HTTP 응답 코드: {resp.status_code}"
            )

        raw_text = resp.text.strip()
        if not raw_text:
            raise TransientAPIError(f"[{service_id}] 빈 응답 수신")

        # 식품안전나라 게이트웨이가 잘못된 인증키/초과 시 HTML <script>alert(...)</script>를 반환하는 현상 방어
        try:
            payload = resp.json()
        except ValueError as exc:
            if "인증키" in raw_text or "유효하지 않" in raw_text:
                raise InvalidApiKeyError(
                    "유효하지 않은 식품안전나라 인증키입니다. 인증키를 다시 확인해주세요."
                ) from exc
            if "초과" in raw_text or "호출건수" in raw_text:
                raise QuotaExceededError(
                    "식품안전나라 OpenAPI 일일/시간당 호출 한도를 초과했습니다. (캐시 자동 저장됨)"
                ) from exc
            raise TransientAPIError(
                f"[{service_id}] JSON 파싱 실패 (일시적 게이트웨이 응답 이상): {raw_text[:180]}"
            ) from exc

        # 서비스 루트 또는 최상위 RESULT 추출
        svc_body = payload.get(service_id) if isinstance(payload, dict) else None
        result_obj = None
        if isinstance(svc_body, dict) and isinstance(svc_body.get("RESULT"), dict):
            result_obj = svc_body["RESULT"]
        elif isinstance(payload, dict) and isinstance(payload.get("RESULT"), dict):
            result_obj = payload["RESULT"]

        code = str((result_obj or {}).get("CODE", "")).strip()
        msg = str((result_obj or {}).get("MSG", "")).strip()

        if code == "INFO-000":
            total_count = int(str((svc_body or {}).get("total_count", "0") or "0"))
            rows = (svc_body or {}).get("row", [])
            if not isinstance(rows, list):
                rows = []
            return total_count, rows, code, msg

        if code == "INFO-200":
            # 해당하는 데이터가 없음
            return 0, [], code, msg

        if code in ("INFO-100", "ERROR-301"):
            raise InvalidApiKeyError(f"인증키 오류 ({code}): {msg}")

        if code == "INFO-110":
            raise InvalidApiKeyError(
                f"[{service_id}] 서비스 활용신청 미완료 ({code}): 인증키는 정상 확인되었으나 "
                f"식품안전나라 OpenAPI 페이지에서 '{service_id}' 서비스의 [활용신청]을 추가로 클릭해야 합니다."
            )

        if code in ("INFO-300", "INFO-320"):
            raise QuotaExceededError(
                f"API 호출 제한 도달 ({code}: {msg}) - 현재까지 수집된 데이터는 로컬 SQLite 캐시에 안전하게 저장되었습니다. 한도 초기화 후 이어받기가 가능합니다."
            )

        if code.startswith("ERROR-5") or code in ("ERROR-001", ""):
            raise TransientAPIError(f"[{service_id}] 서버 응답 오류 ({code}): {msg}")

        raise FoodSafetyAPIError(f"[{service_id}] API 오류 ({code}): {msg}")

    def test_connection(self) -> Dict[str, Any]:
        """API 연결 및 주요 서비스(I2500, I1300, I1320, C006) 활용신청 여부 통합 점검"""
        service_checks: Dict[str, str] = {}
        sample_rows: List[Dict[str, Any]] = []
        active_code = ""
        active_msg = ""
        active_total = 0

        for svc, extra in (
            ("I2500", {"INDUTY_CD_NM": "식육가공업"}),
            ("I1300", None),
            ("I1320", None),
            ("C006", None),
        ):
            try:
                tot, rows, c_code, c_msg = self.fetch_page(svc, 1, 5, params=extra)
                service_checks[svc] = f"✅ 정상 ({c_code}, 총 {tot:,}건)"
                if not sample_rows and rows:
                    sample_rows = rows[:5]
                    active_code = c_code
                    active_msg = f"{svc}: {c_msg}"
                    active_total = tot
            except FoodSafetyAPIError as exc:
                service_checks[svc] = f"❌ {exc}"

        if not active_code:
            # 하나도 성공한 서비스가 없으면 대표 예외 발생
            self.fetch_page("I1320", 1, 5)

        return {
            "success": bool(active_code),
            "code": active_code,
            "message": active_msg,
            "total_count": active_total,
            "sample_rows": sample_rows,
            "service_checks": service_checks,
        }

    def collect_i2500_businesses(
        self,
        resume: bool = True,
        progress_cb: ProgressCallback = None,
    ) -> Dict[str, int]:
        """
        1) I2500 영업 중인 인허가 업소 정보 수집
           INDUTY_CD_NM = 식육가공업, 식육포장처리업, 축산물가공업 각각 1,000건 단위 전체 순회 후 합집합 저장
        """
        if self.api_key.lower() == "sample":
            raise InvalidApiKeyError(
                "'sample' 테스트 키는 1~5건 샘플 조회만 허용됩니다. 전체 수집을 위해 발급받은 정식 인증키를 입력해주세요."
            )

        stats: Dict[str, int] = {}
        num_ind = len(I2500_TARGET_INDUSTRIES)

        for idx, industry in enumerate(I2500_TARGET_INDUSTRIES):
            task_key = f"I2500:{industry}"
            cached_total, last_end, is_done = self.cache.get_task_progress(task_key)

            if resume and is_done:
                stats[industry] = cached_total
                if progress_cb:
                    progress_cb(
                        "I2500",
                        (idx + 1) / num_ind,
                        f"[I2500] '{industry}' 캐시 완료본 확인 ({cached_total:,}건)",
                    )
                continue

            start_idx = (last_end + 1) if (resume and last_end > 0) else 1
            total_count = cached_total

            while True:
                end_idx = start_idx + MAX_PAGE_SIZE - 1
                page_total, rows, code, _ = self.fetch_page(
                    service_id="I2500",
                    start_idx=start_idx,
                    end_idx=end_idx,
                    params={"INDUTY_CD_NM": industry},
                )

                if code == "INFO-200" or page_total == 0 or not rows:
                    self.cache.update_task_progress(task_key, total_count, total_count, True)
                    stats[industry] = total_count
                    break

                total_count = page_total
                self.cache.upsert_i2500_rows(rows, query_industry=industry)
                actual_end = min(end_idx, total_count)
                done_now = actual_end >= total_count
                self.cache.update_task_progress(task_key, total_count, actual_end, done_now)

                if progress_cb:
                    sub_ratio = (actual_end / total_count) if total_count > 0 else 1.0
                    overall = (idx + sub_ratio) / num_ind
                    progress_cb(
                        "I2500",
                        min(overall, 1.0),
                        f"[I2500] '{industry}' 수집 중: {actual_end:,} / {total_count:,}건",
                    )

                if done_now:
                    stats[industry] = total_count
                    break

                start_idx = end_idx + 1
                time.sleep(0.05)

        return stats

    def collect_i1300_supplement(
        self,
        resume: bool = True,
        progress_cb: ProgressCallback = None,
    ) -> Dict[str, int]:
        """
        1-보조) I1300 축산물 가공업 영업허가 현황 수집 및 LCNS_NO 기준 보완 병합
               (I2500 미신청 시에도 축산물가공업-식육가공업 전수 확보 가능)
        """
        if self.api_key.lower() == "sample":
            raise InvalidApiKeyError(
                "'sample' 테스트 키로는 전체 페이지네이션 수집을 실행할 수 없습니다."
            )

        task_key = "I1300:ALL"
        cached_total, last_end, is_done = self.cache.get_task_progress(task_key)
        inserted_total = 0
        supplemented_total = 0

        if resume and is_done:
            if progress_cb:
                progress_cb(
                    "I1300",
                    1.0,
                    f"[I1300] 축산물가공업 영업허가 현황 캐시 완료본 확인 ({cached_total:,}건)",
                )
            return {
                "total_count": cached_total,
                "inserted": 0,
                "supplemented": 0,
            }

        start_idx = (last_end + 1) if (resume and last_end > 0) else 1
        total_count = cached_total

        while True:
            end_idx = start_idx + MAX_PAGE_SIZE - 1
            page_total, rows, code, _ = self.fetch_page(
                service_id="I1300",
                start_idx=start_idx,
                end_idx=end_idx,
            )

            if code == "INFO-200" or page_total == 0 or not rows:
                self.cache.update_task_progress(task_key, total_count, total_count, True)
                break

            total_count = page_total
            ins_cnt, sup_cnt = self.cache.merge_i1320_rows(rows)
            inserted_total += ins_cnt
            supplemented_total += sup_cnt

            actual_end = min(end_idx, total_count)
            done_now = actual_end >= total_count
            self.cache.update_task_progress(task_key, total_count, actual_end, done_now)

            if progress_cb:
                ratio = (actual_end / total_count) if total_count > 0 else 1.0
                progress_cb(
                    "I1300",
                    min(ratio, 1.0),
                    f"[I1300] 축산물가공업 영업허가 현황 수집·병합 중: {actual_end:,} / {total_count:,}건 "
                    f"(신규 +{inserted_total:,}건)",
                )

            if done_now:
                break

            start_idx = end_idx + 1
            time.sleep(0.05)

        return {
            "total_count": total_count,
            "inserted": inserted_total,
            "supplemented": supplemented_total,
        }

    def collect_i1320_supplement(
        self,
        resume: bool = True,
        progress_cb: ProgressCallback = None,
    ) -> Dict[str, int]:
        """
        2) I1320 축산물 식육포장처리업 영업허가대장 수집 및 LCNS_NO 기준 누락분 보완 병합
        """
        if self.api_key.lower() == "sample":
            raise InvalidApiKeyError(
                "'sample' 테스트 키로는 전체 페이지네이션 수집을 실행할 수 없습니다."
            )

        task_key = "I1320:ALL"
        cached_total, last_end, is_done = self.cache.get_task_progress(task_key)
        inserted_total = 0
        supplemented_total = 0

        if resume and is_done:
            if progress_cb:
                progress_cb(
                    "I1320",
                    1.0,
                    f"[I1320] 식육포장처리업 허가대장 캐시 완료본 확인 ({cached_total:,}건)",
                )
            return {
                "total_count": cached_total,
                "inserted": 0,
                "supplemented": 0,
            }

        start_idx = (last_end + 1) if (resume and last_end > 0) else 1
        total_count = cached_total

        while True:
            end_idx = start_idx + MAX_PAGE_SIZE - 1
            page_total, rows, code, _ = self.fetch_page(
                service_id="I1320",
                start_idx=start_idx,
                end_idx=end_idx,
            )

            if code == "INFO-200" or page_total == 0 or not rows:
                self.cache.update_task_progress(task_key, total_count, total_count, True)
                break

            total_count = page_total
            ins_cnt, sup_cnt = self.cache.merge_i1320_rows(rows)
            inserted_total += ins_cnt
            supplemented_total += sup_cnt

            actual_end = min(end_idx, total_count)
            done_now = actual_end >= total_count
            self.cache.update_task_progress(task_key, total_count, actual_end, done_now)

            if progress_cb:
                ratio = (actual_end / total_count) if total_count > 0 else 1.0
                progress_cb(
                    "I1320",
                    min(ratio, 1.0),
                    f"[I1320] 식육포장처리업 허가대장 보완 병합 중: {actual_end:,} / {total_count:,}건 "
                    f"(신규 보완 +{inserted_total:,}건)",
                )

            if done_now:
                break

            start_idx = end_idx + 1
            time.sleep(0.05)

        return {
            "total_count": total_count,
            "inserted": inserted_total,
            "supplemented": supplemented_total,
        }

    def _fetch_c006_single_license(self, lcns_no: str) -> Tuple[str, int]:
        """단일 인허가번호(LCNS_NO)에 대한 C006 품목제조보고 전 페이지 조회 및 저장"""
        start_idx = 1
        saved_count = 0
        while True:
            end_idx = start_idx + MAX_PAGE_SIZE - 1
            total_count, rows, code, _ = self.fetch_page(
                service_id="C006",
                start_idx=start_idx,
                end_idx=end_idx,
                params={"LCNS_NO": lcns_no},
            )
            if code == "INFO-200" or total_count == 0 or not rows:
                break

            saved_count += self.cache.upsert_c006_rows(rows)
            if end_idx >= total_count:
                break
            start_idx = end_idx + 1

        self.cache.mark_c006_license_checked(lcns_no, saved_count)
        return lcns_no, saved_count

    def collect_c006_by_licenses(
        self,
        target_businesses: Sequence[Dict[str, Any]],
        include_packaging_shops: bool = False,
        resume: bool = True,
        progress_cb: ProgressCallback = None,
    ) -> Dict[str, int]:
        """
        3-A) 1차 지역 필터링된 대상 업소들의 LCNS_NO 리스트 기준 C006 개별 호출 매핑.
             식육가공업/축산물가공업 업소를 최우선으로 조회하여 일일 호출 한도 효율을 극대화한다.
        """
        priority_licenses: List[str] = []
        secondary_licenses: List[str] = []

        seen: Set[str] = set()
        for biz in target_businesses:
            lcns = str(biz.get("LCNS_NO", "") or "").strip()
            if not lcns or lcns in seen:
                continue
            seen.add(lcns)
            induty = str(biz.get("INDUTY_CD_NM", "") or "")
            # 식육가공업·축산물가공업은 품목제조보고(C006)가 양념육 분류의 핵심이므로 최우선 배치
            if "식육가공업" in induty or "축산물가공업" in induty:
                priority_licenses.append(lcns)
            elif include_packaging_shops:
                secondary_licenses.append(lcns)

        ordered_licenses = priority_licenses + secondary_licenses
        already_checked = self.cache.get_checked_c006_licenses() if resume else set()
        pending_licenses = [lc for lc in ordered_licenses if lc not in already_checked]

        total_targets = len(ordered_licenses)
        skipped_cached = total_targets - len(pending_licenses)

        if total_targets == 0 or not pending_licenses:
            if progress_cb:
                progress_cb(
                    "C006",
                    1.0,
                    f"[C006] 대상 인허가번호 {total_targets:,}건 모두 캐시 매핑 완료",
                )
            return {
                "total_targets": total_targets,
                "queried_now": 0,
                "cached_skipped": skipped_cached,
                "matched_rows": 0,
            }

        completed_now = 0
        matched_rows = 0

        # 동시성 제어(ThreadPoolExecutor)로 속도 확보 및 과도한 트래픽 방지
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            future_to_lcns = {
                executor.submit(self._fetch_c006_single_license, lc): lc
                for lc in pending_licenses
            }
            try:
                for future in as_completed(future_to_lcns):
                    _, row_cnt = future.result()
                    completed_now += 1
                    matched_rows += row_cnt
                    done_overall = skipped_cached + completed_now
                    if progress_cb and (
                        completed_now % 5 == 0 or done_overall == total_targets
                    ):
                        progress_cb(
                            "C006",
                            min(done_overall / max(total_targets, 1), 1.0),
                            f"[C006] 업소별 품목제조보고 조회 중: {done_overall:,} / {total_targets:,}개소 "
                            f"(누적 품목행 {matched_rows:,}건)",
                        )
            except Exception:
                # QuotaExceededError 등 발생 시 대기 중인 작업 취소 후 즉시 상위로 전달 (완료분은 이미 SQLite 커밋됨)
                for f in future_to_lcns:
                    f.cancel()
                raise

        return {
            "total_targets": total_targets,
            "queried_now": completed_now,
            "cached_skipped": skipped_cached,
            "matched_rows": matched_rows,
        }

    def collect_c006_full_pagination(
        self,
        allowed_licenses: Optional[Set[str]] = None,
        resume: bool = True,
        progress_cb: ProgressCallback = None,
    ) -> Dict[str, int]:
        """
        3-B) C006 전체 데이터를 1,000건 단위로 페이지네이션 순회하여 대상 업소(LCNS_NO) 품목 일괄 매핑.
             대상 업소가 수천 개 이상일 때 약 570~580회 호출만으로 식육포장처리업 포함 전 업종을 매핑 가능.
        """
        task_key = "C006:FULL_SCAN"
        cached_total, last_end, is_done = self.cache.get_task_progress(task_key)
        saved_total = 0

        if resume and is_done:
            if progress_cb:
                progress_cb(
                    "C006",
                    1.0,
                    f"[C006] 전체 품목제조보고 페이지네이션 캐시 완료본 확인 ({cached_total:,}건)",
                )
            return {"total_count": cached_total, "saved_rows": 0}

        start_idx = (last_end + 1) if (resume and last_end > 0) else 1
        total_count = cached_total

        while True:
            end_idx = start_idx + MAX_PAGE_SIZE - 1
            page_total, rows, code, _ = self.fetch_page(
                service_id="C006",
                start_idx=start_idx,
                end_idx=end_idx,
            )

            if code == "INFO-200" or page_total == 0 or not rows:
                self.cache.update_task_progress(task_key, total_count, total_count, True)
                break

            total_count = page_total
            saved_total += self.cache.upsert_c006_rows(
                rows, allowed_licenses=allowed_licenses
            )

            actual_end = min(end_idx, total_count)
            done_now = actual_end >= total_count
            self.cache.update_task_progress(task_key, total_count, actual_end, done_now)

            if progress_cb:
                ratio = (actual_end / total_count) if total_count > 0 else 1.0
                progress_cb(
                    "C006",
                    min(ratio, 1.0),
                    f"[C006] 전체 품목제조보고 스캔·매핑 중: {actual_end:,} / {total_count:,}건 "
                    f"(대상 업소 매칭 +{saved_total:,}건)",
                )

            if done_now:
                break

            start_idx = end_idx + 1
            time.sleep(0.04)

        return {"total_count": total_count, "saved_rows": saved_total}
