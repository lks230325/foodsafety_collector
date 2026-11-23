#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GitHub Pages 정적 웹 배포 및 AES-256-GCM 보안 암호화 빌드 스크립트 (build_gh_pages.py)
- 로컬 실행 시: SQLite 캐시(`data/foodsafety_cache.sqlite`)에서 정제된 전체 가공장 JSON 및
  5개 시트 서식 완비 엑셀을 생성한 뒤, `.env`의 `DASHBOARD_PASSWORD`로 PBKDF2-HMAC-SHA256 + AES-256-GCM
  암호화하여 `static_data/*.enc.json`만 저장 (평문 JSON/엑셀은 디스크·Git에 남기지 않음)
- GitHub Actions 실행 시(`--from-static-only`): 커밋된 암호화 파일(`*.enc.json`)과 `web_template/index.html`,
  검색엔진 차단 `robots.txt`를 결합하여 `./dist` 빌드
"""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
import secrets
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).resolve().parent
ENV_PATH = BASE_DIR / ".env"
CACHE_DB_PATH = BASE_DIR / "data" / "foodsafety_cache.sqlite"
STATIC_DATA_DIR = BASE_DIR / "static_data"
TEMPLATE_PATH = BASE_DIR / "web_template" / "index.html"
DIST_DIR = BASE_DIR / "dist"

ENC_JSON_FILENAME = "processors_latest.enc.json"
ENC_EXCEL_FILENAME = "meat_processors_5sheets.enc.json"
PBKDF2_ITERATIONS = 600_000


def encrypt_bytes_aes256_gcm(plaintext: bytes, password: str) -> Dict[str, object]:
    """PBKDF2-HMAC-SHA256(600,000회) + AES-256-GCM으로 바이트 데이터를 암호화한다 (Web Crypto API 100% 호환)."""
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM

    salt = secrets.token_bytes(16)
    iv = secrets.token_bytes(12)
    key = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        PBKDF2_ITERATIONS,
        dklen=32,
    )
    ciphertext_and_tag = AESGCM(key).encrypt(iv, plaintext, None)
    return {
        "alg": "AES-256-GCM",
        "kdf": "PBKDF2-SHA256",
        "iter": PBKDF2_ITERATIONS,
        "salt": base64.b64encode(salt).decode("ascii"),
        "iv": base64.b64encode(iv).decode("ascii"),
        "ct": base64.b64encode(ciphertext_and_tag).decode("ascii"),
    }


def export_encrypted_static_data(password: str) -> int:
    """로컬 SQLite 캐시에서 최신 정제 데이터프레임 및 5개 시트 엑셀을 생성 후 AES-256-GCM 암호화 저장한다."""
    if not CACHE_DB_PATH.exists():
        raise FileNotFoundError(f"로컬 SQLite 캐시 파일이 없습니다: {CACHE_DB_PATH}")
    if not password:
        raise ValueError(
            "암호화 빌드를 위해 .env 또는 환경변수에 DASHBOARD_PASSWORD가 설정되어 있어야 합니다."
        )

    from api_client import SQLiteCacheStore
    from excel_exporter import generate_multisheet_excel_bytes
    from processor import build_processed_dataframe, filter_businesses_by_region

    STATIC_DATA_DIR.mkdir(parents=True, exist_ok=True)

    # 기존 평문 파일이 남아 있다면 즉시 삭제하여 Git/배포 유출 원천 방지
    for legacy_plain in ("processors_latest.json", "meat_processors_5sheets.xlsx", "충청전북_양념육_세절육_가공장현황.xlsx"):
        lp = STATIC_DATA_DIR / legacy_plain
        if lp.exists():
            lp.unlink()

    cache = SQLiteCacheStore(CACHE_DB_PATH)
    all_biz = cache.load_all_businesses()
    reg_biz = filter_businesses_by_region(all_biz, ["충남", "충북", "대전", "전북"])
    c006_map = cache.load_c006_items_for_licenses(
        [str(b.get("LCNS_NO", "") or "").strip() for b in reg_biz]
    )
    df = build_processed_dataframe(reg_biz, c006_map, merge_same_facility=True)

    payload_bytes = json.dumps(
        {
            "meta": {
                "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "total_count": int(len(df)),
                "regions": ["충청남도", "충청북도", "대전광역시", "전북특별자치도(전라북도)"],
            },
            "rows": df.to_dict(orient="records"),
        },
        ensure_ascii=False,
        separators=(",", ":"),
    ).encode("utf-8")

    excel_bytes = generate_multisheet_excel_bytes(df)

    enc_json_bundle = encrypt_bytes_aes256_gcm(payload_bytes, password)
    enc_excel_bundle = encrypt_bytes_aes256_gcm(excel_bytes, password)

    (STATIC_DATA_DIR / ENC_JSON_FILENAME).write_text(
        json.dumps(enc_json_bundle, separators=(",", ":")),
        encoding="utf-8",
    )
    (STATIC_DATA_DIR / ENC_EXCEL_FILENAME).write_text(
        json.dumps(enc_excel_bundle, separators=(",", ":")),
        encoding="utf-8",
    )
    return int(len(df))


def build_dist(from_static_only: bool = False) -> Path:
    """암호화된 static_data/*.enc.json 및 보안 웹 템플릿을 ./dist 정적 웹사이트로 패키징한다."""
    if not from_static_only and CACHE_DB_PATH.exists():
        from dotenv import load_dotenv

        load_dotenv(ENV_PATH, encoding="utf-8-sig", override=True)
        pw = os.getenv("DASHBOARD_PASSWORD", "").strip()
        row_cnt = export_encrypted_static_data(password=pw)
        print(f"[Encrypt] SQLite -> AES-256-GCM 암호화 스냅샷 생성 완료 ({row_cnt:,}개 가공장)")

    src_enc_json = STATIC_DATA_DIR / ENC_JSON_FILENAME
    src_enc_xlsx = STATIC_DATA_DIR / ENC_EXCEL_FILENAME
    if not src_enc_json.exists() or not src_enc_xlsx.exists():
        raise FileNotFoundError(
            f"암호화된 정적 데이터 파일이 없습니다: {src_enc_json} 또는 {src_enc_xlsx}"
        )
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"웹 템플릿 파일이 없습니다: {TEMPLATE_PATH}")

    if DIST_DIR.exists():
        shutil.rmtree(DIST_DIR)
    dist_data_dir = DIST_DIR / "data"
    dist_data_dir.mkdir(parents=True, exist_ok=True)

    shutil.copy2(TEMPLATE_PATH, DIST_DIR / "index.html")
    shutil.copy2(src_enc_json, dist_data_dir / ENC_JSON_FILENAME)
    shutil.copy2(src_enc_xlsx, dist_data_dir / ENC_EXCEL_FILENAME)

    # 검색엔진(Google, Naver, Bing 등) 크롤링·색인 원천 차단
    (DIST_DIR / "robots.txt").write_text(
        "User-agent: *\nDisallow: /\n",
        encoding="utf-8",
    )
    (DIST_DIR / ".nojekyll").write_text("", encoding="utf-8")

    print(f"[Build] 보안 정적 사이트 빌드 성공 -> {DIST_DIR}")
    return DIST_DIR


def main() -> int:
    parser = argparse.ArgumentParser(description="Build encrypted static site for GitHub Pages")
    parser.add_argument(
        "--from-static-only",
        action="store_true",
        help="SQLite 재생성 없이 커밋된 암호화 스냅샷(*.enc.json)만으로 dist/ 빌드",
    )
    args = parser.parse_args()
    build_dist(from_static_only=args.from_static_only)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
