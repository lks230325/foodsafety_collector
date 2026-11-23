# -*- coding: utf-8 -*-
"""
CLI 원클릭 자동 수집 및 멀티시트 엑셀(.xlsx) 생성 스크립트
.env 파일에 FOOD_API_KEY가 저장되어 있으면 브라우저 조작 없이 전수 수집부터 엑셀 파일 저장까지 완료합니다.
"""

from __future__ import annotations

import os
import sys
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv

from api_client import (
    FoodSafetyAPIClient,
    InvalidApiKeyError,
    QuotaExceededError,
    SQLiteCacheStore,
)
from excel_exporter import generate_multisheet_excel_bytes
from processor import build_processed_dataframe, filter_businesses_by_region

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env", encoding="utf-8-sig", override=True)
load_dotenv(BASE_DIR.parent / ".env", encoding="utf-8-sig", override=False)

CACHE_DB_PATH = BASE_DIR / "data" / "foodsafety_cache.sqlite"


def main() -> int:
    api_key = os.getenv("FOOD_API_KEY", "").strip()
    if not api_key or api_key.lower() == "sample" or "your_" in api_key:
        print(
            "[안내] .env 파일에 정식 발급된 FOOD_API_KEY가 설정되어 있지 않습니다.\n"
            f"설정 파일 경로: {BASE_DIR / '.env'}",
            file=sys.stderr,
        )
        return 1

    cache = SQLiteCacheStore(CACHE_DB_PATH)
    client = FoodSafetyAPIClient(api_key=api_key, cache_store=cache, timeout=25.0, max_workers=4)

    def log_progress(stage: str, ratio: float, msg: str) -> None:
        print(f"[{stage} {ratio * 100:5.1f}%] {msg}", flush=True)

    try:
        print("=== [1/3단계] 영업 인허가 업소(I2500 · I1300 · I1320) 수집 시작 ===")
        try:
            client.collect_i2500_businesses(resume=True, progress_cb=log_progress)
        except InvalidApiKeyError as exc:
            if "INFO-110" in str(exc):
                print("[안내] I2500 미신청 상태 -> 신청 완료된 I1300(축산물가공업)+I1320(식육포장처리업)으로 대체 수집합니다.")
            else:
                raise

        try:
            client.collect_i1300_supplement(resume=True, progress_cb=log_progress)
        except InvalidApiKeyError as exc:
            if "INFO-110" not in str(exc):
                raise

        client.collect_i1320_supplement(resume=True, progress_cb=log_progress)

        print("=== [2/3단계] 충남·충북·대전·전북 대상 지역 필터링 ===")
        all_biz = cache.load_all_businesses()
        reg_biz = filter_businesses_by_region(all_biz, ["충남", "충북", "대전", "전북"])
        target_lcns = {
            str(b.get("LCNS_NO", "") or "").strip()
            for b in reg_biz
            if str(b.get("LCNS_NO", "") or "").strip()
        }
        print(f"대상 지역 업소 추출 완료: {len(reg_biz):,}개소 (인허가번호 {len(target_lcns):,}건)")

        print("=== [3/3단계] C006 축산물 품목제조보고 매핑 ===")
        try:
            client.collect_c006_full_pagination(
                allowed_licenses=target_lcns,
                resume=True,
                progress_cb=log_progress,
            )
        except InvalidApiKeyError as exc:
            if "INFO-110" in str(exc):
                print("[안내] C006(축산물품목제조보고) 활용신청 전 -> I1300(식육가공업)+I1320(식육포장처리업) 인허가 기준으로 분류합니다.")
            else:
                raise
    except QuotaExceededError as exc:
        print(f"\n[경고] {exc}\n현재까지 수집된 캐시 데이터로 엑셀 파일을 우선 생성합니다.")
    except InvalidApiKeyError as exc:
        print(f"\n[오류] {exc}", file=sys.stderr)
        return 2

    all_biz = cache.load_all_businesses()
    reg_biz = filter_businesses_by_region(all_biz, ["충남", "충북", "대전", "전북"])
    c006_map = cache.load_c006_items_for_licenses(
        [str(b.get("LCNS_NO", "") or "").strip() for b in reg_biz]
    )
    df = build_processed_dataframe(reg_biz, c006_map, merge_same_facility=True)

    out_path = BASE_DIR / f"충청전북_양념육_세절육_가공장현황_{datetime.now().strftime('%Y%m%d')}.xlsx"
    out_path.write_bytes(generate_multisheet_excel_bytes(df))
    print(f"\n[완료] 총 {len(df):,}개 가공장 정제 및 엑셀 생성 완료: {out_path}")
    print("지역별 현황:", df["시/도"].value_counts().to_dict())
    print("가공구분별 현황:", df["가공구분"].value_counts().to_dict())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
