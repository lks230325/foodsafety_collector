# -*- coding: utf-8 -*-
"""
식품안전나라(식품의약품안전처) OpenAPI 연동
충남·충북·대전·전북 '양념육' 및 '세절육(식육포장처리)' 가공장 통합 수집·정제·엑셀 다운로드 웹 애플리케이션
"""

from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path
from typing import Dict, List

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from api_client import (
    FoodSafetyAPIClient,
    FoodSafetyAPIError,
    InvalidApiKeyError,
    QuotaExceededError,
    SQLiteCacheStore,
)
from excel_exporter import generate_multisheet_excel_bytes
from processor import (
    CATEGORY_BOTH,
    CATEGORY_SEASONED,
    CATEGORY_SLICED,
    OUTPUT_COLUMNS,
    build_processed_dataframe,
    filter_businesses_by_region,
)

BASE_DIR = Path(__file__).resolve().parent
ENV_FILE_PATH = BASE_DIR / ".env"
# 현재 폴더(.env) 및 상위 워크스페이스 폴더(.env) 모두 자동 탐색 지원 (BOM 방어 포함)
load_dotenv(ENV_FILE_PATH, encoding="utf-8-sig", override=True)
load_dotenv(BASE_DIR.parent / ".env", encoding="utf-8-sig", override=False)

CACHE_DB_PATH = BASE_DIR / "data" / "foodsafety_cache.sqlite"
PARQUET_SNAPSHOT_PATH = BASE_DIR / "data" / "latest_processed.parquet"


def save_api_key_to_env(api_key: str) -> None:
    """유효한 인증키를 foodsafety_collector/.env 파일에 안전하게 영구 기록한다."""
    clean_key = (api_key or "").strip()
    if not clean_key or clean_key.lower() == "sample":
        return
    ENV_FILE_PATH.write_text(
        "# 식품안전나라(식품의약품안전처) OpenAPI 인증키\n"
        f"FOOD_API_KEY={clean_key}\n",
        encoding="utf-8",
    )
    os.environ["FOOD_API_KEY"] = clean_key

st.set_page_config(
    page_title="충청·전북 양념육·세절육 가공장 수집기 (식품안전나라 OpenAPI)",
    page_icon="🥩",
    layout="wide",
)


@st.cache_resource
def get_cache_store() -> SQLiteCacheStore:
    return SQLiteCacheStore(CACHE_DB_PATH)


def rebuild_dataframe_from_cache(
    cache_store: SQLiteCacheStore,
    selected_regions: List[str],
    merge_same_facility: bool,
) -> pd.DataFrame:
    """SQLite 로컬 캐시에 적재된 원본 업소·품목 데이터를 바탕으로 대상 지역 DataFrame을 즉시 재구성한다."""
    all_businesses = cache_store.load_all_businesses()
    if not all_businesses or not selected_regions:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    regional_biz = filter_businesses_by_region(all_businesses, selected_regions)
    if not regional_biz:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    lcns_list = [str(b.get("LCNS_NO", "") or "").strip() for b in regional_biz]
    c006_map = cache_store.load_c006_items_for_licenses(lcns_list)
    df = build_processed_dataframe(
        regional_businesses=regional_biz,
        c006_by_lcns=c006_map,
        merge_same_facility=merge_same_facility,
    )
    return df


def render_sidebar(cache_store: SQLiteCacheStore) -> Dict[str, object]:
    """사이드바 설정 패널 렌더링 및 사용자 입력값 반환"""
    st.sidebar.title("⚙️ 수집 설정 및 인증")
    st.sidebar.caption("식품의약품안전처(식품안전나라) 실시간 OpenAPI 연동")

    env_api_key = os.getenv("FOOD_API_KEY", "").strip()
    api_key = st.sidebar.text_input(
        "식품안전나라 OpenAPI 인증키",
        value=env_api_key,
        type="password",
        placeholder="발급받은 API 인증키 입력 (.env 자동 로드)",
        help=(
            ".env 파일에 FOOD_API_KEY가 설정되어 있으면 자동으로 불러옵니다. "
            "연결 테스트만 먼저 확인하려면 'sample'을 입력해 5건 조회를 테스트할 수 있습니다."
        ),
    ).strip()

    if env_api_key and api_key == env_api_key:
        st.sidebar.caption("✅ `.env` 파일에서 `FOOD_API_KEY`가 자동 로드되었습니다.")

    btn_k1, btn_k2 = st.sidebar.columns(2)
    save_key_clicked = btn_k1.button(
        "💾 키 .env 저장",
        width="stretch",
        help="입력창의 인증키를 foodsafety_collector/.env 파일에 즉시 저장합니다.",
    )
    test_clicked = btn_k2.button(
        "🔌 샘플 5건 테스트",
        width="stretch",
    )
    if save_key_clicked:
        if api_key and api_key.lower() != "sample":
            save_api_key_to_env(api_key)
            st.sidebar.success("`.env` 파일에 인증키가 저장되었습니다!")
        else:
            st.sidebar.warning("정식 발급 인증키를 입력한 뒤 저장해 주세요.")

    st.sidebar.divider()
    st.sidebar.subheader("📍 대상 지역 선택")
    col_r1, col_r2 = st.sidebar.columns(2)
    chk_cn = col_r1.checkbox("충청남도 (충남)", value=True)
    chk_cb = col_r2.checkbox("충청북도 (충북)", value=True)
    chk_dj = col_r1.checkbox("대전광역시 (대전)", value=True)
    chk_jb = col_r2.checkbox("전북특별자치도 (전북)", value=True)

    selected_regions: List[str] = []
    if chk_cn:
        selected_regions.append("충남")
    if chk_cb:
        selected_regions.append("충북")
    if chk_dj:
        selected_regions.append("대전")
    if chk_jb:
        selected_regions.append("전북")

    st.sidebar.divider()
    st.sidebar.subheader("🔍 품목 매핑(C006) 및 병합 옵션")

    c006_mode = st.sidebar.radio(
        "C006(축산물 품목제조보고) 수집 방식",
        options=[
            "priority_lcns",
            "full_pagination",
            "all_lcns",
        ],
        format_func=lambda x: {
            "priority_lcns": "식육·축산물가공업 LCNS_NO 조회 (권장·호출 절약)",
            "full_pagination": "C006 전체 페이지네이션 후 일괄 매핑 (~575회 호출)",
            "all_lcns": "대상 지역 전 업소 LCNS_NO 개별 조회 (포장처리업 포함)",
        }[x],
        help=(
            "• 식육포장처리업은 품목제조보고 신고 면제 대상이 대부분이므로 "
            "'식육·축산물가공업 LCNS_NO 조회'를 선택하면 일일 API 호출 한도를 크게 아낄 수 있습니다.\n"
            "• 식육포장처리업 중 자발적 품목보고 업소까지 적은 호출수로 전수 매핑하려면 "
            "'C006 전체 페이지네이션 후 일괄 매핑'을 선택하세요."
        ),
    )

    merge_same_facility = st.sidebar.checkbox(
        "동일 사업장(업체명+주소 일치) 복수 인허가 통합",
        value=True,
        help=(
            "한 공장에서 '식육가공업'과 '식육포장처리업' 인허가를 각각 보유한 경우 "
            "하나의 행으로 병합하여 '양념육 + 세절·포장육 겸업'으로 정확히 분류합니다."
        ),
    )

    st.sidebar.divider()
    st.sidebar.subheader("💾 로컬 SQLite 캐시 (이어받기)")
    cache_info = cache_store.get_cache_summary()
    completed_tasks = sum(1 for t in cache_info["tasks"] if t["is_completed"])
    total_tasks = len(cache_info["tasks"])

    st.sidebar.caption(
        f"• 캐시 파일: `{CACHE_DB_PATH.name}`\n"
        f"• 저장된 인허가 업소: **{cache_info['business_count']:,}건**\n"
        f"• 저장된 품목보고(C006): **{cache_info['c006_item_count']:,}건** "
        f"(개별조회 완료 `{cache_info['c006_checked_licenses']:,}`개소)\n"
        f"• 페이지네이션 완료 작업: **{completed_tasks} / {total_tasks}개**"
    )

    run_clicked = st.sidebar.button(
        "🚀 데이터 수집 실행 (이어받기 지원)",
        type="primary",
        width="stretch",
    )

    col_c1, col_c2 = st.sidebar.columns(2)
    refresh_cache_clicked = col_c1.button(
        "🔄 캐시 재집계",
        width="stretch",
        help="API 추가 호출 없이 로컬 SQLite 캐시 데이터로 현재 지역/옵션 기준 표를 다시 생성합니다.",
    )
    clear_cache_clicked = col_c2.button(
        "🗑️ 캐시 초기화",
        width="stretch",
        help="로컬 SQLite 캐시를 비우고 처음부터 새로 수집할 때 사용합니다.",
    )

    return {
        "api_key": api_key,
        "test_clicked": test_clicked,
        "selected_regions": selected_regions,
        "c006_mode": c006_mode,
        "merge_same_facility": merge_same_facility,
        "run_clicked": run_clicked,
        "refresh_cache_clicked": refresh_cache_clicked,
        "clear_cache_clicked": clear_cache_clicked,
    }


def handle_connection_test(api_key: str, cache_store: SQLiteCacheStore) -> None:
    """사이드바 API 연결 테스트(샘플 5건 조회) 결과 표시"""
    if not api_key:
        st.error("인증키를 입력한 뒤 API 연결 테스트를 실행해주세요. (샘플 확인 시 `sample` 입력 가능)")
        return

    with st.spinner("식품안전나라 OpenAPI(`I2500` 식육가공업 샘플 5건) 연결 확인 중..."):
        try:
            client = FoodSafetyAPIClient(api_key=api_key, cache_store=cache_store)
            res = client.test_connection()
            if res["success"]:
                save_api_key_to_env(api_key)
            st.success(
                f"✅ API 연결 성공! 응답코드: `{res['code']}` ({res['message']}) | "
                f"대표 건수: **{res['total_count']:,}건**"
            )
            if res.get("service_checks"):
                for svc_id, svc_msg in res["service_checks"].items():
                    st.caption(f"• **`{svc_id}`**: {svc_msg}")
            if res["sample_rows"]:
                sample_df = pd.DataFrame(res["sample_rows"])
                cols = [
                    c
                    for c in [
                        "LCNS_NO",
                        "BSSH_NM",
                        "INDUTY_CD_NM",
                        "INDUTY_NM",
                        "PRSDNT_NM",
                        "TELNO",
                        "ADDR",
                        "LOCP_ADDR",
                        "PRMS_DT",
                    ]
                    if c in sample_df.columns
                ]
                with st.expander("📋 실시간 수신 샘플 5건 확인", expanded=True):
                    st.dataframe(sample_df[cols], width="stretch", hide_index=True)
        except FoodSafetyAPIError as exc:
            st.error(f"❌ API 연결 테스트 실패: {exc}")
        except Exception as exc:
            st.error(f"❌ 예기치 못한 오류 발생: {exc}")


def run_collection_pipeline(
    api_key: str,
    selected_regions: List[str],
    c006_mode: str,
    merge_same_facility: bool,
    cache_store: SQLiteCacheStore,
) -> None:
    """3단계 수집·필터링·품목매핑 파이프라인 실행 및 실시간 진행률 표시"""
    if not api_key:
        st.error("식품안전나라 OpenAPI 인증키를 입력해주세요.")
        return
    if not selected_regions:
        st.warning("최소 1개 이상의 대상 지역(충남, 충북, 대전, 전북)을 선택해주세요.")
        return

    progress_bar = st.progress(0.0, text="수집 파이프라인 준비 중...")
    status_box = st.status("🚀 식품안전나라 OpenAPI 데이터 수집 및 정제 진행 중...", expanded=True)

    try:
        client = FoodSafetyAPIClient(
            api_key=api_key,
            cache_store=cache_store,
            timeout=25.0,
            max_workers=4,
        )
        save_api_key_to_env(api_key)

        # [1단계] I2500(3개 업종 합집합) + I1300(축산물가공업 허가현황) + I1320(식육포장처리업 허가대장)
        status_box.write("1️⃣ **[1/3단계] 영업 인허가 업소 수집 (`I2500` · `I1300` · `I1320` 통합 병합)** 시작...")

        def cb_i2500(_stage: str, ratio: float, msg: str) -> None:
            overall = 0.02 + (ratio * 0.18)
            progress_bar.progress(min(overall, 0.20), text=f"[1/3단계 업소 수집] {msg}")

        try:
            i2500_stats = client.collect_i2500_businesses(resume=True, progress_cb=cb_i2500)
            status_box.write(
                "   - `I2500` 업종별 확인 완료: "
                + ", ".join(f"{k}({v:,}건)" for k, v in i2500_stats.items())
            )
        except InvalidApiKeyError as exc:
            if "INFO-110" in str(exc):
                status_box.write("   - ℹ️ `I2500` 미신청 상태(`INFO-110`) → 신청 완료된 `I1300`+`I1320` 허가현황으로 대체 수집합니다.")
            else:
                raise

        def cb_i1300(_stage: str, ratio: float, msg: str) -> None:
            overall = 0.20 + (ratio * 0.12)
            progress_bar.progress(min(overall, 0.32), text=f"[1/3단계 축산물가공업 수집] {msg}")

        try:
            i1300_stats = client.collect_i1300_supplement(resume=True, progress_cb=cb_i1300)
            status_box.write(
                f"   - `I1300` 축산물가공업 허가현황 병합 완료: 총 {i1300_stats['total_count']:,}건 "
                f"(신규 +{i1300_stats['inserted']:,}건)"
            )
        except InvalidApiKeyError as exc:
            if "INFO-110" not in str(exc):
                raise

        def cb_i1320(_stage: str, ratio: float, msg: str) -> None:
            overall = 0.32 + (ratio * 0.13)
            progress_bar.progress(min(overall, 0.45), text=f"[1/3단계 식육포장처리업 병합] {msg}")

        i1320_stats = client.collect_i1320_supplement(resume=True, progress_cb=cb_i1320)
        status_box.write(
            f"   - `I1320` 식육포장처리업 허가대장 보완 완료: 총 {i1320_stats['total_count']:,}건 스캔 "
            f"(누락 업소 신규 추가 +{i1320_stats['inserted']:,}건, 상세정보 보완 +{i1320_stats['supplemented']:,}건)"
        )

        # [2단계] 대상 지역(충남·충북·대전·전북) 필터링
        progress_bar.progress(0.48, text="[2/3단계 지역 필터링] 충남·충북·대전·전북 소재지 추출 중...")
        status_box.write("2️⃣ **[2/3단계] 대상 지역(`ADDR` 기준) 필터링 및 전북 명칭 통합** 진행...")

        all_businesses = cache_store.load_all_businesses()
        regional_businesses = filter_businesses_by_region(all_businesses, selected_regions)
        target_lcns_set = {
            str(b.get("LCNS_NO", "") or "").strip()
            for b in regional_businesses
            if str(b.get("LCNS_NO", "") or "").strip()
        }

        status_box.write(
            f"   - 전국 수집 업소 {len(all_businesses):,}개소 중 선택 지역 해당 업소: "
            f"**{len(regional_businesses):,}개소** (인허가번호 `{len(target_lcns_set):,}`건)"
        )
        progress_bar.progress(0.52, text=f"[2/3단계 완료] 대상 지역 업소 {len(regional_businesses):,}개소 추출 완료")

        # [3단계] 축산물 품목제조보고(C006) 매핑
        status_box.write("3️⃣ **[3/3단계] 축산물 품목제조보고(`C006`) 수집·매핑 및 가공구분 분류** 시작...")

        def cb_c006(_stage: str, ratio: float, msg: str) -> None:
            overall = 0.52 + (ratio * 0.43)  # 52% ~ 95%
            progress_bar.progress(min(overall, 0.95), text=f"[3/3단계 품목 매핑] {msg}")

        try:
            if c006_mode == "full_pagination":
                c006_stats = client.collect_c006_full_pagination(
                    allowed_licenses=target_lcns_set,
                    resume=True,
                    progress_cb=cb_c006,
                )
                status_box.write(
                    f"   - `C006` 전체 페이지네이션 스캔 완료 (전체 {c006_stats['total_count']:,}건, "
                    f"신규 매칭 저장 +{c006_stats['saved_rows']:,}건)"
                )
            else:
                include_pkg = c006_mode == "all_lcns"
                c006_stats = client.collect_c006_by_licenses(
                    target_businesses=regional_businesses,
                    include_packaging_shops=include_pkg,
                    resume=True,
                    progress_cb=cb_c006,
                )
                status_box.write(
                    f"   - `C006` 인허가번호별 매핑 완료: 대상 {c006_stats['total_targets']:,}개소 "
                    f"(신규 조회 {c006_stats['queried_now']:,}개소, 캐시 재사용 {c006_stats['cached_skipped']:,}개소)"
                )
        except InvalidApiKeyError as exc:
            if "INFO-110" in str(exc):
                status_box.write(
                    "   - ⚠️ `C006`(축산물품목제조보고) 활용신청 전(`INFO-110`) → 현재 허가업종(`I1300` 식육가공업 + `I1320` 식육포장처리업) 기준으로 선행 분류합니다."
                )
            else:
                raise

        # 최종 DataFrame 병합·정제 및 스냅샷 저장
        progress_bar.progress(0.97, text="최종 데이터 정제 및 양념육·세절육 가공구분 집계 중...")
        final_df = rebuild_dataframe_from_cache(
            cache_store=cache_store,
            selected_regions=selected_regions,
            merge_same_facility=merge_same_facility,
        )
        st.session_state["processed_df"] = final_df
        st.session_state["last_updated"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # 선택적 Parquet 스냅샷 보조 저장 (pyarrow/fastparquet 설치 환경 시 자동 저장, 미설치 시 SQLite만 유지)
        try:
            PARQUET_SNAPSHOT_PATH.parent.mkdir(parents=True, exist_ok=True)
            final_df.to_parquet(PARQUET_SNAPSHOT_PATH, index=False)
        except Exception:
            pass

        progress_bar.progress(1.0, text=f"✅ 수집 및 정제 완료! 총 {len(final_df):,}개 가공장 구축")
        status_box.update(
            label=f"✅ 수집 완료: 총 {len(final_df):,}개 양념육·세절육 가공장 데이터 구축됨",
            state="complete",
            expanded=False,
        )

    except QuotaExceededError as exc:
        # 호출 한도 도달 시에도 지금까지 캐시에 저장된 데이터로 즉시 결과 표 및 엑셀 제공
        partial_df = rebuild_dataframe_from_cache(
            cache_store=cache_store,
            selected_regions=selected_regions,
            merge_same_facility=merge_same_facility,
        )
        st.session_state["processed_df"] = partial_df
        st.session_state["last_updated"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S") + " (부분완료)"
        status_box.update(
            label="⚠️ OpenAPI 호출 한도 도달 (진행분 SQLite 캐시 안전 저장 완료)",
            state="error",
            expanded=True,
        )
        st.warning(
            f"⚠️ {exc}\n\n"
            f"현재까지 수집된 캐시 기준으로 **{len(partial_df):,}개소**의 데이터를 아래 화면에서 바로 조회·다운로드할 수 있으며, "
            "자정(00:00) 한도 초기화 후 다시 **데이터 수집 실행**을 누르면 중단된 지점부터 자동 이어받기됩니다."
        )
    except InvalidApiKeyError as exc:
        status_box.update(label="❌ 인증키 오류", state="error", expanded=True)
        st.error(f"❌ {exc}")
    except FoodSafetyAPIError as exc:
        status_box.update(label="❌ OpenAPI 통신 오류", state="error", expanded=True)
        st.error(f"❌ OpenAPI 오류: {exc} (진행 중이던 데이터는 SQLite 캐시에 자동 저장되었습니다.)")
    except Exception as exc:
        status_box.update(label="❌ 수집 중 예외 발생", state="error", expanded=True)
        st.error(f"❌ 오류 발생: {exc}")


def render_main_dashboard(df: pd.DataFrame) -> None:
    """메인 화면 요약 지표(Metric 카드), 검색/필터 데이터프레임 미리보기 및 멀티시트 엑셀 다운로드"""
    st.subheader("📊 지역별 · 가공구분별 가공장 현황 요약")

    total_cnt = len(df)
    seasoned_cnt = int((df["가공구분"] == CATEGORY_SEASONED).sum()) if not df.empty else 0
    sliced_cnt = int((df["가공구분"] == CATEGORY_SLICED).sum()) if not df.empty else 0
    both_cnt = int((df["가공구분"] == CATEGORY_BOTH).sum()) if not df.empty else 0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("전체 대상 가공장", f"{total_cnt:,}개소")
    m2.metric("양념육 가공장", f"{seasoned_cnt:,}개소")
    m3.metric("식육포장처리(세절·포장육)", f"{sliced_cnt:,}개소")
    m4.metric("양념육 + 세절·포장육 겸업", f"{both_cnt:,}개소")

    r1, r2, r3, r4 = st.columns(4)
    cn_cnt = int((df["시/도"] == "충청남도").sum()) if not df.empty else 0
    cb_cnt = int((df["시/도"] == "충청북도").sum()) if not df.empty else 0
    dj_cnt = int((df["시/도"] == "대전광역시").sum()) if not df.empty else 0
    jb_cnt = int((df["시/도"] == "전북특별자치도(전라북도)").sum()) if not df.empty else 0

    r1.metric("충청남도", f"{cn_cnt:,}개소")
    r2.metric("충청북도", f"{cb_cnt:,}개소")
    r3.metric("대전광역시", f"{dj_cnt:,}개소")
    r4.metric("전북특별자치도(전라북도)", f"{jb_cnt:,}개소")

    if df.empty:
        st.info(
            "💡 아직 표시할 수집 데이터가 없습니다. 좌측 사이드바에서 **식품안전나라 OpenAPI 인증키**를 입력하고 "
            "**🚀 데이터 수집 실행** 버튼을 클릭해주세요."
        )
        return

    # 지역 x 가공구분 교차 요약표 (접기/펼치기)
    with st.expander("📈 지역별 × 가공구분 교차 집계표 및 연락처 확보율 보기", expanded=False):
        pivot_df = pd.crosstab(
            df["시/도"],
            df["가공구분"],
            margins=True,
            margins_name="합계",
        )
        tel_valid_cnt = int((df["연락처(전화번호)"].str.strip() != "").sum())
        tel_rate = (tel_valid_cnt / total_cnt * 100.0) if total_cnt > 0 else 0.0
        st.caption(
            f"전화번호 기재 업체: **{tel_valid_cnt:,}개소 / {total_cnt:,}개소 ({tel_rate:.1f}%)**"
        )
        st.dataframe(pivot_df, width="stretch")

    st.divider()

    # 엑셀(.xlsx) 원클릭 다운로드 영역
    dl_col1, dl_col2 = st.columns([2, 3])
    today_str = datetime.now().strftime("%Y%m%d")
    excel_bytes = generate_multisheet_excel_bytes(df)

    with dl_col1:
        st.download_button(
            label=f"📥 전체 결과 엑셀(.xlsx) 다운로드 (5개 시트 · {len(df):,}건)",
            data=excel_bytes,
            file_name=f"충청전북_양념육_세절육_가공장현황_{today_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            width="stretch",
        )
    with dl_col2:
        last_ts = st.session_state.get("last_updated", "-")
        st.caption(
            "📌 엑셀 시트 자동 구성: `[전체_통합]`, `[충청남도]`, `[충청북도]`, `[대전광역시]`, `[전북특별자치도]` "
            f"(1행 헤더 고정·자동필터·열 너비 자동 조절 완료 | 최근 집계: {last_ts})"
        )

    st.divider()
    st.subheader("🔎 가공장 상세 목록 조회 및 검색")

    f_col1, f_col2, f_col3, f_col4 = st.columns([1.3, 1.4, 1.8, 0.9])
    available_sidos = list(df["시/도"].unique())
    available_cats = list(df["가공구분"].unique())

    sel_sidos = f_col1.multiselect(
        "시/도 필터",
        options=available_sidos,
        default=available_sidos,
    )
    sel_cats = f_col2.multiselect(
        "가공구분 필터",
        options=available_cats,
        default=available_cats,
    )
    keyword = f_col3.text_input(
        "통합 키워드 검색",
        placeholder="업체명, 시/군/구, 주소, 생산품목(불고기/갈비 등), 인허가번호 검색",
    ).strip()
    only_with_phone = f_col4.checkbox("연락처 보유만", value=False)

    filtered_df = df[
        df["시/도"].isin(sel_sidos) & df["가공구분"].isin(sel_cats)
    ].copy()

    if only_with_phone:
        filtered_df = filtered_df[filtered_df["연락처(전화번호)"].str.strip() != ""]

    if keyword:
        kw_lower = keyword.lower()
        search_target = (
            filtered_df["업체명"].fillna("")
            + " "
            + filtered_df["시/군/구"].fillna("")
            + " "
            + filtered_df["소재지 주소"].fillna("")
            + " "
            + filtered_df["주요 품목유형"].fillna("")
            + " "
            + filtered_df["대표 생산품목 리스트"].fillna("")
            + " "
            + filtered_df["대표자"].fillna("")
            + " "
            + filtered_df["인허가번호"].fillna("")
        ).str.lower()
        filtered_df = filtered_df[search_target.str.contains(kw_lower, regex=False)]

    filtered_df = filtered_df.reset_index(drop=True)

    sub_bar1, sub_bar2 = st.columns([3, 1.5])
    sub_bar1.markdown(
        f"검색 결과: **{len(filtered_df):,}개소** (전체 `{len(df):,}`개소 중)"
    )
    if len(filtered_df) != len(df) and not filtered_df.empty:
        filtered_excel = generate_multisheet_excel_bytes(filtered_df)
        sub_bar2.download_button(
            label=f"📥 현재 필터 결과만 엑셀 다운로드 ({len(filtered_df):,}건)",
            data=filtered_excel,
            file_name=f"충청전북_가공장_필터결과_{today_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

    st.dataframe(
        filtered_df[OUTPUT_COLUMNS],
        width="stretch",
        height=560,
        hide_index=True,
        column_config={
            "등록 품목수": st.column_config.NumberColumn("등록 품목수", format="%d건"),
        },
    )


def main() -> None:
    st.title("🥩 충청·전북 양념육 및 세절·포장육 가공장 통합 수집기")
    st.markdown(
        "식품의약품안전처(식품안전나라) OpenAPI **`I2500`(인허가 업소)** · **`I1320`(식육포장처리업 허가대장)** · "
        "**`C006`(축산물 품목제조보고)** 실데이터를 연동하여 **충청남도·충청북도·대전광역시·전북특별자치도(구 전라북도)** "
        "소재 양념육 및 세절·포장육 가공장을 자동 병합·분류하고 멀티시트 엑셀(`.xlsx`)로 내보냅니다."
    )

    cache_store = get_cache_store()
    controls = render_sidebar(cache_store)

    if controls["clear_cache_clicked"]:
        cache_store.clear_all()
        if PARQUET_SNAPSHOT_PATH.exists():
            try:
                PARQUET_SNAPSHOT_PATH.unlink()
            except OSError:
                pass
        st.session_state.pop("processed_df", None)
        st.session_state.pop("last_updated", None)
        st.toast("로컬 SQLite 캐시가 초기화되었습니다.", icon="🗑️")
        st.rerun()

    if controls["test_clicked"]:
        handle_connection_test(str(controls["api_key"]), cache_store)

    if controls["run_clicked"]:
        run_collection_pipeline(
            api_key=str(controls["api_key"]),
            selected_regions=list(controls["selected_regions"]),  # type: ignore[arg-type]
            c006_mode=str(controls["c006_mode"]),
            merge_same_facility=bool(controls["merge_same_facility"]),
            cache_store=cache_store,
        )
    elif controls["refresh_cache_clicked"] or "processed_df" not in st.session_state:
        cached_summary = cache_store.get_cache_summary()
        if cached_summary["business_count"] > 0:
            st.session_state["processed_df"] = rebuild_dataframe_from_cache(
                cache_store=cache_store,
                selected_regions=list(controls["selected_regions"]),  # type: ignore[arg-type]
                merge_same_facility=bool(controls["merge_same_facility"]),
            )
            if "last_updated" not in st.session_state:
                st.session_state["last_updated"] = "로컬 SQLite 캐시 복원본"

    current_df = st.session_state.get("processed_df", pd.DataFrame(columns=OUTPUT_COLUMNS))
    render_main_dashboard(current_df)


if __name__ == "__main__":
    main()
