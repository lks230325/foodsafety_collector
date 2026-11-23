# -*- coding: utf-8 -*-
"""
충남·충북·대전·전북 양념육 및 세절·포장육 가공장 필터링, 병합, 정제 프로세서
"""

from __future__ import annotations

import re
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set, Tuple

import pandas as pd

# 최종 산출물 및 엑셀 다운로드 표준 컬럼 순서 (13개 컬럼)
OUTPUT_COLUMNS: List[str] = [
    "시/도",
    "시/군/구",
    "가공구분",
    "업체명",
    "대표자",
    "인허가업종",
    "연락처(전화번호)",
    "소재지 주소",
    "주요 품목유형",
    "대표 생산품목 리스트",
    "등록 품목수",
    "인허가번호",
    "허가일자",
]

# 가공구분 표준 명칭
CATEGORY_SEASONED = "양념육 가공장"
CATEGORY_SLICED = "식육포장처리(세절·포장육)"
CATEGORY_BOTH = "양념육 + 세절·포장육 겸업"

# 대상 지역 매핑 정의 (UI 체크박스 키 -> (표준 시/도 명칭, 시작 접두어 튜플))
REGION_RULES: Dict[str, Tuple[str, Tuple[str, ...]]] = {
    "충남": ("충청남도", ("충청남도", "충남")),
    "충북": ("충청북도", ("충청북도", "충북")),
    "대전": ("대전광역시", ("대전광역시", "대전시", "대전")),
    # 전북은 구 '전라북도'와 신규 '전북특별자치도'를 모두 수집 후 '전북특별자치도(전라북도)'로 통일
    "전북": ("전북특별자치도(전라북도)", ("전북특별자치도", "전라북도", "전북")),
}

SEASONED_KEYWORDS: Tuple[str, ...] = (
    "양념육",
    "양념",
    "불고기",
    "갈비",
    "주물럭",
    "분쇄가공육",
    "육가공",
)

SLICED_KEYWORDS: Tuple[str, ...] = (
    "포장육",
    "세절",
    "절단",
    "슬라이스",
    "정육",
)

_LEADING_ZIP_RE = re.compile(r"^[\(\[\s]*\d{5,6}[\)\]\s,-]*")
_MULTISPACE_RE = re.compile(r"\s+")


def clean_address(raw_addr: Any) -> str:
    """주소 앞뒤 공백 및 불필요한 우편번호 표기를 정리한다."""
    if raw_addr is None:
        return ""
    addr = str(raw_addr).strip()
    if not addr or addr.lower() in ("nan", "none", "null", "-"):
        return ""
    addr = _LEADING_ZIP_RE.sub("", addr)
    return _MULTISPACE_RE.sub(" ", addr).strip()


def parse_target_region(
    raw_addr: Any, selected_region_keys: Optional[Sequence[str]] = None
) -> Optional[Tuple[str, str, str]]:
    """
    소재지 주소(ADDR)가 대상 지역(충남, 충북, 대전, 전북)으로 시작하는지 판별한다.

    Returns:
        매칭 시 (region_key, 표준_시도명, 시군구명), 비대상 지역이면 None
    """
    addr = clean_address(raw_addr)
    if not addr:
        return None

    active_keys = (
        list(selected_region_keys)
        if selected_region_keys is not None
        else list(REGION_RULES.keys())
    )

    for reg_key in active_keys:
        if reg_key not in REGION_RULES:
            continue
        canonical_sido, prefixes = REGION_RULES[reg_key]
        # 긴 접두어('전북특별자치도', '충청남도' 등)부터 먼저 매칭
        for prefix in sorted(prefixes, key=len, reverse=True):
            if not addr.startswith(prefix):
                continue
            rest = addr[len(prefix) :]
            # 2글자 약칭(충남, 충북, 대전, 전북)은 공백 또는 시/군/구 행정구역명이 이어질 때만 인정 ('대전로' 오탐 방지)
            if len(prefix) == 2 and rest:
                if not (
                    rest[0].isspace()
                    or re.match(r"^[가-힣]{1,4}[시군구](?:\s|$)", rest)
                ):
                    continue
            rest_clean = rest.strip()
            sigungu = _extract_sigungu(canonical_sido, rest_clean)
            return reg_key, canonical_sido, sigungu

    return None


def _extract_sigungu(canonical_sido: str, rest_addr: str) -> str:
    """시/도 이후 주소 문자열에서 시/군/구 행정구역명을 추출한다."""
    if not rest_addr:
        return ""
    tokens = rest_addr.split()
    if not tokens:
        return ""

    first = tokens[0].strip(",")
    if canonical_sido == "대전광역시":
        if first.endswith(("구", "군")):
            return first
        return first

    # 도 단위(충남, 충북, 전북): '천안시 서북구', '청주시 흥덕구', '전주시 덕진구' 등 일반구 포함
    if first.endswith("시") and len(tokens) >= 2 and tokens[1].strip(",").endswith("구"):
        return f"{first} {tokens[1].strip(',')}"
    if first.endswith(("시", "군", "구")):
        return first
    return first


def filter_businesses_by_region(
    businesses: Sequence[Dict[str, Any]],
    selected_region_keys: Optional[Sequence[str]] = None,
) -> List[Dict[str, Any]]:
    """1차 수집된 전국 업소 리스트에서 선택된 충청·전북권 업소만 필터링한다."""
    filtered: List[Dict[str, Any]] = []
    for biz in businesses:
        parsed = parse_target_region(biz.get("ADDR", ""), selected_region_keys)
        if parsed is None:
            continue
        reg_key, sido, sigungu = parsed
        enriched = dict(biz)
        enriched["_REGION_KEY"] = reg_key
        enriched["_SIDO"] = sido
        enriched["_SIGUNGU"] = sigungu
        enriched["ADDR"] = clean_address(biz.get("ADDR", ""))
        filtered.append(enriched)
    return filtered


def format_korean_phone(raw_tel: Any) -> str:
    """
    전화번호(TELNO)를 한국 표준 하이픈(-) 포맷으로 정제하고 결측치·무효값은 공란('') 처리한다.
    """
    if raw_tel is None:
        return ""
    text = str(raw_tel).strip()
    if not text or text.lower() in ("nan", "none", "null", "-", "--", "없음", "미상"):
        return ""

    # 복수 번호가 '/' 또는 ','로 기재된 경우 첫 번째 유효 번호 사용
    first_chunk = re.split(r"[/,;~]", text)[0].strip()
    digits = re.sub(r"[^0-9]", "", first_chunk)

    # 국제번호(+82) 보정
    if digits.startswith("82") and len(digits) >= 10:
        digits = "0" + digits[2:]

    # 모두 0이거나 너무 짧은 무효 번호 공란 처리
    if not digits or set(digits) == {"0"} or len(digits) < 7 or len(digits) > 12:
        return ""

    # 대표번호 (1588-XXXX, 1644-XXXX, 1899-XXXX 등 8자리)
    if len(digits) == 8 and digits[:2] in ("15", "16", "18"):
        return f"{digits[:4]}-{digits[4:]}"

    # 서울 지역번호 (02)
    if digits.startswith("02"):
        if len(digits) == 9:
            return f"02-{digits[2:5]}-{digits[5:]}"
        if len(digits) == 10:
            return f"02-{digits[2:6]}-{digits[6:]}"
        return ""

    # 안심번호·평생번호 (050X - 11~12자리)
    if digits.startswith("050") and len(digits) in (11, 12):
        mid = len(digits) - 4
        return f"{digits[:4]}-{digits[4:mid]}-{digits[mid:]}"

    # 일반 지역번호(041 충남, 042 대전, 043 충북, 063 전북 등) 및 휴대폰/인터넷전화(010, 070, 080)
    if digits.startswith("0"):
        if len(digits) == 10:
            return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
        if len(digits) == 11:
            return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
        # 간혹 9자리 오기재 방어
        if len(digits) == 9:
            return f"{digits[:3]}-{digits[3:5]}-{digits[5:]}"
        return ""

    # 지역번호가 생략된 7~8자리 국번 데이터 보정
    if len(digits) == 7:
        return f"{digits[:3]}-{digits[3:]}"
    if len(digits) == 8:
        return f"{digits[:4]}-{digits[4:]}"

    return ""


def format_permit_date(raw_dt: Any) -> str:
    """허가일자(YYYYMMDD)를 YYYY-MM-DD 형식으로 정제한다."""
    if raw_dt is None:
        return ""
    s = re.sub(r"[^0-9]", "", str(raw_dt).strip())
    if len(s) == 8 and s != "00000000":
        return f"{s[:4]}-{s[4:6]}-{s[6:8]}"
    raw_clean = str(raw_dt).strip()
    return "" if raw_clean.lower() in ("nan", "none", "null", "0", "-") else raw_clean


def _normalize_facility_key(bssh_nm: str, addr: str) -> Optional[Tuple[str, str]]:
    """동일 사업장(업체명 + 도로명/지번 핵심 주소) 병합 키 생성"""
    clean_name = re.sub(r"\s+", "", (bssh_nm or "").strip())
    if not clean_name or not addr:
        return None
    # 괄호 상세동·층수 표기 차이 보정 전 기본 주소 정규화
    addr_norm = re.sub(r"\s+", " ", addr.strip())
    return (clean_name, addr_norm)


def _deduplicate_strings(values: Iterable[str]) -> List[str]:
    seen: Set[str] = set()
    result: List[str] = []
    for v in values:
        s = (v or "").strip()
        if s and s not in seen:
            seen.add(s)
            result.append(s)
    return result


def classify_and_summarize(
    induty_cd_nm: str,
    c006_items: Sequence[Dict[str, Any]],
    c006_mapped_globally: bool = True,
) -> Tuple[Optional[str], str, str, int]:
    """
    인허가업종 및 C006 품목제조보고 리스트를 기반으로:
    1) 가공구분 ('양념육 가공장' / '식육포장처리(세절·포장육)' / '양념육 + 세절·포장육 겸업') 판별
    2) 주요 품목유형 (중복 제거된 유형명 나열)
    3) 대표 생산품목 리스트 (상위 10개 외 N건)
    4) 등록 품목수 (품목제조보고번호/품목명 기준 고유 품목 수)
    를 계산한다.
    """
    induty = (induty_cd_nm or "").strip()
    is_meat_proc_induty = ("식육가공업" in induty) or ("축산물가공업" in induty)
    is_packaging_induty = "식육포장처리업" in induty

    # C006 원재료 행 중복 제거 (품목제조보고번호 우선, 없으면 품목명 기준)
    unique_products: Dict[str, Dict[str, str]] = {}
    for item in c006_items:
        report_no = str(item.get("PRDLST_REPORT_NO", "") or "").strip()
        prd_nm = str(item.get("PRDLST_NM", "") or "").strip()
        dcnm = str(item.get("PRDLST_DCNM", "") or "").strip()
        if not prd_nm and not dcnm:
            continue
        dedup_key = report_no if report_no else f"NM:{prd_nm}"
        if dedup_key not in unique_products:
            unique_products[dedup_key] = {
                "PRDLST_NM": prd_nm,
                "PRDLST_DCNM": dcnm,
            }

    dedup_list = list(unique_products.values())
    product_count = len(dedup_list)

    has_seasoned_item = False
    has_sliced_item = False

    priority_names: List[str] = []
    other_names: List[str] = []
    priority_types: List[str] = []
    other_types: List[str] = []

    for prod in dedup_list:
        prd_nm = prod["PRDLST_NM"]
        dcnm = prod["PRDLST_DCNM"]
        combined_text = f"{dcnm} {prd_nm}"

        matched_s = any(kw in combined_text for kw in SEASONED_KEYWORDS)
        matched_p = any(kw in combined_text for kw in SLICED_KEYWORDS)

        if matched_s:
            has_seasoned_item = True
        if matched_p:
            has_sliced_item = True

        if matched_s or matched_p:
            if prd_nm:
                priority_names.append(prd_nm)
            if dcnm:
                priority_types.append(dcnm)
        else:
            if prd_nm:
                other_names.append(prd_nm)
            if dcnm:
                other_types.append(dcnm)

    # 양념육 가공장 조건:
    # 1) C006 데이터가 연동된 상태라면 C006 유형/품목에 양념육 관련 키워드가 포함된 업체
    # 2) C006 활용신청 전(c006_mapped_globally=False)이라면 '식육가공업' 인허가 보유 업체를 우선 포함 (유가공/알가공 제외)
    is_strictly_meat_proc = "식육가공업" in induty
    if c006_mapped_globally:
        is_seasoned = (is_meat_proc_induty and has_seasoned_item) or has_seasoned_item
    else:
        is_seasoned = is_strictly_meat_proc or has_seasoned_item

    # 세절·포장육 가공장 조건: 식육포장처리업(신고 면제 기본 인정) 또는 포장육/세절 관련 품목 보유
    is_sliced = is_packaging_induty or has_sliced_item

    if is_seasoned and is_sliced:
        category: Optional[str] = CATEGORY_BOTH
    elif is_seasoned:
        category = CATEGORY_SEASONED
    elif is_sliced:
        category = CATEGORY_SLICED
    else:
        category = None

    unique_types = _deduplicate_strings(priority_types + other_types)
    if not unique_types and not c006_mapped_globally and is_strictly_meat_proc:
        type_summary = "식육가공업 (C006 품목보고 매핑 대기)"
    else:
        type_summary = ", ".join(unique_types)

    unique_prd_names = _deduplicate_strings(priority_names + other_names)
    if len(unique_prd_names) <= 10:
        product_summary = ", ".join(unique_prd_names)
    else:
        top10 = ", ".join(unique_prd_names[:10])
        remain = len(unique_prd_names) - 10
        product_summary = f"{top10} 외 {remain:,}건"

    return category, type_summary, product_summary, product_count


def build_processed_dataframe(
    regional_businesses: Sequence[Dict[str, Any]],
    c006_by_lcns: Dict[str, List[Dict[str, Any]]],
    merge_same_facility: bool = True,
) -> pd.DataFrame:
    """
    지역 필터링된 업소 리스트와 C006 품목 데이터를 결합하여 최종 표준 DataFrame을 생성한다.

    Args:
        regional_businesses: filter_businesses_by_region 결과 리스트
        c006_by_lcns: {LCNS_NO: [C006 item dicts]}
        merge_same_facility: 동일 사업장(업체명+주소 동일)이 식육가공업과 식육포장처리업 별도 인허가를
                             동시 보유한 경우 하나로 병합하여 '양념육 + 세절·포장육 겸업' 판별 정확도 향상
    """
    if not regional_businesses:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    grouped_units: List[Dict[str, Any]] = []

    if merge_same_facility:
        facility_map: Dict[Tuple[str, str], Dict[str, Any]] = {}
        for biz in regional_businesses:
            bssh_nm = str(biz.get("BSSH_NM", "") or "").strip()
            addr = str(biz.get("ADDR", "") or "").strip()
            lcns_no = str(biz.get("LCNS_NO", "") or "").strip()
            fkey = _normalize_facility_key(bssh_nm, addr)

            if fkey is None:
                grouped_units.append(
                    {
                        "sido": biz.get("_SIDO", ""),
                        "sigungu": biz.get("_SIGUNGU", ""),
                        "bssh_nm": bssh_nm,
                        "prsdnt_list": [str(biz.get("PRSDNT_NM", "") or "").strip()],
                        "induty_list": [
                            x.strip()
                            for x in str(biz.get("INDUTY_CD_NM", "") or "").split(",")
                            if x.strip()
                        ],
                        "tel_list": [str(biz.get("TELNO", "") or "").strip()],
                        "addr": addr,
                        "lcns_list": [lcns_no] if lcns_no else [],
                        "prms_list": [str(biz.get("PRMS_DT", "") or "").strip()],
                    }
                )
                continue

            if fkey not in facility_map:
                unit = {
                    "sido": biz.get("_SIDO", ""),
                    "sigungu": biz.get("_SIGUNGU", ""),
                    "bssh_nm": bssh_nm,
                    "prsdnt_list": [str(biz.get("PRSDNT_NM", "") or "").strip()],
                    "induty_list": [
                        x.strip()
                        for x in str(biz.get("INDUTY_CD_NM", "") or "").split(",")
                        if x.strip()
                    ],
                    "tel_list": [str(biz.get("TELNO", "") or "").strip()],
                    "addr": addr,
                    "lcns_list": [lcns_no] if lcns_no else [],
                    "prms_list": [str(biz.get("PRMS_DT", "") or "").strip()],
                }
                facility_map[fkey] = unit
                grouped_units.append(unit)
            else:
                unit = facility_map[fkey]
                unit["prsdnt_list"].append(str(biz.get("PRSDNT_NM", "") or "").strip())
                unit["induty_list"].extend(
                    x.strip()
                    for x in str(biz.get("INDUTY_CD_NM", "") or "").split(",")
                    if x.strip()
                )
                unit["tel_list"].append(str(biz.get("TELNO", "") or "").strip())
                if lcns_no:
                    unit["lcns_list"].append(lcns_no)
                unit["prms_list"].append(str(biz.get("PRMS_DT", "") or "").strip())
    else:
        for biz in regional_businesses:
            lcns_no = str(biz.get("LCNS_NO", "") or "").strip()
            grouped_units.append(
                {
                    "sido": biz.get("_SIDO", ""),
                    "sigungu": biz.get("_SIGUNGU", ""),
                    "bssh_nm": str(biz.get("BSSH_NM", "") or "").strip(),
                    "prsdnt_list": [str(biz.get("PRSDNT_NM", "") or "").strip()],
                    "induty_list": [
                        x.strip()
                        for x in str(biz.get("INDUTY_CD_NM", "") or "").split(",")
                        if x.strip()
                    ],
                    "tel_list": [str(biz.get("TELNO", "") or "").strip()],
                    "addr": str(biz.get("ADDR", "") or "").strip(),
                    "lcns_list": [lcns_no] if lcns_no else [],
                    "prms_list": [str(biz.get("PRMS_DT", "") or "").strip()],
                }
            )

    has_global_c006 = any(len(v) > 0 for v in c006_by_lcns.values())
    records: List[Dict[str, Any]] = []
    for unit in grouped_units:
        lcns_list = _deduplicate_strings(unit["lcns_list"])
        induty_joined = ", ".join(_deduplicate_strings(unit["induty_list"]))

        combined_items: List[Dict[str, Any]] = []
        for lc in lcns_list:
            combined_items.extend(c006_by_lcns.get(lc, []))

        category, type_summary, product_summary, product_count = classify_and_summarize(
            induty_cd_nm=induty_joined,
            c006_items=combined_items,
            c006_mapped_globally=has_global_c006,
        )

        # 양념육 및 세절·포장육 가공장 대상에 해당하지 않는 비대상 업소(유가공·알가공·단순 육포/소시지 등) 제외
        if category is None:
            continue

        # 유효 전화번호 선택
        cleaned_tel = ""
        for raw_t in unit["tel_list"]:
            formatted = format_korean_phone(raw_t)
            if formatted:
                cleaned_tel = formatted
                break

        prsdnt_clean = next(
            (p for p in _deduplicate_strings(unit["prsdnt_list"]) if p and p != "-"),
            "",
        )

        valid_dates = sorted(
            d for d in (format_permit_date(x) for x in unit["prms_list"]) if d
        )
        permit_date = valid_dates[0] if valid_dates else ""

        records.append(
            {
                "시/도": unit["sido"],
                "시/군/구": unit["sigungu"],
                "가공구분": category,
                "업체명": unit["bssh_nm"],
                "대표자": prsdnt_clean,
                "인허가업종": induty_joined,
                "연락처(전화번호)": cleaned_tel,
                "소재지 주소": unit["addr"],
                "주요 품목유형": type_summary,
                "대표 생산품목 리스트": product_summary,
                "등록 품목수": int(product_count),
                "인허가번호": ", ".join(lcns_list),
                "허가일자": permit_date,
            }
        )

    df = pd.DataFrame(records, columns=OUTPUT_COLUMNS)
    if not df.empty:
        # 정렬: 시/도 -> 가공구분(겸업·양념육 우선) -> 시/군/구 -> 업체명
        cat_order = {
            CATEGORY_BOTH: 0,
            CATEGORY_SEASONED: 1,
            CATEGORY_SLICED: 2,
        }
        sido_order = {
            "충청남도": 0,
            "충청북도": 1,
            "대전광역시": 2,
            "전북특별자치도(전라북도)": 3,
        }
        df["_sido_rank"] = df["시/도"].map(lambda x: sido_order.get(x, 99))
        df["_cat_rank"] = df["가공구분"].map(lambda x: cat_order.get(x, 99))
        df = (
            df.sort_values(
                by=["_sido_rank", "_cat_rank", "시/군/구", "업체명"],
                ascending=[True, True, True, True],
            )
            .drop(columns=["_sido_rank", "_cat_rank"])
            .reset_index(drop=True)
        )

    return df
