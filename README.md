# 충청·전북 양념육 및 세절·포장육 가공장 수집기 (식품안전나라 OpenAPI)

식품의약품안전처(식품안전나라) 실제 OpenAPI 3종(`I2500`, `I1320`, `C006`)을 연동하여 **충청남도, 충청북도, 대전광역시, 전북특별자치도(구 전라북도)** 지역의 **양념육** 및 **세절·포장육(식육포장처리)** 가공장 주소지·연락처·주요 생산품목을 수집·병합하고 5개 시트로 구성된 엑셀(`.xlsx`) 파일로 다운로드하는 Streamlit 웹 애플리케이션입니다.

---

## 1. 연동 OpenAPI 규격

기본 호출 URL:
```text
http://openapi.foodsafetykorea.go.kr/api/{API_KEY}/{SERVICE_ID}/json/{START_IDX}/{END_IDX}/{추가파라미터}
```
- 1회 최대 `1,000`건 단위 자동 페이지네이션 (`total_count` 도달 시까지 순회, `RESULT.CODE == "INFO-000"` 검증)
- `requests.Session` 커넥션 풀링 + `tenacity` 지수 백오프(Exponential Backoff) 자동 재시도 적용
- `data/foodsafety_cache.sqlite` 로컬 SQLite 캐시를 통해 네트워크 단절 또는 일일 호출 한도(`INFO-300`) 도달 시에도 **중단 지점부터 이어받기** 지원

| 서비스 ID | 서비스명 | 역할 및 활용 필드 |
| :--- | :--- | :--- |
| **`I2500`** | 영업 중인 인허가 업소 정보 | `INDUTY_CD_NM=식육가공업`, `식육포장처리업`, `축산물가공업` 각각 조회 후 `LCNS_NO` 기준 합집합 구성 (`LCNS_NO`, `BSSH_NM`, `INDUTY_CD_NM`, `PRSDNT_NM`, `TELNO`, `ADDR`, `PRMS_DT`) |
| **`I1320`** | 축산물 식육포장처리업 영업허가대장 | `I2500`에서 누락된 식육포장처리업소를 `LCNS_NO` 기준으로 보완 병합 및 공란 전화번호·대표자 보완 (`LOCP_ADDR` → `ADDR`, `INDUTY_NM` → `INDUTY_CD_NM`, 폐업·취소 제외) |
| **`C006`** | 축산물 품목제조보고(원재료) | 대상 지역 업소의 생산품목(`PRDLST_NM`), 축산물 유형(`PRDLST_DCNM`), 원재료(`RAWMTRL_NM`) 매핑 (`PRDLST_REPORT_NO` 기준 중복 원재료 행 제거) |

---

## 2. 지역 및 가공구분 정제 로직

- **대상 지역 필터링 (`ADDR` 기준)**
  - `충청남도`(`충남`) → `충청남도`
  - `충청북도`(`충북`) → `충청북도`
  - `대전광역시`(`대전`) → `대전광역시`
  - `전북특별자치도` 및 구 `전라북도`(`전북`) → **`전북특별자치도(전라북도)`** 단일 명칭으로 통합
- **가공구분 분류**
  - **양념육 가공장**: `식육가공업`/`축산물가공업` 중 `C006` 유형(`PRDLST_DCNM`) 또는 품목명(`PRDLST_NM`)에 `양념육`, `양념`, `불고기`, `갈비`, `주물럭`, `분쇄가공육`, `육가공` 포함 업체
  - **식육포장처리(세절·포장육)**: 업종이 `식육포장처리업`(품목제조보고 면제 대상 기본 인정)이거나 품목/유형에 `포장육`, `세절`, `절단`, `슬라이스`, `정육` 포함 업체
  - **양념육 + 세절·포장육 겸업**: 두 조건을 모두 충족하는 업체 (동일 사업장 복수 인허가 자동 병합 지원)
- **연락처 및 품목 요약**
  - 정규식 기반 한국 전화번호 하이픈(`-`) 표준화 및 결측·무효 번호 공란 처리
  - 중복 제거된 `주요 품목유형` 및 `대표 생산품목 리스트`(상위 10개 외 `N`건) 자동 요약

---

## 3. 프로젝트 구조

```text
foodsafety_collector/
├── app.py              # Streamlit 메인 대시보드 UI 및 수집 파이프라인 실행
├── api_client.py       # 식품안전나라 OpenAPI 수집기 (Session + tenacity + SQLite 이어받기 캐시)
├── processor.py        # 지역 필터링, 전북 명칭 통합, 가공구분 분류, 연락처/품목 정제
├── excel_exporter.py   # openpyxl 기반 5개 시트 분리 및 헤더 고정/자동필터/한글 열너비 조절
├── requirements.txt    # 필수 파이썬 패키지 목록
├── .env.example        # 환경변수 설정 예시 파일 (FOOD_API_KEY)
├── run_app.bat         # Windows 원클릭 실행 스크립트
└── README.md           # 사용 설명서
```

---

## 4. 설치 및 실행 방법

> **현재 PC 환경 설정 완료 안내**: 현재 PC의 승인된 Python 3.12 환경에 `streamlit`, `tenacity`, `pyarrow`, `openpyxl`, `pandas`, `requests`, `python-dotenv` 등 모든 필수 라이브러리 설치가 완료되어 있습니다. 추가 설치 없이 바로 실행 가능합니다.

### 1) OpenAPI 인증키 설정
[식품안전나라 OpenAPI 홈페이지](https://www.foodsafetykorea.go.kr/apiMain.do)에서 발급받은 인증키를 `.env` 파일의 `FOOD_API_KEY=` 뒤에 저장하거나, 실행된 웹 대시보드 좌측 사이드바 입력창에 직접 붙여넣으세요.

### 2) 애플리케이션 실행
Windows 탐색기에서 [`run_app.bat`](run_app.bat)을 더블클릭하거나 터미널에서 아래 명령으로 실행합니다.
```powershell
cd foodsafety_collector
.\run_app.bat
```
- 접속 주소: `http://localhost:8501`
- 정식 키 입력 전 통신 확인: 인증키 입력란에 `sample` 입력 후 **`🔌 샘플 5건 테스트`** 클릭

---

## 5. GitHub Pages 무료 정적 웹 대시보드 배포 (AES-256-GCM 비밀번호 보안 적용)

- **배포 URL**: https://lks230325.github.io/foodsafety_collector/
- **보안 구조**:
  - `.env`에 설정된 `DASHBOARD_PASSWORD`를 이용해 데이터 JSON 및 5개 시트 `.xlsx` 원본을 **PBKDF2-HMAC-SHA256 (600,000회) + AES-256-GCM**으로 암호화한 파일(`static_data/*.enc.json`)만 Git 및 배포 서버에 업로드합니다.
  - 평문 JSON과 평문 엑셀은 저장소 및 서버에 존재하지 않으며, 접속 시 올바른 비밀번호를 입력해야만 브라우저 내 Web Crypto API로 복호화되어 화면과 다운로드가 열립니다.
  - `robots.txt` 및 `noindex` 메타 태그로 검색엔진 노출을 차단합니다.
- 로컬에서 새로운 데이터를 수집한 뒤 암호화 스냅샷(`static_data/*.enc.json`)을 갱신하려면 아래 명령을 실행하고 `main` 브랜치에 푸시합니다.
  ```powershell
  python -X utf8 build_gh_pages.py
  git add static_data/
  git commit -m "chore: update encrypted meat processor snapshot"
  git push origin main
  ```
