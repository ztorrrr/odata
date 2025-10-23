# CLAUDE.md

Claude Code (claude.ai/code) 사용 시 참고할 프로젝트 가이드 문서

## 프로젝트 개요 

BigQuery 데이터를 OData v4 REST API로 제공하는 서비스. Excel, Power BI 등 OData 호환 도구에서 데이터를 조회할 수 있음. GCS에 저장된 CSV 파일을 BigQuery로 로드하고 OData v4 쿼리 기능을 제공함.

## 아키텍처

### 인증 흐름
- GCP service account 자격증명은 AWS Secret Manager에 저장 (파일 저장 방식 사용 안 함)
- Secret key 형식: `{environment}/gen-ai/google/auth` (예: `dev/gen-ai/google/auth`)
- Secret은 JSON 객체 또는 `{"service_account_key": "<json-string>"}` 형태로 저장
- `app/utils/gcp_auth.py`에서 AWS Secret Manager로부터 자격증명을 가져와 인증 설정

### 데이터 흐름
1. CSV 파일이 GCS bucket에 저장
2. `app/services/data_loader.py`가 CSV를 BigQuery로 로드 (컬럼명 정제 포함)
3. FastAPI 서비스가 OData v4 REST API 제공
4. 클라이언트 도구(Excel, Power BI)가 OData protocol로 데이터 조회

### Excel 연결 방식 (4가지 옵션)

#### 1. Excel 템플릿 연결 방식 (추천 - 신규)
- 엔드포인트: `/odata/{table_name}/excel-template`
- **특징**: 템플릿 파일 기반, Power Query 연결만 포함 (데이터 미포함)
- 대용량 데이터 처리 최적화 (~수십 KB 파일 크기)
- 다운로드 후 Excel에서 "데이터 새로고침"만 하면 됨
- 필터/선택/정렬 조건이 연결에 포함되어 반복 사용 가능
- 사용 예: `curl "http://localhost:8888/odata/musinsa_data/excel-template?$filter=Media eq 'Naver'" -o data.xlsx`

#### 2. ODC 파일 방식
- 엔드포인트: `/odata/{table_name}/connection`
- 경량 파일 (~1KB) 다운로드 후 더블클릭
- Excel에서 OData 연결 자동 생성
- 크로스 플랫폼 (macOS/Linux 서버에서도 생성 가능)

#### 3. Excel 가이드 템플릿 방식
- 엔드포인트: `/odata/{table_name}/template`
- 사용 방법 안내 + Power Query M 코드 + 샘플 데이터 포함
- Excel 파일 다운로드 후 안내에 따라 연결 설정
- 초보자를 위한 상세 가이드 제공

#### 4. Windows COM 방식 (선택사항)
- 엔드포인트: `/odata/{table_name}/excel-live`
- 별도 Windows 서버 필요 (Windows Excel Service)
- Power Query가 내장된 실제 Excel 파일 생성
- 파일 열면 자동으로 데이터 로드
- 마이크로서비스 아키텍처: macOS 메인 서버 → Windows Excel Service (COM 자동화)
- 설정: `WINDOWS_EXCEL_SERVICE_URL` 환경 변수
- 자세한 내용: `WINDOWS_EXCEL_SERVICE_SETUP.md` 참조

### 주요 컴포넌트

**app/main.py**: FastAPI 애플리케이션 진입점 및 lifespan 관리
- 시작 시 GCP 인증 초기화
- BigQuery service singleton 설정
- Excel/브라우저 접근을 위한 CORS 설정

**app/routers/odata.py**: OData v4 endpoint 구현
- `/odata/` - Service document
- `/odata/$metadata` - Schema를 설명하는 XML metadata
- `/odata/{table_name}` - $filter, $select, $orderby, $top, $skip, $count를 지원하는 entity set 쿼리
- `/odata/{table_name}/$count` - 개수만 반환하는 endpoint
- `/odata/{table_name}/export` - CSV 파일 다운로드 endpoint
- `/odata/{table_name}/connection` - ODC (Office Data Connection) 파일 생성
- `/odata/{table_name}/template` - Excel 가이드 템플릿 파일 (사용 방법 안내 + Power Query M 코드 + 샘플 데이터)
- `/odata/{table_name}/excel-template` - **[신규]** 템플릿 기반 Excel 파일 (Power Query 연결만 포함, 데이터 미포함, 대용량 처리 최적화)
- `/odata/{table_name}/excel-live` - Windows Excel Service를 통한 실시간 OData 연결 Excel 파일 생성 (선택사항)

**app/services/bigquery_service.py**: BigQuery 작업 처리
- `get_bigquery_service()`로 singleton 인스턴스 접근
- 컬럼명 정제 (BOM 제거, 특수문자 처리, 숫자 시작 처리)
- `use_string_schema=True` 옵션 사용 시 모든 컬럼을 STRING 타입으로 로드하여 타입 추론 오류 방지
- OData 파라미터를 BigQuery SQL로 변환

**app/services/odata_query_parser.py**: OData 쿼리 변환
- OData 연산자(eq, ne, gt 등)를 SQL 연산자로 변환
- OData 함수(contains, startswith, endswith)를 SQL LIKE로 변환
- 예약어 충돌 방지를 위해 필드명에 backtick 추가

**app/services/odata_metadata.py**: OData metadata 생성
- BigQuery schema로부터 XML metadata document 생성
- BigQuery 타입을 EDM 타입으로 매핑 (STRING→Edm.String, INT64→Edm.Int64 등)
- 첫 번째 필드 또는 'id'/'key'/'code' 필드를 entity key로 자동 선택

**app/utils/setting.py**: 설정 관리
- 환경별 설정 (DEV/PROD/TEST)
- AWS Secret Manager에서 GCP service account key 조회
- `@cache` decorator로 singleton config 구현

**app/services/data_loader.py**: 데이터 로딩 스크립트
- GCS에서 BigQuery로 CSV 데이터 로드
- 환경 설정 로드 및 GCP 인증 초기화
- BigQuery dataset 자동 생성 (필요 시)
- 컬럼명 정제 및 STRING schema로 로드 (타입 오류 방지)
- 로드 완료 후 테이블 정보 및 샘플 데이터 출력

**app/services/excel_connection_modifier.py**: **[신규]** Excel 연결 수정 서비스
- Excel 파일(.xlsx)의 Power Query 연결 정보를 동적으로 수정
- Excel 파일을 ZIP으로 압축 해제 → XML 수정 → 재압축
- 수정 대상: `xl/connections.xml`, `xl/queries/*.xml`, `customXml/*.xml`
- 템플릿 파일(`app/template/odata_template.xlsx`)의 OData URL을 요청된 엔드포인트로 변경
- 대용량 데이터 처리에 최적화 (데이터 미포함, 연결 정보만 저장)

**app/utils/gcp_auth.py**: GCP 인증
- `get_gcp_auth()`로 singleton 인스턴스 접근
- credential 객체 및 임시 파일 인증 방식 모두 지원
- 적절한 자격증명이 설정된 BigQuery 및 Storage client 반환

**main.py**: 서버 실행 스크립트
- OData 서비스 실행을 위한 진입점
- 환경 설정 로드 및 서버 정보 출력
- uvicorn을 통해 FastAPI 앱 실행
- DEV 환경에서 auto-reload 활성화

**Windows Excel Service** (선택사항 - 별도 배포)
- Windows Server에서 독립적으로 실행되는 마이크로서비스
- Excel COM 자동화로 OData 연결이 내장된 Excel 파일 생성
- macOS 메인 서버의 `/odata/{table_name}/excel-live` 엔드포인트가 이 서비스에 요청 위임
- 구현 방법은 `WINDOWS_EXCEL_SERVICE_SETUP.md` 참조

## 주요 명령어

### 개발 환경 설정
```bash
# uv를 사용한 의존성 설치
pip install uv
uv sync
```

### GCS에서 BigQuery로 데이터 로드
```bash
uv run python -m app.services.data_loader
```
필요 시 BigQuery dataset 생성, 모든 컬럼을 STRING 타입으로 CSV 로드 (타입 오류 방지), 테이블 정보 출력

### OData 서버 실행
```bash
uv run python main.py
```
또는 직접 실행:
```bash
uv run uvicorn app.main:app --host 0.0.0.0 --port 8888 --reload
```

서버 endpoint (기본 포트: 8888):
- Service: `http://localhost:8888/odata/`
- Metadata: `http://localhost:8888/odata/$metadata`
- Data: `http://localhost:8888/odata/musinsa_data`
- Health: `http://localhost:8888/odata/health`
- CSV Export: `http://localhost:8888/odata/musinsa_data/export`
- **Excel Template (신규)**: `http://localhost:8888/odata/musinsa_data/excel-template`

### Excel 템플릿 연결 사용 예제 (신규 기능)
```bash
# 기본 사용
curl "http://localhost:8888/odata/musinsa_data/excel-template" -o data.xlsx

# 필터 적용 (Media가 'Naver'인 데이터만)
curl "http://localhost:8888/odata/musinsa_data/excel-template?\$filter=Media%20eq%20'Naver'" -o naver_data.xlsx

# 특정 필드만 선택
curl "http://localhost:8888/odata/musinsa_data/excel-template?\$select=Date,Campaign,Clicks" -o selected_fields.xlsx

# 정렬 조건 적용
curl "http://localhost:8888/odata/musinsa_data/excel-template?\$orderby=Date%20desc" -o sorted_data.xlsx

# 조합 사용
curl "http://localhost:8888/odata/musinsa_data/excel-template?\$filter=Media%20eq%20'Naver'&\$select=Date,Campaign&\$orderby=Date%20desc" -o filtered_sorted.xlsx
```

브라우저에서 직접 접속하여 다운로드:
```
http://localhost:8888/odata/musinsa_data/excel-template
http://localhost:8888/odata/musinsa_data/excel-template?$filter=Media eq 'Naver'
```

### 환경 설정
`.env.example`을 `.env`로 복사 후 설정:
- Secret Manager 접근용 AWS 자격증명
- GCP project ID
- GCS bucket 및 CSV 파일명
- BigQuery dataset 및 table명
- 서버 host 및 port

## 주요 구현 세부사항

### 컬럼명 정제
BigQuery는 엄격한 컬럼 명명 규칙을 적용함. `bigquery_service.py:50-92`의 `_sanitize_column_name` 메서드:
- BOM 문자 (UTF-8 byte order mark) 제거
- 공백 및 특수문자를 underscore로 변환
- 숫자로 시작하는 경우 `col_` prefix 추가
- 300자로 truncate

### Schema 로드 전략
데이터 품질 문제가 있는 CSV의 경우 `use_string_schema=True` 사용 권장. 모든 컬럼을 STRING 타입으로 생성하여 타입 추론 오류를 방지함.

### OData Pagination
결과 개수가 `$top` 값과 일치하면 자동으로 `@odata.nextLink` 추가. Excel에서 대용량 데이터를 점진적으로 로드 가능. 기본 page size는 1000 (`ODATA_MAX_PAGE_SIZE`로 설정 가능).

### Singleton Pattern
`BigQueryService`와 `GCPAuth` 모두 singleton pattern 사용하여 재인증 및 재초기화 방지. 각각 `get_bigquery_service()` 및 `get_gcp_auth()`로 접근.

### CSV Export
`/odata/{table_name}/export` endpoint를 통해 데이터를 CSV 파일로 다운로드 가능:
- OData 쿼리 파라미터 재사용 ($filter, $select, $orderby, $top, $skip)
- pandas DataFrame으로 변환 후 CSV 생성
- UTF-8 BOM 추가로 Excel 호환성 확보 (한글 깨짐 방지)
- 파일명 자동 생성: `{table_name}_{timestamp}.csv` 형식
- 기본 최대 행 수: 100,000 (최대 1,000,000까지 설정 가능)
- 사용 예: `curl "http://localhost:8888/odata/musinsa_data/export?$filter=Media eq 'Naver'&$top=1000" -o data.csv`

### Excel 템플릿 연결 수정 (신규)
`/odata/{table_name}/excel-template` endpoint를 통해 템플릿 기반 Excel 파일 생성:
- **동작 원리**: Excel 파일(.xlsx)은 ZIP 압축된 XML 파일들의 모음
  1. 템플릿 파일(`app/template/odata_template.xlsx`)을 ZIP으로 압축 해제
  2. Power Query 연결 정보가 담긴 XML 파일들 수정
     - `xl/connections.xml`: 연결 문자열의 Location URL
     - `xl/queries/*.xml`: Power Query M 코드의 OData.Feed URL
     - `customXml/*.xml`: 커스텀 연결 정보 (있는 경우)
  3. 수정된 파일들을 다시 ZIP으로 압축하여 .xlsx 생성
- **장점**:
  - 데이터를 파일에 포함하지 않아 파일 크기 최소화 (~수십 KB)
  - 대용량 데이터 처리 가능 (Excel에서 필요한 만큼만 로드)
  - 필터/선택/정렬 조건이 연결에 포함되어 반복 사용 가능
  - 사용자는 다운로드 후 Excel에서 "데이터 새로고침"만 하면 됨
- 파일명 자동 생성: `{table_name}_connection_{timestamp}.xlsx` 형식
- 임시 파일은 자동으로 삭제됨 (FileResponse background 처리)

## 프로젝트 의존성

주요 라이브러리 (pyproject.toml 참조):
- FastAPI + uvicorn: 웹 서비스
- google-cloud-bigquery, google-cloud-storage: GCP 연동
- boto3: AWS Secret Manager
- pandas: 데이터 처리 (CSV 로딩 및 export 시 사용)
- lxml: XML metadata 생성
- cachetools: Secret Manager 응답 캐싱
- openpyxl: Excel 템플릿 파일 생성
- httpx: Windows Excel Service 연동용 HTTP 클라이언트 (선택사항)

## 트러블슈팅

BigQuery table을 찾을 수 없는 경우: `uv run python -m app.services.data_loader` 먼저 실행

GCP 인증 실패 시: AWS Secret Manager의 `{environment}/gen-ai/google/auth`에 GCP service account key 존재 여부 확인

Excel 연결 실패 시: 서버 실행 상태 확인 및 OData feed URL로 `http://localhost:8888/odata` 사용 (기본 포트 8888)
