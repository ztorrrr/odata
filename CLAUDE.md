# CLAUDE.md

Claude Code (claude.ai/code) 사용 시 참고할 프로젝트 가이드 문서

## 프로젝트 개요

BigQuery 데이터를 OData v4 REST API로 제공하는 서비스. Excel, Power BI 등 OData 호환 도구에서 데이터를 조회할 수 있음.

## 아키텍처

### GCP 인증 흐름
- **우선순위 1**: Application Default Credentials (ADC) - 로컬 개발 및 Google Sheets/Drive API 사용
- **우선순위 2**: AWS Secret Manager의 Service Account 키 - Fallback
- Secret key 형식: `{environment}/gen-ai/google/auth` (예: `dev/gen-ai/google/auth`)
- `app/utils/gcp_auth.py`에서 인증 처리
- ADC 설정: `gcloud auth application-default login --scopes=...` (Sheets, Drive, BigQuery 스코프 필요)

### OData API 사용자 인증
- HTTP Basic Authentication 사용 (Excel, Power BI 등과 호환)
- 사용자 계정 정보는 AWS Secret Manager에 JSON 형식으로 저장
- Secret key 형식: `{environment}/gen-ai/odata/users`
- 데이터 조회 엔드포인트에만 인증 필요 (metadata는 인증 불필요)
- FastAPI Dependency Injection으로 구현 (`app/utils/auth.py`)

### 데이터 흐름
1. CSV 파일이 GCS bucket에 저장
2. `app/services/data_loader.py`가 CSV를 BigQuery로 로드
3. FastAPI 서비스가 OData v4 REST API 제공
4. 클라이언트 도구(Excel, Power BI)가 OData protocol로 데이터 조회

## 주요 컴포넌트

### Core Files

**main.py**: 서버 실행 스크립트
- uvicorn을 통해 FastAPI 앱 실행
- 환경 설정 로드 및 서버 정보 출력
- DEV 환경에서 auto-reload 활성화

**app/main.py**: FastAPI 애플리케이션
- lifespan 관리 (시작/종료 시 리소스 관리)
- GCP 인증 초기화
- BigQuery service singleton 설정
- CORS 설정 (Excel/브라우저 접근용)

### Routers

**app/routers/odata.py**: OData v4 엔드포인트
- `/odata/` - Service document
- `/odata/$metadata` - XML metadata
- `/odata/{table_name}` - Entity set 쿼리 ($filter, $select, $orderby, $top, $skip, $count 지원)
- `/odata/{table_name}/$count` - 개수 조회
- `/odata/{table_name}/export` - CSV 내보내기
- `/odata/{table_name}/excel-com` - Windows COM 기반 Excel 파일 생성

**app/routers/spreadsheet.py**: Google Spreadsheet 연동 엔드포인트
- `/spreadsheet/create-connected-bigquery` - BigQuery Connected Sheets 자동 생성 (네이티브 연결)
- `/spreadsheet/create-sample-view` - BigQuery 샘플 View 생성
- `/spreadsheet/sample-data` - 데이터 미리보기
- `/spreadsheet/modify-view-test` - 테스트용 View 수정
- `/spreadsheet/restore-view` - View 복원

### Services

**app/services/bigquery_service.py**: BigQuery 작업 처리
- Singleton pattern (`get_bigquery_service()`)
- 컬럼명 정제 (특수문자, BOM 제거)
- OData 파라미터를 BigQuery SQL로 변환

**app/services/odata_query_parser.py**: OData 쿼리 파싱
- OData 연산자를 SQL로 변환 (eq, ne, gt, lt, contains 등)
- 필드명에 backtick 추가로 예약어 충돌 방지

**app/services/odata_metadata.py**: OData metadata 생성
- BigQuery schema를 EDM 타입으로 매핑
- XML metadata document 생성

**app/services/excel_com_generator.py**: Windows COM Excel 생성
- Windows COM 자동화로 Excel 파일 생성
- OData 연결 정보 포함
- Power Query 설정 시도 (실패 시 연결 정보만 텍스트로 제공)

**app/services/data_loader.py**: 데이터 로딩
- GCS에서 BigQuery로 CSV 로드
- 모든 컬럼을 STRING 타입으로 로드 (타입 오류 방지)
- BigQuery dataset 자동 생성

**app/services/spreadsheet_connector.py**: Google Spreadsheet 연동
- BigQuery Connected Sheets 네이티브 연결 생성
- 자동 폴더 검색 및 파일명 생성 (bigquery_connector_YYMMDD_HHMMSS)
- BigQuery 샘플 View 생성/수정/복원
- Data Source API 활용으로 Apps Script 불필요

### Utils

**app/utils/setting.py**: 설정 관리
- 환경별 설정 (DEV/PROD/TEST)
- AWS Secret Manager 연동
- `@cache` decorator로 singleton 구현

**app/utils/gcp_auth.py**: GCP 인증
- `get_gcp_auth()`로 singleton 접근
- BigQuery 및 Storage client 생성

**app/utils/auth.py**: OData API 사용자 인증
- HTTP Basic Authentication 구현
- AWS Secret Manager에서 사용자 정보 로드
- FastAPI Dependency로 제공 (`get_current_user()`)
- 타이밍 공격 방지 (`secrets.compare_digest()` 사용)

## 주요 명령어

### 개발 환경 설정
```bash
# 의존성 설치
pip install uv
uv sync
```

### 데이터 로드
```bash
uv run python -m app.services.data_loader
```

### 서버 실행
```bash
uv run python main.py
# 또는
uv run uvicorn app.main:app --host 0.0.0.0 --port 8888 --reload
```

### Windows COM Excel 생성 (직접 실행)
```python
from app.services.excel_com_generator import create_excel_with_odata_com
output_path = create_excel_with_odata_com('http://localhost:8888/odata/musinsa_data', 'Data')
```

## 엔드포인트

기본 포트: 8888

### OData v4 엔드포인트
- Service Document: `http://localhost:8888/odata/`
- Metadata: `http://localhost:8888/odata/$metadata`
- Data: `http://localhost:8888/odata/musinsa_data`
- Count: `http://localhost:8888/odata/musinsa_data/$count`
- CSV Export: `http://localhost:8888/odata/musinsa_data/export`
- Excel COM: `http://localhost:8888/odata/musinsa_data/excel-com`
- Health: `http://localhost:8888/odata/health`

### Google Spreadsheet 엔드포인트
- Connected Sheets 생성: `GET/POST http://localhost:8888/spreadsheet/create-connected-bigquery`
- 샘플 View 생성: `POST http://localhost:8888/spreadsheet/create-sample-view`
- 데이터 미리보기: `GET http://localhost:8888/spreadsheet/sample-data`

## 환경 설정

`.env` 파일 필요:
```
ENVIRONMENT=DEV
AWS_DEFAULT_REGION=ap-northeast-2
GCP_PROJECT_ID=your-project
GCS_BUCKET_NAME=your-bucket
GCS_FILE_NAME=data.csv
BIGQUERY_DATASET_ID=odata_dataset
BIGQUERY_TABLE_NAME=musinsa_data
```

## 구현 세부사항

### API 인증
HTTP Basic Authentication으로 구현:
- **인증 필요 엔드포인트**: 데이터 조회, CSV export, Excel 생성
- **인증 불필요 엔드포인트**: service document, metadata, health check

AWS Secret Manager 설정:
- Secret 이름: `dev/gen-ai/odata/users` (환경별로 다름)
- JSON 형식:
```json
{
  "users": [
    {"username": "user1", "password": "pass1"},
    {"username": "user2", "password": "pass2"}
  ]
}
```

Excel에서 사용:
1. 데이터 > OData 피드 연결
2. 인증 방식: "기본" (Basic) 선택
3. 사용자명/암호 입력

개발 환경:
- Secret Manager에 사용자 정보가 없으면 인증 우회 (DEV 모드만)
- 로그에 경고 메시지 출력

### 컬럼명 정제
BigQuery 컬럼 명명 규칙 적용:
- BOM 문자 제거
- 특수문자를 underscore로 변환
- 숫자로 시작하면 `col_` prefix 추가
- 300자로 제한

### OData Pagination
- 기본 page size: 1000 (`ODATA_MAX_PAGE_SIZE`)
- 결과가 page size와 일치하면 `@odata.nextLink` 자동 추가

### CSV Export
- UTF-8 BOM 추가로 Excel 호환성 확보
- 최대 100,000행 (설정 가능)

### Windows COM Excel 생성
- Windows 환경에서만 작동
- Excel 설치 필요
- Power Query 추가 시도 (실패 시 연결 정보만 제공)

### BigQuery Connected Sheets
- Google Sheets API v4의 Data Source 기능 사용
- 네이티브 BigQuery 연결 (Apps Script 불필요)
- 자동 기능:
  - 파일명 자동 생성: `bigquery_connector_YYMMDD_HHMMSS`
  - "odata_test" 폴더 자동 검색 및 저장
  - 기본 빈 시트 자동 제거
- 사용자는 스프레드시트에서 "데이터 > 새로고침"으로 최신 데이터 조회
- GET 요청 지원으로 브라우저에서 직접 호출 가능 (`?token=xxx`)

**사용 예시:**
```bash
# 브라우저에서 (GET)
http://localhost:8888/spreadsheet/create-connected-bigquery?token=test-token

# API 호출 (POST)
curl -X POST "http://localhost:8888/spreadsheet/create-connected-bigquery" \
  -H "Authorization: Bearer test-token"

# 파라미터 없이 호출하면 모든 것이 자동:
# - 파일명: bigquery_connector_251029_131530 (현재 시각)
# - 폴더: odata_test (자동 검색)
# - 데이터: 기본 샘플 View (100행)
```

## 프로젝트 의존성

주요 라이브러리:
- FastAPI + uvicorn
- google-cloud-bigquery, google-cloud-storage
- google-api-python-client (Sheets, Drive API)
- google-auth, google-auth-httplib2, google-auth-oauthlib
- boto3 (AWS Secret Manager)
- pandas (데이터 처리)
- lxml (XML metadata)
- pywin32 (Windows COM)