# OData v4 Service for BigQuery

BigQuery 데이터를 OData v4 REST API로 제공하는 서비스.

## 주요 기능

- **OData v4 API**: BigQuery 데이터를 OData 프로토콜로 제공
- **Windows COM Excel**: OData 연결 정보가 포함된 Excel 파일 생성
- **CSV Export**: 쿼리 결과를 CSV로 내보내기
- **자동 캐시 정리**: 서버 시작 시 Python 바이트코드 캐시 자동 정리

## 시작하기

### 1. 의존성 설치

```bash
pip install uv
uv sync
```

### 2. 환경 설정

`.env` 파일을 생성하고 다음 내용 입력:

```env
ENVIRONMENT=DEV
AWS_DEFAULT_REGION=ap-northeast-2
GCP_PROJECT_ID=your-project-id
GCS_BUCKET_NAME=your-bucket
GCS_FILE_NAME=data.csv
BIGQUERY_DATASET_ID=odata_dataset
BIGQUERY_TABLE_NAME=your_table
```

### 3. 데이터 로드 (최초 1회)

```bash
uv run python -m app.services.data_loader
```

### 4. 서버 실행

```bash
uv run python main.py
```

서버가 실행되면:
- **포트**: 8888
- **OData 엔드포인트**: http://localhost:8888/odata/
- **Health Check**: http://localhost:8888/odata/health

## API 엔드포인트

### OData 표준 엔드포인트

- `GET /odata/` - Service document
- `GET /odata/$metadata` - Schema metadata (XML)
- `GET /odata/{table_name}` - 데이터 조회
- `GET /odata/{table_name}/$count` - 레코드 개수

### 쿼리 예시

```bash
# 필터링
curl "http://localhost:8888/odata/musinsa_data?\$filter=Media eq 'Naver'"

# 필드 선택
curl "http://localhost:8888/odata/musinsa_data?\$select=Date,Campaign,Clicks"

# 정렬
curl "http://localhost:8888/odata/musinsa_data?\$orderby=Date desc"

# 페이징
curl "http://localhost:8888/odata/musinsa_data?\$top=100&\$skip=0"

# 개수 포함
curl "http://localhost:8888/odata/musinsa_data?\$count=true"
```

### 추가 엔드포인트

- `GET /odata/{table_name}/export` - CSV 내보내기
- `GET /odata/{table_name}/excel-com` - Windows COM 기반 Excel 파일 생성

## Excel 연결

### Windows COM 방식

```bash
curl "http://localhost:8888/odata/musinsa_data/excel-com" -o data.xlsx
```

다운로드한 Excel 파일 열기:
1. OData 연결 정보가 포함되어 있음
2. 데이터 탭 → 모두 새로고침 클릭
3. 실시간 데이터 로드


## 프로젝트 구조

```
odata/
├── main.py                          # 서버 실행 (캐시 정리 포함)
├── app/
│   ├── main.py                      # FastAPI 앱
│   ├── routers/odata.py             # OData 엔드포인트
│   ├── services/
│   │   ├── bigquery_service.py      # BigQuery 처리
│   │   ├── odata_*.py               # OData 관련 서비스
│   │   ├── excel_com_generator.py   # COM Excel 생성
│   │   └── data_loader.py           # 데이터 로더
│   └── utils/
│       ├── setting.py               # 설정 관리
│       └── gcp_auth.py              # GCP 인증
├── .env                             # 환경 변수
├── pyproject.toml                   # 의존성 및 설정
└── CLAUDE.md                        # 개발 가이드
```

## 사용 인프라

- **FastAPI**: 웹 프레임워크
- **uvicorn**: ASGI 서버
- **Google Cloud BigQuery**: 데이터 저장소
- **AWS Secret Manager**: 자격증명 관리
- **pywin32**: Windows COM 자동화
- **pandas**: 데이터 처리
- **lxml**: XML 메타데이터 생성