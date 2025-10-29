# Google Spreadsheet BigQuery Connector

BigQuery 데이터를 Google Spreadsheet와 실시간 연동하는 기능

## 개요

이 기능은 BigQuery 테이블의 샘플 데이터를 View로 생성하고, Google Sheets의 Connected Sheets 기능을 통해 스프레드시트와 연동할 수 있도록 지원합니다.

## 주요 기능

- BigQuery 샘플 View 자동 생성 (기본 100개 행)
- Google Sheets 연동 가이드 제공
- 데이터 미리보기
- RESTful API 엔드포인트 제공

## API 엔드포인트

### 1. 전체 설정 프로세스
```
POST /spreadsheet/setup-spreadsheet/{spreadsheet_id}
```
- 샘플 View 생성과 연동 가이드를 한번에 제공
- 파라미터:
  - `spreadsheet_id`: Google Spreadsheet ID (필수)
  - `sample_size`: 샘플 데이터 행 수 (기본: 100)
  - `source_table`: 원본 테이블 이름 (옵션)

### 2. 샘플 View 생성
```
POST /spreadsheet/create-sample-view
```
- BigQuery에 샘플 데이터 View 생성
- 파라미터:
  - `sample_size`: 샘플 행 수 (기본: 100)
  - `source_table`: 원본 테이블 이름 (옵션)
  - `view_name`: View 이름 (옵션, 자동 생성)
  - `force_recreate`: 기존 View 재생성 여부 (기본: false)

### 3. 연결 설정 정보
```
GET /spreadsheet/connection-config
```
- Connected Sheets 연결에 필요한 설정 정보 반환
- 파라미터:
  - `spreadsheet_id`: Google Spreadsheet ID (필수)
  - `view_id`: BigQuery View ID (옵션)

### 4. 설정 가이드
```
GET /spreadsheet/setup-guide
```
- 수동 설정 가이드 제공
- 파라미터:
  - `spreadsheet_id`: Google Spreadsheet ID (필수)
  - `view_id`: BigQuery View ID (옵션)

### 5. 데이터 미리보기
```
GET /spreadsheet/sample-data
```
- 샘플 View의 데이터 미리보기
- 파라미터:
  - `view_id`: BigQuery View ID (옵션)
  - `limit`: 미리보기 행 수 (기본: 10)

## 인증

모든 엔드포인트는 인증이 필요합니다. 다음 방식 중 하나를 사용:

### Bearer Token (권장)
```bash
curl -H "Authorization: Bearer YOUR_TOKEN" http://localhost:8888/spreadsheet/...
```

### HTTP Basic Auth
```bash
curl -u username:password http://localhost:8888/spreadsheet/...
```

### Query Parameter Token
```bash
curl "http://localhost:8888/spreadsheet/...?token=YOUR_TOKEN"
```

## 사용 예시

### 1. 스프레드시트 연동 설정
```bash
# Bearer 토큰 사용
curl -X POST "http://localhost:8888/spreadsheet/setup-spreadsheet/YOUR_SPREADSHEET_ID?sample_size=100" \
  -H "Authorization: Bearer your_token"

# Basic Auth 사용
curl -X POST "http://localhost:8888/spreadsheet/setup-spreadsheet/YOUR_SPREADSHEET_ID?sample_size=100" \
  -u username:password
```

### 2. 샘플 데이터 확인
```bash
curl "http://localhost:8888/spreadsheet/sample-data?limit=5" \
  -H "Authorization: Bearer your_token"
```

## Google Sheets에서 연결하기

### 방법 1: BigQuery 데이터 커넥터 사용

1. Google Sheets 열기
2. **데이터 > 데이터 커넥터 > BigQuery에 연결** 메뉴 선택
3. Google Cloud 프로젝트 선택
4. 데이터셋 찾기 (예: `odata_dataset`)
5. 생성된 View 선택 (예: `musinsa_data_sample_100`)
6. "연결" 클릭
7. 가져올 열 선택
8. 새로고침 일정 설정 (선택사항)

### 방법 2: 사용자 지정 쿼리 사용

1. Google Sheets에서 **데이터 > 데이터 커넥터 > BigQuery에 연결**
2. "사용자 지정 쿼리" 옵션 선택
3. SQL 쿼리 입력:
   ```sql
   SELECT * FROM `project.dataset.view_name`
   ```
4. "연결" 클릭

## Connected Sheets 특징

- **실시간 동기화**: BigQuery 데이터 변경 시 수동/자동 새로고침 가능
- **대용량 데이터 처리**: 최대 10,000개 행까지 직접 표시
- **보안**: Google Cloud IAM 권한 기반 접근 제어
- **협업**: 여러 사용자가 동일한 데이터 소스 공유 가능

## 제한사항

- Google Workspace 계정 필요 (Connected Sheets 사용)
- BigQuery 접근 권한 필요
- 샘플 View는 정적 데이터 (원본 테이블 변경 시 재생성 필요)

## 개발 환경 설정

### 환경 변수
```env
ENVIRONMENT=DEV
GCP_PROJECT_ID=your-project-id
BIGQUERY_DATASET_ID=odata_dataset
BIGQUERY_TABLE_NAME=musinsa_data
```

### DEV 환경 특징
- Bearer 토큰 인증 시 모든 토큰 허용
- Secret Manager 설정 없이 인증 우회 가능

## 문제 해결

### View 생성 실패
- BigQuery 데이터셋 존재 확인
- GCP 권한 확인 (BigQuery Data Editor 이상)
- 원본 테이블 존재 확인

### 스프레드시트 연결 실패
- Google Cloud 프로젝트 접근 권한 확인
- BigQuery API 활성화 확인
- View ID가 올바른지 확인

### 인증 오류
- DEV 환경: Bearer 토큰 사용 권장
- PROD 환경: AWS Secret Manager에 사용자 정보 설정 필요