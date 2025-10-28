# Google Spreadsheet BigQuery 연동 기능 구현

날짜: 2025-10-28

## 작업 개요
BigQuery 테이블 데이터를 Google Spreadsheet와 연동하는 기능을 구현했습니다. 기존 OData 서비스에 새로운 엔드포인트를 추가하여 BigQuery View를 생성하고 Google Sheets의 Connected Sheets 기능과 연결할 수 있도록 했습니다.

## 브랜치 정보
- 기존 브랜치: feature/webapi-test
- 새 브랜치: feature/spreadsheet-connector

## 구현 내용

### 1. SpreadsheetConnector 서비스
파일: `app/services/spreadsheet_connector.py`
- BigQuery 샘플 View 생성 (100개 행 제한)
- Connected Sheets 연결 설정 정보 제공
- View 수정 기능 (컬럼에 suffix 추가)
- View 복원 기능

### 2. API 엔드포인트 추가
파일: `app/routers/spreadsheet.py`
- POST `/spreadsheet/setup-spreadsheet/{spreadsheet_id}` - 전체 설정 프로세스
- POST `/spreadsheet/create-sample-view` - 샘플 View 생성
- GET `/spreadsheet/connection-config` - 연결 설정 정보
- GET `/spreadsheet/setup-guide` - 설정 가이드
- GET `/spreadsheet/sample-data` - 데이터 미리보기
- POST `/spreadsheet/modify-view-test` - View 수정 (테스트용)
- POST `/spreadsheet/restore-view` - View 원상복구

### 3. 인증 개선
파일: `app/routers/spreadsheet.py`
- 기존 `get_current_user`에서 `get_current_user_with_header_token`으로 변경
- Bearer 토큰 지원 추가 (DEV 환경에서 모든 토큰 허용)

### 4. 메인 앱 수정
파일: `app/main.py`
- spreadsheet 라우터 추가
- 루트 엔드포인트에 spreadsheet_endpoint 정보 추가

## 권한 설정
Google Sheets에서 BigQuery 프로젝트 접근을 위해 IAM 권한 부여:
```bash
gcloud projects add-iam-policy-binding dataconsulting-imagen2-test \
    --member="user:jckang@madup.com" \
    --role="roles/bigquery.user"
```

## 테스트
### 테스트 시트
- URL: https://docs.google.com/spreadsheets/d/14v8oM27b8WN5gQFWLf-VvdCDyN2FJY3yqf7J8Zyn_vY

### 생성된 View
- View ID: `dataconsulting-imagen2-test.odata_dataset.musinsa_data_sample_100`
- 샘플 데이터: 100개 행

### View 수정 테스트
Type 컬럼에 '_테스트' suffix를 추가하여 Google Sheets 실시간 동기화 확인:
- 수정 전: "검색"
- 수정 후: "검색_테스트"

## 테스트 스크립트
- `test_spreadsheet_connector.py` - 전체 연동 프로세스 테스트
- `test_view_modification.py` - View 수정 및 복원 테스트

## 문서
- `SPREADSHEET_CONNECTOR.md` - 기능 설명 및 사용 가이드

## 주요 발견사항
1. Google Sheets Connected Sheets는 자동 실시간 동기화를 지원하지 않음
2. 데이터 변경 후 수동 새로고침 필요
3. BigQuery View 수정이 즉시 반영되나 Sheets에서는 새로고침 필요

## 미해결 과제
- Google Sheets 자동 새로고침 구현 (Apps Script 활용 고려)
- 대용량 데이터 처리 최적화
- View 버전 관리 기능