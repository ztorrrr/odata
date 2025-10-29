# BigQuery 데이터 연결 Google Spreadsheet 자동 생성 기능 구현

**작업 날짜**: 2025-10-29
**목표**: GCS → BigQuery로 변환된 데이터를 Google Spreadsheet로 자동 생성하고, BigQuery 데이터 커넥터를 통해 연결

## 프로젝트 개요

- **GCP 프로젝트**: `dataconsulting-imagen2-test`
- **BigQuery 데이터셋**: `odata_dataset`
- **테스트 테이블**: `musinsa_data_sample_100` (100개 행 샘플 뷰)
- **타겟 폴더**: https://drive.google.com/drive/folders/1HJdKCg9RBs0ky79QfsUA66r0615pJy78
- **인증 계정**: `dc_team@madup.com` (Application Default Credentials 사용)
- **서비스 계정**: `gen-ai@dataconsulting-imagen2-test.iam.gserviceaccount.com` (권한 부족으로 ADC 사용)

## 구현 완료 사항

### 1. 라이브러리 추가 (`pyproject.toml`)

```toml
dependencies = [
    # ... 기존 라이브러리들 ...
    "google-api-python-client>=2.149.0",
    "google-auth>=2.35.0",
    "google-auth-httplib2>=0.2.0",
    "google-auth-oauthlib>=1.2.1",
    # ...
]
```

**설치 완료**: `uv sync` 실행됨

### 2. GCP 인증 스코프 추가 (`app/utils/gcp_auth.py`)

#### 변경된 스코프 목록:
```python
scopes=[
    "https://www.googleapis.com/auth/cloud-platform",
    "https://www.googleapis.com/auth/bigquery",
    "https://www.googleapis.com/auth/devstorage.read_only",
    "https://www.googleapis.com/auth/spreadsheets",      # 추가
    "https://www.googleapis.com/auth/drive",             # 추가
    "https://www.googleapis.com/auth/script.projects",   # 추가
]
```

#### ADC (Application Default Credentials) 우선 사용 로직 추가:

**신규 메서드**: `GCPAuth.authenticate_with_adc()`
```python
def authenticate_with_adc(self) -> bool:
    """
    Application Default Credentials (ADC)로 인증을 시도합니다.

    Returns:
        인증 성공 여부
    """
    try:
        credentials, project = default(scopes=[...])
        self.credentials = credentials
        self.project_id = project or os.environ.get("GOOGLE_CLOUD_PROJECT")
        self._initialized = True
        logger.info(f"ADC authentication successful. Project: {self.project_id}")
        return True
    except DefaultCredentialsError as e:
        logger.warning(f"ADC not available: {e}")
        return False
```

#### 인증 우선순위 변경 (`app/utils/setting.py`):

```python
def setup_gcp_auth():
    """
    GCP 인증을 설정합니다.
    1. Application Default Credentials (ADC) 우선 시도
    2. 실패 시 AWS Secret Manager에서 서비스 계정 키를 가져와 설정
    """
    gcp_auth = get_gcp_auth()

    # 1. ADC로 인증 시도
    if gcp_auth.authenticate_with_adc():
        logger.info(f"Authenticated with Application Default Credentials")
        return gcp_auth

    # 2. ADC 실패 시 AWS Secret Manager 사용
    logger.info("ADC not available, using service account from AWS Secret Manager")
    gcp_auth.authenticate_from_secret(get_config().GCP_SERVICE_ACCOUNT_KEY)
    return gcp_auth
```

### 3. SpreadsheetConnector 서비스 확장 (`app/services/spreadsheet_connector.py`)

#### 신규 메서드들:

##### `_get_sheets_service()`, `_get_drive_service()`, `_get_script_service()`
Google API 클라이언트 생성

##### `create_spreadsheet_with_bigquery()`
**메인 함수 - 스프레드시트 자동 생성**

파라미터:
- `spreadsheet_title`: 생성할 스프레드시트 이름
- `view_id`: BigQuery View ID (None이면 기본 샘플 view 사용)
- `folder_id`: Google Drive 폴더 ID (None이면 루트에 생성)
- `include_apps_script`: Apps Script 코드 포함 여부

동작 흐름:
1. 스프레드시트 생성 (`_create_spreadsheet`)
2. 지정된 폴더로 이동 (`_move_to_folder`)
3. BigQuery 데이터 로드 (`_load_bigquery_data_to_sheet`)
4. Apps Script 코드 생성 (`_add_bigquery_apps_script`)
5. 설정 가이드 시트 추가 (`_add_guide_sheet`)

##### `_create_spreadsheet(title: str)`
빈 스프레드시트 생성 (헤더 행 고정 포함)

##### `_move_to_folder(spreadsheet_id: str, folder_id: str)`
스프레드시트를 특정 Google Drive 폴더로 이동

##### `_load_bigquery_data_to_sheet(spreadsheet_id: str, view_id: str)`
BigQuery에서 데이터 조회 후 시트에 작성 (최대 1000행)
- 헤더 행 자동 포맷팅 (파란색 배경, 흰색 볼드 텍스트)

##### `_format_header_row(spreadsheet_id: str, sheet_id: int)`
헤더 행 스타일 적용

##### `_add_bigquery_apps_script(...)`
BigQuery 데이터 새로고침을 위한 Apps Script 코드 생성

**생성되는 스크립트 기능**:
- `refreshBigQueryData()`: BigQuery API를 통해 데이터 조회 및 시트 업데이트
- `onOpen()`: 스프레드시트 열 때 "BigQuery" 메뉴 자동 생성

**사용 방법**:
1. 확장 프로그램 > Apps Script
2. 제공된 코드 붙여넣기
3. BigQuery API 서비스 추가
4. 저장 후 "BigQuery > 데이터 새로고침" 메뉴 사용

##### `_add_guide_sheet(...)`
설정 가이드를 포함한 시트 생성
- 현재 연결 정보 표시
- 데이터 새로고침 방법 안내 (Apps Script / Connected Sheets)
- 참고사항

##### `_add_apps_script_code_sheet(...)`
Apps Script 코드를 별도 시트에 작성 (복사 편의성)

### 4. API 엔드포인트 추가 (`app/routers/spreadsheet.py`)

#### 신규 엔드포인트: `POST /spreadsheet/create-with-bigquery`

**Query Parameters**:
- `title` (required): 스프레드시트 제목
- `view_id` (optional): BigQuery View ID
- `folder_id` (optional): Google Drive 폴더 ID
- `include_apps_script` (optional, default=True): Apps Script 포함 여부

**Response**:
```json
{
  "success": true,
  "spreadsheet_id": "...",
  "spreadsheet_url": "https://docs.google.com/spreadsheets/d/...",
  "view_id": "project.dataset.view",
  "folder_id": "...",
  "apps_script_added": true,
  "script_info": { ... },
  "message": "Spreadsheet 'title' created successfully with BigQuery data"
}
```

**인증**: HTTP Bearer Token 필요 (`Authorization: Bearer test-token`)

## 현재 상태 및 이슈

### ✅ 해결된 문제들

1. **모듈 Import 에러** → `uv sync`로 해결
2. **서비스 계정 권한 부족** → ADC 사용으로 전환
3. **ADC 미인증 상태** → `gcloud auth application-default login` 실행
4. **ADC Scopes 부족** → 올바른 scopes로 재인증 완료

### ⚠️ 현재 이슈

**문제**: 현재 터미널 세션에 ADC 변경사항이 반영되지 않음

**원인**:
- ADC는 `%APPDATA%\gcloud\application_default_credentials.json` 파일로 관리됨
- 이미 실행 중인 프로세스는 이전 인증 정보를 캐싱하고 있을 수 있음

**해결 방법**:
1. **서버 완전히 종료** (현재 실행 중인 모든 Python 프로세스)
2. **새 터미널 세션에서 서버 재시작**
3. 서버 시작 로그에서 다음 확인:
   ```
   Authenticated with Application Default Credentials (project: ...)
   ```

### 🔍 인증 확인 테스트 결과

**테스트 스크립트**: `test_sheets_api.py`

**새 세션 테스트 결과**:
```
Project ID: None
Credentials type: <class 'google.oauth2.credentials.Credentials'>  ← ADC 작동 확인!
[OK] Google Sheets API service created successfully

[ERROR] Request had insufficient authentication scopes  ← scopes 문제 (해결됨)
```

**scopes 재설정 후 예상 결과**: 정상 작동

## 다음 단계 (새 세션에서 수행)

### Step 1: 서버 재시작

```bash
# 기존 서버 완전 종료 후
uv run python main.py
```

**확인할 로그**:
```
2025-10-29 XX:XX:XX - app.utils.gcp_auth - INFO - ADC authentication successful. Project: dataconsulting-imagen2-test
2025-10-29 XX:XX:XX - app.utils.setting - INFO - Authenticated with Application Default Credentials (project: dataconsulting-imagen2-test)
```

### Step 2: 테스트 실행

#### 테스트 1: 폴더 지정 없이 생성 (루트)
```bash
curl -X POST "http://localhost:8888/spreadsheet/create-with-bigquery?title=odata_test_no_folder" \
  -H "Authorization: Bearer test-token" \
  -H "Content-Type: application/json"
```

**예상 결과**: 스프레드시트 생성 성공, URL 반환

#### 테스트 2: 특정 폴더에 생성
```bash
curl -X POST "http://localhost:8888/spreadsheet/create-with-bigquery?title=odata_test&folder_id=1HJdKCg9RBs0ky79QfsUA66r0615pJy78" \
  -H "Authorization: Bearer test-token" \
  -H "Content-Type: application/json"
```

**타겟 폴더**: https://drive.google.com/drive/folders/1HJdKCg9RBs0ky79QfsUA66r0615pJy78

#### 테스트 3: 커스텀 View 사용
```bash
curl -X POST "http://localhost:8888/spreadsheet/create-with-bigquery?title=custom_test&view_id=dataconsulting-imagen2-test.odata_dataset.musinsa_data_sample_100&folder_id=1HJdKCg9RBs0ky79QfsUA66r0615pJy78" \
  -H "Authorization: Bearer test-token" \
  -H "Content-Type: application/json"
```

### Step 3: 생성된 스프레드시트 확인

생성 성공 시 응답에서 `spreadsheet_url`을 받아 브라우저에서 확인:

**확인 사항**:
1. ✅ "Data" 시트에 BigQuery 데이터 로드됨
2. ✅ "Setup Guide" 시트에 설정 가이드 작성됨
3. ✅ "Apps Script Code" 시트에 스크립트 코드 작성됨
4. ✅ 헤더 행이 파란색 배경으로 포맷팅됨
5. ✅ 지정한 폴더에 파일이 위치함

### Step 4: Apps Script 설정 (선택사항)

데이터 자동 새로고침을 원하는 경우:

1. 스프레드시트에서 **확장 프로그램 > Apps Script** 클릭
2. "Apps Script Code" 시트의 코드 복사
3. Apps Script 에디터에 붙여넣기
4. **서비스 +** 클릭 → **BigQuery API** 추가
5. 저장 (Ctrl+S)
6. 스프레드시트로 돌아가기
7. **BigQuery > 데이터 새로고침** 메뉴 클릭

## 주요 코드 변경 파일 목록

### 수정된 파일
- `pyproject.toml` - Google API 라이브러리 추가
- `app/utils/gcp_auth.py` - ADC 인증 로직 추가, 스코프 확장
- `app/utils/setting.py` - ADC 우선 인증 설정
- `app/services/spreadsheet_connector.py` - 스프레드시트 자동 생성 기능 구현 (500+ lines 추가)
- `app/routers/spreadsheet.py` - API 엔드포인트 추가

### 생성된 파일
- `test_sheets_api.py` - Google Sheets API 접근 테스트 스크립트
- `test_output.txt` - 테스트 결과 출력

## 기술적 세부사항

### Google Sheets API 사용

**사용된 API 메서드**:
- `spreadsheets().create()` - 스프레드시트 생성
- `spreadsheets().values().update()` - 데이터 쓰기
- `spreadsheets().batchUpdate()` - 포맷팅 적용
- `files().update()` (Drive API) - 폴더 이동

### BigQuery 통합

**데이터 로드 방식**:
1. BigQuery View에서 `SELECT * LIMIT 1000` 쿼리 실행
2. 결과를 Python 리스트로 변환
3. Google Sheets API로 일괄 업데이트

**Apps Script에서 BigQuery 접근**:
- BigQuery Advanced Service 사용
- `BigQuery.Jobs.query()` 메서드로 쿼리 실행
- 결과를 시트에 직접 작성

### 인증 메커니즘

**ADC (Application Default Credentials) 파일 위치**:
- Windows: `%APPDATA%\gcloud\application_default_credentials.json`
- Linux/Mac: `~/.config/gcloud/application_default_credentials.json`

**인증 우선순위**:
1. ADC 파일 (사용자 계정: `dc_team@madup.com`)
2. AWS Secret Manager의 서비스 계정 (fallback)

## 문제 해결 가이드

### 403 Permission Denied

**원인**:
- API가 활성화되지 않음
- 계정에 권한 없음
- 폴더 공유 설정 안됨

**해결**:
1. GCP Console에서 Google Sheets API, Drive API 활성화
2. ADC 재설정: `gcloud auth application-default login`
3. 폴더를 계정(`dc_team@madup.com`)과 공유

### ADC not available

**원인**: ADC 파일이 없거나 유효하지 않음

**해결**:
```bash
gcloud auth application-default login
```

### Insufficient authentication scopes

**원인**: ADC 생성 시 필요한 scopes가 포함되지 않음

**해결**:
```bash
gcloud auth application-default login --scopes=https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/cloud-platform
```

## 참고 자료

- [Google Sheets API v4 Documentation](https://developers.google.com/sheets/api/reference/rest)
- [Google Drive API v3 Documentation](https://developers.google.com/drive/api/v3/reference)
- [Application Default Credentials](https://cloud.google.com/docs/authentication/provide-credentials-adc)
- [Apps Script BigQuery Service](https://developers.google.com/apps-script/advanced/bigquery)

## 완료 체크리스트

- [x] Google API 라이브러리 설치
- [x] GCP 인증 스코프 추가
- [x] ADC 우선 인증 로직 구현
- [x] 스프레드시트 자동 생성 기능 구현
- [x] BigQuery 데이터 로드 기능
- [x] Apps Script 코드 생성
- [x] 가이드 시트 생성
- [x] API 엔드포인트 추가
- [x] ADC 인증 설정 완료
- [ ] **서버 재시작 필요**
- [ ] **최종 테스트 실행**
- [ ] **폴더 내 스프레드시트 생성 확인**

---

**다음 세션 시작 시 바로 실행할 명령어**:

```bash
# 1. 서버 재시작 (새 터미널에서)
uv run python main.py

# 2. 테스트 실행 (다른 터미널에서)
curl -X POST "http://localhost:8888/spreadsheet/create-with-bigquery?title=odata_test&folder_id=1HJdKCg9RBs0ky79QfsUA66r0615pJy78" -H "Authorization: Bearer test-token" -H "Content-Type: application/json"
```
