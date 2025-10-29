# Web API 인증 테스트 가이드

## 개요
Excel의 "웹 API" 탭을 사용하여 API Key로 인증하는 새로운 엔드포인트를 테스트하는 방법입니다.

## 새로운 엔드포인트

### `/odata/musinsa_data/excel-com-webapi-key`
- **용도**: Web API 인증 방식의 Excel 파일 생성
- **인증**: Bearer Token (API Key)
- **인증 방법**: Excel에서 "웹 API" 탭 사용

## 테스트 절차

### 1. 서버 실행
```bash
uv run python main.py
```

### 2. Excel 파일 다운로드
브라우저에서 다음 URL 접속:
```
http://localhost:8888/odata/musinsa_data/excel-com-webapi-key
```

이 엔드포인트는 Bearer 토큰 인증이 필요.
- DEV 환경에서는 토큰 검증이 우회됨.
- PROD 환경에서는 AWS Secret Manager에서 api_tokens 설정 필요.

### 3. Excel에서 데이터 새로고침

1. 다운로드받은 Excel 파일 열기
2. 데이터 탭 → 쿼리 및 연결 클릭
3. 우측 패널에서 쿼리 우클릭 → 새로고침
4. 인증 대화상자가 나타나면:
   - **"웹 API" 탭 선택**
   - **키 입력란에 다음 형식으로 입력**:
     ```
     Bearer tok_abc123xyz
     ```
   - "연결" 클릭

### 4. 인증 설정 확인

#### DEV 환경
- 어떤 토큰 값을 입력해도 인증 통과
- 로그에 "Bearer token authentication bypassed in DEV mode" 메시지 확인

#### PROD 환경
AWS Secret Manager에 다음 형식으로 토큰 설정:
```json
{
  "users": [
    {"username": "user1", "password": "pass1"}
  ],
  "api_tokens": [
    "tok_abc123xyz",
    "tok_def456uvw"
  ]
}
```

## 주요 변경사항

### 1. `app/services/excel_com_generator.py`
- `create_odata_connection_with_webapi_auth()` 메서드 추가
- Power Query M 코드에 `ApiKeyName` 파라미터 추가
- `create_excel_with_webapi_auth_com()` 편의 함수 추가

### 2. `app/routers/odata.py`
- `/excel-com-webapi-key` 엔드포인트 추가
- `get_current_user_with_header_token` 의존성 사용
- Web API 인증 방식 지원

### 3. `app/utils/auth.py` (기존 기능 활용)
- `get_current_user_with_header_token()`: Bearer 토큰 인증 지원
- URL 쿼리 파라미터로 Authorization 전달 지원 (Excel Web API 방식)

## Excel Web API 인증 흐름

1. Excel이 데이터 새로고침 시 Power Query의 `ApiKeyName` 설정을 감지
2. "웹 API" 인증 탭이 활성화됨
3. 사용자가 입력한 키를 URL 쿼리 파라미터로 전송:
   ```
   ?Authorization=Bearer%20tok_abc123xyz
   ```
4. FastAPI가 `get_current_user_with_header_token()`에서 토큰 검증
5. 인증 성공 시 데이터 반환

## 테스트 시나리오

### 시나리오 1: 기본 인증 (기존 방식)
- 엔드포인트: `/odata/musinsa_data/excel-com`
- Excel 인증: "기본" 탭에서 사용자명/비밀번호 입력

### 시나리오 2: Web API 인증 (새 방식)
- 엔드포인트: `/odata/musinsa_data/excel-com-webapi-key`
- Excel 인증: "웹 API" 탭에서 Bearer 토큰 입력