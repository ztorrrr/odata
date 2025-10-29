# Web API 인증 구현 작업 기록
2025-10-28

## 작업 배경
Excel의 OData 연결 시 "웹 API" 탭을 통한 API Key 인증 방식 구현 요청. 기존에는 Basic Auth(ID/PW) 방식만 지원.

## 브랜치 관리
- 기존 변경사항 reset 후 `feature/webapi-test` 브랜치 생성
- master 브랜치의 마지막 커밋(ec70045)부터 작업 시작

## 구현 내용

### 1. API Key 인증 기능 추가 (app/utils/auth.py)

**추가된 함수들:**
- `get_api_tokens_config()`: AWS Secret Manager에서 API 토큰 목록 로드
- `verify_bearer_token()`: Bearer 토큰 검증 (DEV 환경에서는 모든 토큰 허용)
- `get_current_user_with_header_token()`: 다중 인증 방식 지원
  - Authorization Header의 Bearer token
  - URL Query Parameter의 Bearer token (Excel Web API 방식)
  - URL Query Parameter의 token
  - HTTP Basic Auth

### 2. Excel COM Generator 개선 (app/services/excel_com_generator.py)

**새로운 메서드:**
- `create_odata_connection_with_webapi_auth()`: Web API 인증용 Excel 생성
- `create_excel_with_webapi_auth_com()`: 편의 함수

**Power Query M 코드에 ApiKeyName 파라미터 추가:**
```m
OData.Feed(url, null, [Implementation="2.0", ApiKeyName="Authorization"])
```

**COM 안정성 개선:**
- `DispatchEx` 사용으로 새 Excel 인스턴스 생성
- 재시도 메커니즘 추가 (최대 3회)
- `CoInitializeEx(COINIT_APARTMENTTHREADED)`로 멀티스레드 지원
- cleanup 과정 개선

### 3. 새 엔드포인트 추가 (app/routers/odata.py)

**Web API 인증용 엔드포인트:**
- `/odata/musinsa_data/excel-com-webapi-key`: Web API 템플릿 다운로드 (인증 불필요)

**기존 엔드포인트 수정:**
- `/odata/musinsa_data/excel-com`: Basic Auth 템플릿 다운로드 (인증 불필요로 변경)

### 4. 데이터 엔드포인트 인증 확장

**다중 인증 지원으로 업그레이드:**
- `/odata/musinsa_data`: `get_current_user` → `get_current_user_with_header_token`
- `/odata/musinsa_data/$count`: 동일 변경
- `/odata/musinsa_data/export`: 동일 변경

모든 데이터 엔드포인트가 Basic Auth와 Bearer Token 인증을 모두 지원.

## 해결된 문제들

### Excel COM 오류 (OLE error 0x800ac472)
- 원인: Excel이 이미 실행 중일 때 COM 충돌
- 해결: DispatchEx로 새 인스턴스 생성, 재시도 메커니즘 추가

### 401 Unauthorized 오류
- 원인: 데이터 엔드포인트가 Basic Auth만 지원
- 해결: get_current_user_with_header_token으로 모든 인증 방식 지원

### 중복 인증 문제
- 원인: Excel 템플릿 다운로드와 데이터 새로고침 시 각각 인증 필요
- 해결: 템플릿 다운로드는 인증 불필요로 변경, 데이터 접근 시에만 인증

## 최종 상태

**템플릿 다운로드 (인증 불필요):**
- `/excel-com`: Basic Auth 템플릿
- `/excel-com-webapi-key`: Web API 템플릿

**데이터 접근 (인증 필요):**
- Basic Auth 지원
- Bearer Token 지원 (Header, Query Parameter)
- Excel Web API 방식 지원 (?Authorization=Bearer token)

**사용 방법:**
1. 템플릿 다운로드 (브라우저에서 직접 접근 가능)
2. Excel에서 데이터 새로고침
3. 인증 대화상자에서 선택:
   - "기본" 탭: ID/PW 입력
   - "웹 API" 탭: Bearer token 입력

## 테스트 가이드 문서
- WEBAPI_AUTH_TEST_GUIDE.md 생성 (테스트 절차 및 설정 방법 기록)