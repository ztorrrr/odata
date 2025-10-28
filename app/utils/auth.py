"""
OData 서비스 인증 모듈
HTTP Basic Authentication을 사용하여 사용자 인증 처리
"""
import logging
import secrets
from functools import cache
from typing import Optional

from fastapi import Depends, Header, HTTPException, Request, status
from fastapi.security import HTTPBasic, HTTPBasicCredentials

from app.utils import aws_secret_manager as secret_manager
from app.utils.setting import get_config

logger = logging.getLogger(__name__)

# HTTP Basic Auth 설정
security = HTTPBasic()


@cache
def get_users_config():
    """
    Secret format (JSON):
    {
        "users": [
            {
                "username": "user1",
                "password": "plain_password1"
            },
            {
                "username": "user2",
                "password": "plain_password2"
            }
        ]
    }

    Returns:
        dict: 사용자명을 키로 하는 패스워드 딕셔너리
    """
    config = get_config()
    try:
        secret_data = secret_manager.get_secret(config.ODATA_USERS_SECRET_KEY)
        users = secret_data.get("users", [])

        # username을 키로 하는 딕셔너리로 변환
        users_dict = {
            user["username"]: user["password"]
            for user in users
            if "username" in user and "password" in user
        }

        logger.info(f"Loaded {len(users_dict)} users from Secret Manager")
        return users_dict

    except Exception as e:
        logger.error(f"Failed to load users from Secret Manager: {str(e)}")
        # 개발 환경에서 Secret이 없는 경우 빈 딕셔너리 반환
        if config.ENVIRONMENT == "DEV":
            logger.warning("Authentication is disabled in DEV mode without Secret Manager configuration")
            return {}
        raise


def verify_credentials(username: str, password: str) -> bool:
    """
    Args:
        username: 사용자명
        password: 평문 패스워드

    Returns:
        bool: 인증 성공 여부
    """
    users = get_users_config()

    # 개발 환경에서 사용자 설정이 없으면 인증 패스
    if not users:
        config = get_config()
        if config.ENVIRONMENT == "DEV":
            logger.warning(f"Authentication bypassed for user '{username}' in DEV mode")
            return True
        return False

    # 사용자 존재 여부 확인
    if username not in users:
        return False

    # 타이밍 공격 방지를 위한 상수 시간 비교
    stored_password = users[username]
    password_match = secrets.compare_digest(password.encode(), stored_password.encode())

    return password_match


async def get_current_user(
    credentials: HTTPBasicCredentials = Depends(security)
) -> str:
    """
    현재 요청의 사용자를 인증하고 반환.
    FastAPI Dependency 사용.

    Args:
        credentials: HTTP Basic Auth 자격증명

    Returns:
        str: 인증된 사용자명

    Raises:
        HTTPException: 인증 실패 시 401 Unauthorized
    """
    is_valid = verify_credentials(credentials.username, credentials.password)

    if not is_valid:
        logger.warning(f"Authentication failed for user: {credentials.username}")
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Incorrect username or password",
            headers={"WWW-Authenticate": "Basic"},
        )

    logger.info(f"User authenticated: {credentials.username}")
    return credentials.username


# Optional: 특정 사용자만 허용하는 의존성
def require_user(allowed_users: list[str]):
    """
    특정 사용자만 접근을 허용하는 의존성 팩토리

    Args:
        allowed_users: 허용된 사용자명 리스트

    Returns:
        Dependency function
    """
    async def check_user(username: str = Depends(get_current_user)):
        if username not in allowed_users:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="Access denied for this user"
            )
        return username

    return check_user


@cache
def get_api_tokens_config():
    """
    API 토큰 설정을 Secret Manager에서 로드

    Secret format (JSON):
    {
        "api_tokens": [
            "tok_abc123xyz",
            "tok_def456uvw"
        ]
    }

    Returns:
        list: 유효한 토큰 리스트
    """
    config = get_config()
    try:
        secret_data = secret_manager.get_secret(config.ODATA_USERS_SECRET_KEY)
        tokens = secret_data.get("api_tokens", [])

        logger.info(f"Loaded {len(tokens)} API tokens from Secret Manager")
        return tokens

    except Exception as e:
        logger.error(f"Failed to load API tokens from Secret Manager: {str(e)}")
        # 개발 환경에서 Secret이 없는 경우 빈 리스트 반환
        if config.ENVIRONMENT == "DEV":
            logger.warning("API tokens not configured in Secret Manager")
            return []
        raise


def verify_bearer_token(token: str) -> bool:
    """
    Bearer 토큰 검증

    Args:
        token: Bearer 토큰 문자열

    Returns:
        bool: 토큰 유효성 여부
    """
    config = get_config()

    # 개발 환경에서는 모든 토큰 허용 (먼저 체크)
    if config.ENVIRONMENT == "DEV":
        logger.warning(f"Bearer token authentication bypassed in DEV mode: {token[:10]}...")
        return True

    # 프로덕션 환경에서만 Secret Manager 체크
    valid_tokens = get_api_tokens_config()

    # 토큰 설정이 없으면 거부
    if not valid_tokens:
        logger.error("No API tokens configured in Secret Manager")
        return False

    # 타이밍 공격 방지를 위한 상수 시간 비교
    for valid_token in valid_tokens:
        if secrets.compare_digest(token.encode(), valid_token.encode()):
            return True

    return False


async def get_current_user_with_header_token(
    request: Request,
    authorization: Optional[str] = Header(None, description="Authorization header with Bearer token")
) -> str:
    """
    다중 인증 방식 지원 (Excel 웹 API 인증 포함):
    1. Authorization Header의 Bearer token (Authorization: Bearer <token>)
    2. Authorization Query Parameter의 Bearer token (?Authorization=Bearer <token>) - Excel Web API
    3. URL Query Parameter의 token (?token=xxx)
    4. HTTP Basic Auth (username/password)

    Excel 웹 API 인증 사용 시:
    - 인증 유형: "웹 API"
    - 키 이름: Authorization (M 코드의 ApiKeyName으로 설정됨)
    - 키: Bearer test_token_123 형식으로 입력
    - Excel이 자동으로 "?Authorization=Bearer test_token_123" 형태로 전송

    Args:
        request: FastAPI Request 객체
        authorization: Authorization 헤더 값

    Returns:
        str: 인증된 사용자명 또는 토큰

    Raises:
        HTTPException: 인증 실패 시 401 Unauthorized
    """
    # 1. Authorization 헤더에서 Bearer token 확인
    if authorization:
        # "Bearer <token>" 형식 파싱
        parts = authorization.split()
        if len(parts) == 2 and parts[0].lower() == "bearer":
            token = parts[1]
            if verify_bearer_token(token):
                logger.info(f"Bearer token authenticated from header: {token[:10]}...")
                return f"token_user:{token[:10]}"
            else:
                logger.warning(f"Invalid bearer token from header: {token[:10]}...")
                raise HTTPException(
                    status_code=status.HTTP_401_UNAUTHORIZED,
                    detail="Invalid bearer token",
                    headers={"WWW-Authenticate": "Bearer"},
                )

    # 2. URL Query Parameter에서 Authorization 확인 (Excel Web API 방식)
    auth_param = request.query_params.get("Authorization")
    if auth_param:
        # "Bearer <token>" 형식 파싱
        parts = auth_param.split()
        if len(parts) == 2 and parts[0].lower() == "bearer":
            token = parts[1]
            if verify_bearer_token(token):
                logger.info(f"Bearer token authenticated from Authorization query param: {token[:10]}...")
                return f"token_user:{token[:10]}"
            else:
                logger.warning(f"Invalid bearer token from Authorization query param: {token[:10]}...")
                raise HTTPException(
                    status_code=status.HTTP_401_UNAUTHORIZED,
                    detail="Invalid bearer token",
                    headers={"WWW-Authenticate": "Bearer"},
                )
        else:
            logger.warning(f"Invalid Authorization query param format: {auth_param[:20]}...")

    # 3. URL Query Parameter에서 token 확인
    token_param = request.query_params.get("token")
    if token_param:
        if verify_bearer_token(token_param):
            logger.info(f"Bearer token authenticated from token query param: {token_param[:10]}...")
            return f"token_user:{token_param[:10]}"
        else:
            logger.warning(f"Invalid bearer token from token query param: {token_param[:10]}...")
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Invalid bearer token",
                headers={"WWW-Authenticate": "Bearer"},
            )

    # 4. HTTP Basic Auth 확인
    # Authorization 헤더가 "Basic" 형식인지 확인
    if authorization and authorization.lower().startswith("basic"):
        try:
            # HTTPBasicCredentials로 파싱
            import base64
            encoded_credentials = authorization.split(" ", 1)[1]
            decoded = base64.b64decode(encoded_credentials).decode("utf-8")
            username, password = decoded.split(":", 1)

            is_valid = verify_credentials(username, password)
            if not is_valid:
                logger.warning(f"Authentication failed for user: {username}")
                raise HTTPException(
                    status_code=status.HTTP_401_UNAUTHORIZED,
                    detail="Incorrect username or password",
                    headers={"WWW-Authenticate": "Basic"},
                )

            logger.info(f"User authenticated: {username}")
            return username
        except Exception as e:
            logger.warning(f"Failed to parse Basic auth: {e}")
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Invalid authentication credentials",
                headers={"WWW-Authenticate": "Basic, Bearer"},
            )

    # 인증 정보가 없음
    raise HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Authentication required. Use Bearer token or Basic auth.",
        headers={"WWW-Authenticate": "Bearer, Basic"},
    )
