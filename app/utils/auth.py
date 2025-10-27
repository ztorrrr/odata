"""
OData 서비스 인증 모듈
HTTP Basic Authentication을 사용하여 사용자 인증 처리
"""
import logging
import secrets
from functools import cache
from typing import Optional

from fastapi import Depends, HTTPException, status
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
