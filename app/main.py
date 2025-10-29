"""
FastAPI OData Service
"""
import logging
from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from app.routers import odata, spreadsheet
from app.utils.setting import get_config, setup_gcp_auth
from app.services.bigquery_service import get_bigquery_service

# 로깅 설정
config = get_config()
logging.basicConfig(
    level=getattr(logging, config.LOG_LEVEL),
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


@asynccontextmanager
async def lifespan(app: FastAPI):
    """
    애플리케이션 시작/종료 시 실행되는 코드
    """
    # 시작 시 실행
    logger.info("Starting OData Service...")
    logger.info(f"Environment: {config.ENVIRONMENT}")

    try:
        # GCP 인증 설정
        logger.info("Setting up GCP authentication...")
        gcp_auth = setup_gcp_auth()
        logger.info(f"Authenticated with project: {gcp_auth.project_id}")

        # BigQuery 서비스 초기화
        logger.info("Initializing BigQuery service...")
        bq_service = get_bigquery_service()
        bq_service.initialize()

        # 테이블 정보 확인
        table_info = bq_service.get_table_info()
        if table_info:
            logger.info(f"BigQuery table ready: {table_info['table_id']}")
            logger.info(f"Total rows: {table_info['num_rows']:,}")
        else:
            logger.warning("BigQuery table not found. Please run load_data.py first.")

        logger.info("OData Service started successfully!")

    except Exception as e:
        logger.error(f"Failed to initialize service: {str(e)}", exc_info=True)
        raise

    yield

    # 종료 시 실행
    logger.info("Shutting down OData Service...")


# FastAPI 앱 생성
app = FastAPI(
    title="OData v4 Service for BigQuery",
    description="OData v4 API for accessing BigQuery data from Excel and other tools",
    version=config.ODATA_SERVICE_VERSION,
    docs_url=config.DOCS_URL,
    redoc_url=config.REDOC_URL,
    lifespan=lifespan
)

# CORS 설정 (Excel에서 접근 가능하도록)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 프로덕션에서는 특정 도메인으로 제한
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["OData-Version"],
)


# 라우터 등록
app.include_router(odata.router)
app.include_router(spreadsheet.router)


@app.get("/")
async def root():
    """
    루트 엔드포인트
    """
    return {
        "service": "OData v4 Service for BigQuery",
        "version": config.ODATA_SERVICE_VERSION,
        "odata_endpoint": "/odata",
        "metadata": "/odata/$metadata",
        "spreadsheet_endpoint": "/spreadsheet",
        "health": "/odata/health",
        "documentation": config.DOCS_URL if config.DOCS_URL else "disabled"
    }


@app.get("/health")
async def health():
    """
    헬스 체크
    """
    return {"status": "healthy"}


@app.exception_handler(Exception)
async def global_exception_handler(request, exc):
    """
    전역 예외 핸들러
    """
    logger.error(f"Unhandled exception: {str(exc)}", exc_info=True)
    return JSONResponse(
        status_code=500,
        content={
            "error": {
                "code": "InternalServerError",
                "message": "An internal error occurred"
            }
        }
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "app.main:app",
        host=config.HOST,
        port=config.PORT,
        reload=True if config.ENVIRONMENT == "DEV" else False,
        log_level=config.LOG_LEVEL.lower()
    )