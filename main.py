#!/usr/bin/env python
"""
OData v4 Service for BigQuery
Main entry point for running the server
"""
import os
import sys
import shutil
from pathlib import Path

# 프로젝트 루트를 Python 경로에 추가
sys.path.insert(0, str(Path(__file__).parent))

# Python 바이트코드 캐싱 비활성화
os.environ['PYTHONDONTWRITEBYTECODE'] = '1'

import uvicorn
from app.utils.setting import get_config


def clear_pycache():
    """__pycache__ 디렉토리 정리"""
    project_root = Path(__file__).parent
    cache_dirs = list(project_root.rglob('__pycache__'))

    for cache_dir in cache_dirs:
        try:
            shutil.rmtree(cache_dir)
        except Exception:
            pass


def main():
    """OData 서버 실행"""
    # 캐시 정리
    clear_pycache()

    config = get_config()

    print(f"""
========================================================
         OData v4 Service for BigQuery
========================================================
  Environment: {config.ENVIRONMENT}
  Host:        {config.HOST}
  Port:        {config.PORT}
  Dataset:     {config.BIGQUERY_DATASET_ID}
  Table:       {config.BIGQUERY_TABLE_NAME}
========================================================

Starting server...

OData Endpoint:    http://{config.HOST}:{config.PORT}/odata
Metadata:          http://{config.HOST}:{config.PORT}/odata/$metadata
Service Document:  http://{config.HOST}:{config.PORT}/odata/
Health Check:      http://{config.HOST}:{config.PORT}/odata/health
""")

    uvicorn.run(
        "app.main:app",
        host=config.HOST,
        port=config.PORT,
        reload=config.ENVIRONMENT == "DEV",
        reload_dirs=["app"] if config.ENVIRONMENT == "DEV" else None,
        reload_delay=0.25,  # 파일 변경 감지 지연 시간 (초)
        log_level=config.LOG_LEVEL.lower(),
        use_colors=True,
        access_log=True
    )


if __name__ == "__main__":
    main()
