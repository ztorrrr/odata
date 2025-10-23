ㄱ#!/usr/bin/env python
"""
OData 서버 실행 스크립트
"""
import sys
from pathlib import Path

# 프로젝트 루트를 Python 경로에 추가
sys.path.insert(0, str(Path(__file__).parent))

import uvicorn
from app.utils.setting import get_config

if __name__ == "__main__":
    config = get_config()

    print(f"""
╔═══════════════════════════════════════════════════════════╗
║        OData v4 Service for BigQuery                      ║
╠═══════════════════════════════════════════════════════════╣
║  Environment: {config.ENVIRONMENT:<44} ║
║  Host:        {config.HOST:<44} ║
║  Port:        {config.PORT:<44} ║
║  Dataset:     {config.BIGQUERY_DATASET_ID:<44} ║
║  Table:       {config.BIGQUERY_TABLE_NAME:<44} ║
╚═══════════════════════════════════════════════════════════╝

🚀 Starting server...

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
        log_level=config.LOG_LEVEL.lower()
    )