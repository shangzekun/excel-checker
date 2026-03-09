import os
from typing import List

def _parse_csv_env(value: str, default: List[str]) -> List[str]:
    if not value:
        return default
    items = [x.strip() for x in value.split(",") if x.strip()]
    return items or default

class Settings:
    APP_NAME: str = os.getenv("APP_NAME", "Excel Checker")
    APP_VERSION: str = os.getenv("APP_VERSION", "0.5.0")
    MAX_MB: int = int(os.getenv("MAX_MB", "10"))
    LOG_LEVEL: str = os.getenv("LOG_LEVEL", "INFO").upper()
    ALLOW_ORIGINS: List[str] = _parse_csv_env(
        os.getenv("ALLOW_ORIGINS", "http://127.0.0.1:8000,http://localhost:8000"),
        ["http://127.0.0.1:8000", "http://localhost:8000"],
    )
    ENV: str = os.getenv("ENV", "dev").lower()

settings = Settings()
