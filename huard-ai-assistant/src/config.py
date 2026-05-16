from pathlib import Path
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", env_file_encoding="utf-8", extra="ignore")

    anthropic_api_key: str
    claude_model: str = "claude-sonnet-4-6"
    claude_model_fast: str = "claude-haiku-4-5-20251001"

    ms_tenant_id: str = ""
    ms_client_id: str = ""
    ms_client_secret: str = ""
    ms_user_email: str = ""

    twilio_account_sid: str = ""
    twilio_auth_token: str = ""
    twilio_whatsapp_from: str = ""

    wa_perso_enabled: bool = False
    wa_perso_session_dir: str = "./data/wa_perso_session"

    app_env: str = "development"
    app_host: str = "0.0.0.0"
    app_port: int = 8000
    app_public_url: str = ""
    log_level: str = "INFO"

    database_url: str = "sqlite:///./data/huard_ai.db"
    webhook_secret: str = ""

    knowledge_dir: Path = Path("knowledge")


settings = Settings()
