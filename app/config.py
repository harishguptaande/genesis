from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    app_name: str = "Genesis"
    debug: bool = False
    env: str = "development"
    database_url: str = "sqlite:///./test.db"

    class Config:
        env_file = ".env"


settings = Settings()
