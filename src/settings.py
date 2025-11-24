import os
from pydantic_settings import BaseSettings
from functools import lru_cache
import secrets

class Settings(BaseSettings):
    PORT: int
    CLIENT_ID: str
    CLIENT_SECRET: str
    TENANT_ID: str
    OAUTH_CONNECTION_NAME: str
    OPENAI_API_KEY: str

    class Config:
        env_file = ".env"
    

@lru_cache()
def get_settings():
    return Settings()

