from pydantic import Field
from pydantic_settings import BaseSettings
import os


class Env(BaseSettings):
    forest_data_dir: str = Field(description='Path to the forest data directory which must contain Initial_forest_data.xlsx')
    volunteer_name: str = Field(description='Сhecker\'s name')

    class Config:
        env_file = os.path.dirname(os.path.abspath(__file__))+"/.env"


SETTINGS = Env()
