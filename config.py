import os
from pydantic_settings import BaseSettings

class Settings(BaseSettings):
    base_url: str = "http://10.76.75.20/demo/"
    login: str = "KAI25"
    password: str = "Q123456q"
    report_id: str = "7465547992"
    employer: str = "11421130428"
    excel_path: str = "и38 таблица МИАЦ.xlsx"
    excel_path_ticket: str = "miac_table.xlsx"
    lpu_path: str = "LPU.xlsx"

settings = Settings()
