import time

from fastapi import FastAPI, BackgroundTasks, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel
import tasks
from utils import get_last_week_dates
import os
from config import settings

app = FastAPI(title="MIAC Report Service")


class TaskStatus(BaseModel):
    task_id: str
    status: str
    progress: int = 0
    failures: int = 0


tasks_list = {}


@app.get("/")
async def root():
    return {"message": "MIAC Report Service готов", "docs": "/docs"}


@app.post("/start-report", response_model=dict)
async def start_report(background_tasks: BackgroundTasks):
    try:
        background_tasks.add_task(run_miac_direct)
        task_id = f"direct_{int(time.time())}"
        return {"task_id": task_id, "status": "started"}
    except Exception as e:
        raise HTTPException(500, f"Ошибка запуска: {str(e)}")

def run_miac_direct():
    try:
        from utils import main_process
        duration = main_process()
        print(f"Обработка завершена за {duration:.2f} секунд")
    except Exception as e:
        print(f"Ошибка обработки: {e}")



@app.get("/status/{task_id}", response_model=TaskStatus)
async def get_status(task_id: str):
    if task_id not in tasks_list:
        raise HTTPException(404, "Задача не найдена")

    task = tasks.AsyncResult(task_id)
    status = {
        "task_id": task_id,
        "status": task.status,
        **tasks_list[task_id]
    }

    if task.status == "SUCCESS":
        status["progress"] = 90
    return status


@app.get("/download-report")
async def download_report():
    if not os.path.exists(settings.excel_path):
        raise HTTPException(404, "Отчет не готов")
    return FileResponse(
        settings.excel_path,
        filename="миац_отчет.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.get("/last-week-dates")
async def dates():
    return {"from": get_last_week_dates()[0], "to": get_last_week_dates()[1]}
