import time
import os
from datetime import datetime
from fastapi import FastAPI, BackgroundTasks, HTTPException, Form, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from utils import main_process, get_last_week_dates
from config import settings

app = FastAPI(title="MIAC Report Service")
templates = Jinja2Templates(directory="templates")
scheduler = BackgroundScheduler()

# Состояние планировщика
schedule_config = {
    "enabled": False,
    "day_of_week": 0,  # 0=Понедельник
    "hour": 10,
    "minute": 0
}


def scheduled_report():
    """Автоматическая выгрузка по расписанию"""
    print(f"🕐 [{datetime.now().strftime('%H:%M:%S')}] Автовыгрузка запущена")
    try:
        duration = main_process()
        print(f"✅ [{datetime.now().strftime('%H:%M:%S')}] Автовыгрузка завершена за {duration:.2f}с")
    except Exception as e:
        print(f"❌ [{datetime.now().strftime('%H:%M:%S')}] Ошибка автозагрузки: {e}")


def run_miac_direct():
    """Ручная выгрузка"""
    print(f"▶️ [{datetime.now().strftime('%H:%M:%S')}] Ручная выгрузка запущена")
    try:
        duration = main_process()
        print(f"✅ [{datetime.now().strftime('%H:%M:%S')}] Обработка завершена за {duration:.2f} секунд")
    except Exception as e:
        print(f"❌ [{datetime.now().strftime('%H:%M:%S')}] Ошибка обработки: {e}")


@app.on_event("startup")
async def startup():
    """Запуск планировщика при старте приложения"""
    if schedule_config["enabled"]:
        day_names = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun']
        trigger = CronTrigger(
            day_of_week=day_names[schedule_config["day_of_week"]],
            hour=schedule_config["hour"],
            minute=schedule_config["minute"]
        )
        scheduler.add_job(scheduled_report, trigger)
        scheduler.start()
        print("🕐 Планировщик запущен:", schedule_config)


@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request):
    """Главная страница с настройками расписания"""
    return templates.TemplateResponse("schedule.html", {
        "request": request,
        "config": schedule_config
    })


@app.post("/schedule")
async def set_schedule(
        enabled: bool = Form(False),
        day: int = Form(0),
        hour: int = Form(10),
        minute: int = Form(0)
):
    """Установка расписания"""
    global schedule_config

    # Останавливаем старые задачи
    scheduler.remove_all_jobs()

    schedule_config.update({
        "enabled": enabled,
        "day_of_week": day,
        "hour": hour,
        "minute": minute
    })

    day_names = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье']

    if enabled:
        day_names_cron = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun']
        trigger = CronTrigger(
            day_of_week=day_names_cron[day],
            hour=hour,
            minute=minute
        )
        scheduler.add_job(scheduled_report, trigger)
        scheduler.start()
        return {"status": "ok", "message": f"🕐 Запланировано: каждый {day_names[day]} {hour:02d}:{minute:02d}"}

    return {"status": "ok", "message": "⏹️ Планировщик остановлен"}


@app.post("/start-report", response_model=dict)
async def start_report(background_tasks: BackgroundTasks):
    """Запуск ручной выгрузки (для API)"""
    background_tasks.add_task(run_miac_direct)
    task_id = f"direct_{int(time.time())}"
    return {"task_id": task_id, "status": "started", "time": datetime.now().strftime("%H:%M:%S")}


@app.get("/manual")
async def manual_report(background_tasks: BackgroundTasks):
    """Ручной запуск с главной страницы"""
    background_tasks.add_task(run_miac_direct)
    return {"status": "started", "time": datetime.now().strftime("%H:%M:%S")}


@app.get("/status")
async def status():
    """Статус планировщика и файла"""
    jobs_count = len(scheduler.get_jobs())
    file_size = os.path.getsize(settings.excel_path) if os.path.exists(settings.excel_path) else 0
    file_date = datetime.fromtimestamp(os.path.getmtime(settings.excel_path)).strftime("%d.%m %H:%M") if os.path.exists(
        settings.excel_path) else "Нет файла"

    day_names = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс']
    next_run = f"{day_names[schedule_config['day_of_week']]} {schedule_config['hour']:02d}:{schedule_config['minute']:02d}" if \
    schedule_config['enabled'] else "Выключен"

    return {
        "enabled": schedule_config["enabled"],
        "next_run": next_run,
        "jobs_count": jobs_count,
        "file_size": file_size,
        "file_date": file_date,
        "last_dates": get_last_week_dates()
    }


@app.get("/download-report")
async def download_report():
    """Скачать готовый отчет"""
    if not os.path.exists(settings.excel_path):
        raise HTTPException(404, "❌ Отчет не готов. Запустите выгрузку!")
    return FileResponse(
        settings.excel_path,
        filename="миац_отчет.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.get("/last-week-dates")
async def dates():
    """Даты прошлой недели (для отладки)"""
    return {"from": get_last_week_dates()[0], "to": get_last_week_dates()[1]}


@app.get("/docs")
async def docs_redirect():
    """Перенаправление на Swagger"""
    return {"docs": "/docs", "ui": "http://localhost:8000/docs"}
