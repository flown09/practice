import sys
if sys.platform == "win32":
    import asyncio
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
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
import zipfile
import io
from fastapi.responses import StreamingResponse
from playwright.async_api import async_playwright
import asyncio
from concurrent.futures import ThreadPoolExecutor
import tempfile
import shutil
from playwright.sync_api import sync_playwright

app = FastAPI(title="MIAC Report Service")
templates = Jinja2Templates(directory="templates")

schedule_config = {
    "enabled": False,
    "day_of_week": 0,
    "hour": 10,
    "minute": 0
}

scheduler = None


def sync_auto_scrape():
    """Synchronous Playwright scraper - runs in thread"""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={'width': 1920, 'height': 1080},
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        )
        page = context.new_page()

        # TODO: Replace with real login flow from utils.py
        page.goto("https://info-bi-db.egisz.rosminzdrav.ru/")
        # page.fill("input[name='login']", settings.login)  # Adapt selectors
        # page.fill("input[name='password']", settings.password)
        # page.click("button[type='submit']")
        page.wait_for_load_state("networkidle")

        # Navigate to dashboard
        page.goto(
            "https://info-bi-db.egisz.rosminzdrav.ru/dashboardsViewer?sectionId=27&dashboardId=8d82093225eb470595ae4d49d2edc555&sheetId=75e7b48008b34db482c350b2333e2d45")
        page.wait_for_load_state("networkidle")

        # Inject sd.js scraping logic (paste BODYS + functions here)
        with open("static/extension/sd.js", "r", encoding="utf-8") as f:
            sd_js = f.read()
        page.add_init_script(sd_js)

        # Set dates to last week (from utils.py)
        from utils import get_last_week_dates
        date_from, date_to = get_last_week_dates()
        page.evaluate(f"""
            // Replace DSTART/DSTOP in BODYS with '{date_from}', '{date_to}'
            // Trigger addB() or main scraping function
            window.run_scrape();  // Define this in sd.js injection
        """)

        # Download parse.json (adapt selector)
        download_path = tempfile.mkdtemp()
        with page.expect_download(timeout=30000) as download_info:
            page.evaluate("writeFile('parse.json', JSON.stringify(result));")
        download = download_info.value
        download.save_as(f"{download_path}/parse.json")

        # Copy result
        shutil.copy(f"{download_path}/parse.json", "scraped_parse.json")
        browser.close()
        return {"status": "success", "file": "scraped_parse.json"}


@app.get("/auto-scrape")
async def auto_scrape():
    """Async wrapper - runs sync scraper in thread"""
    loop = asyncio.get_event_loop()
    with ThreadPoolExecutor(max_workers=1) as executor:
        result = await loop.run_in_executor(executor, sync_auto_scrape)
    if os.path.exists("scraped_parse.json"):
        return FileResponse("scraped_parse.json", filename="parse.json", media_type="application/json")
    return {"error": "Scrape failed"}

@app.get("/auto-scrape")
async def auto_scrape():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            viewport={'width': 1920, 'height': 1080},
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        )
        page = await context.new_page()

        # Navigate and login (adapt selectors/credentials from utils.py)
        await page.goto("https://info-bi-db.egisz.rosminzdrav.ru/login")  # Adjust login URL
        await page.fill("input[name='login']", "YOUR_GOSUSLUGI_LOGIN")  # Use env vars
        await page.fill("input[name='password']", "YOUR_GOSUSLUGI_PASSWORD")
        await page.click("button[type='submit']")
        await page.wait_for_load_state("networkidle")

        # Go to dashboard
        await page.goto(
            "https://info-bi-db.egisz.rosminzdrav.ru/dashboardsViewer?sectionId=27&dashboardId=8d82093225eb470595ae4d49d2edc555&sheetId=75e7b48008b34db482c350b2333e2d45")
        await page.wait_for_load_state("networkidle")

        # Inject and run sd.js logic (paste BODYS array and functions from sd.js)
        await page.add_init_script("""
            // Paste full sd.js content here, replace window.onload with immediate execution
            // Set dates to last week via get_last_week_dates() logic
            // Trigger addB() to scrape and generate parse.json
        """)
        await page.evaluate("run_scrape();")  # Custom function wrapping sd.js logic

        # Download parse.json
        download_path = tempfile.mkdtemp()
        async with page.expect_download() as download_info:
            await page.evaluate("writeFile('parse.json', JSON.stringify(result));")  # Trigger download
        download = await download_info.value
        await download.save_as(f"{download_path}/parse.json")

        data = await page.evaluate("JSON.parse(localStorage.getItem('scraped_data'))")  # Fallback

        await browser.close()

        # Process/parse json and integrate with main_process() if needed
        shutil.copy(f"{download_path}/parse.json", "parse.json")
        return FileResponse("parse.json", filename="parse.json")

@app.get("/install-extension")
async def install_extension():
    """Скачивает готовый .crx файл расширения"""

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.write("static/extension/manifest.json", "manifest.json")
        zip_file.write("static/extension/script-loader.js", "script-loader.js")
        zip_file.write("static/extension/sd.js", "sd.js")
        if os.path.exists("static/extension/img/icon16.png"):
            zip_file.write("static/extension/img/icon16.png", "img/icon16.png")
            zip_file.write("static/extension/img/icon48.png", "img/icon48.png")
            zip_file.write("static/extension/img/icon128.png", "img/icon128.png")

    zip_buffer.seek(0)

    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={
            "Content-Disposition": "attachment; filename=miac-dashboard-extension.zip"
        }
    )

def create_scheduler():
    """Создает и возвращает новый чистый планировщик"""
    global scheduler
    return BackgroundScheduler()

def scheduled_report():
    """Автоматическая выгрузка по расписанию"""
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Автовыгрузка запущена")
    try:
        duration = main_process()
        print(f"[{(datetime.now().strftime('%H:%M:%S'))}] Автовыгрузка завершена за {duration:.2f}с")
    except Exception as e:
        print(f"[{(datetime.now().strftime('%H:%M:%S'))}] Ошибка автозагрузки: {e}")

def run_miac_direct():
    """Ручная выгрузка"""
    print(f"[{(datetime.now().strftime('%H:%M:%S'))}] Ручная выгрузка запущена")
    try:
        duration = main_process()
        print(f"[{(datetime.now().strftime('%H:%M:%S'))}] Обработка завершена за {duration:.2f} секунд")
    except Exception as e:
        print(f"[{(datetime.now().strftime('%H:%M:%S'))}] Ошибка обработки: {e}")

@app.on_event("startup")
async def startup():
    """Запуск планировщика при старте приложения"""
    global scheduler
    if schedule_config["enabled"]:
        scheduler = create_scheduler()
        day_names = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun']
        trigger = CronTrigger(
            day_of_week=day_names[schedule_config["day_of_week"]],
            hour=schedule_config["hour"],
            minute=schedule_config["minute"]
        )
        scheduler.add_job(scheduled_report, trigger)
        scheduler.start()
        print("Планировщик запущен:", schedule_config)

@app.post("/schedule")
async def set_schedule(
        enabled: bool = Form(False),
        day: int = Form(0),
        hour: int = Form(10),
        minute: int = Form(0)
):
    """Установка расписания"""
    global scheduler, schedule_config

    if scheduler is not None:
        if scheduler.running:
            scheduler.shutdown(wait=True)  # Ждем завершения
        scheduler = None

    schedule_config.update({
        "enabled": enabled,
        "day_of_week": day,
        "hour": hour,
        "minute": minute
    })

    day_names = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье']

    if enabled:
        scheduler = create_scheduler()
        day_names_cron = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun']
        trigger = CronTrigger(
            day_of_week=day_names_cron[day],
            hour=hour,
            minute=minute
        )
        scheduler.add_job(scheduled_report, trigger)
        scheduler.start()
        print(f"Запланировано: каждый {day_names[day]} {hour:02d}:{minute:02d}")
        return {"status": "ok", "message": f"Запланировано: каждый {day_names[day]} {hour:02d}:{minute:02d}"}

    print("Планировщик остановлен")
    return {"status": "ok", "message": "Планировщик остановлен"}

@app.get("/status")
async def status():
    """Статус планировщика и файла"""
    global scheduler
    jobs_count = len(scheduler.get_jobs()) if scheduler and scheduler.running else 0
    file_size = os.path.getsize(settings.excel_path) if os.path.exists(settings.excel_path) else 0
    file_date = datetime.fromtimestamp(os.path.getmtime(settings.excel_path)).strftime("%d.%m %H:%M") if os.path.exists(
        settings.excel_path) else "Нет файла"

    day_names = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс']
    next_run = f"{day_names[schedule_config['day_of_week']]} {schedule_config['hour']:02d}:{schedule_config['minute']:02d}" if \
    schedule_config['enabled'] else "Выключен"

    return {
        "enabled": schedule_config["enabled"],
        "scheduler_running": scheduler.running if scheduler else False,
        "next_run": next_run,
        "jobs_count": jobs_count,
        "file_size": file_size,
        "file_date": file_date,
        "last_dates": get_last_week_dates()
    }


@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request):
    """Главная страница с настройками расписания"""
    return templates.TemplateResponse("schedule.html", {
        "request": request,
        "config": schedule_config
    })


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


@app.get("/download-report")
async def download_report():
    """Скачать готовый отчет"""
    if not os.path.exists(settings.excel_path):
        raise HTTPException(404, "Отчет не готов. Запустите выгрузку!")
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