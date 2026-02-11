import time
import os
from datetime import datetime
import requests
from requests.exceptions import SSLError as RequestsSSLError, RequestException

from fastapi import FastAPI, BackgroundTasks, HTTPException, Form, Request, Response
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse, PlainTextResponse
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from utils import main_process, get_last_week_dates
from config import settings

app = FastAPI(title="MIAC Report Service")
templates = Jinja2Templates(directory="templates")

# Убираем глобальный scheduler, делаем его локальным
schedule_config = {
    "enabled": False,
    "day_of_week": 0,
    "hour": 10,
    "minute": 0
}

scheduler = None  # Глобальная переменная для текущего планировщика

EXTERNAL_BASE_URL = "https://info-bi-db.egisz.rosminzdrav.ru"

def _rewrite_location_header(location: str) -> str:
    if location.startswith(EXTERNAL_BASE_URL):
        return location[len(EXTERNAL_BASE_URL):] or "/"
    return location

def _egisz_ssl_verify_value():
    """Параметр verify для requests к внешнему дашборду."""
    if settings.egisz_ca_bundle:
        return settings.egisz_ca_bundle
    return settings.egisz_ssl_verify


def _rewrite_html_for_proxy(html: str) -> str:
    """Нормализует абсолютные ссылки внешнего дашборда в same-origin ссылки."""
    html = html.replace(EXTERNAL_BASE_URL, "")
    html = html.replace("//info-bi-db.egisz.rosminzdrav.ru", "")
    return html

def _inject_autostart_script(html: str) -> str:
    script_tag = '<script src="/autostart.js"></script>'
    if script_tag in html:
        return html
    if "</body>" in html:
        return html.replace("</body>", f"{script_tag}</body>")
    return html + script_tag


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

    # Полностью останавливаем старый планировщик
    if scheduler is not None:
        if scheduler.running:
            scheduler.shutdown(wait=True)  # Ждем завершения
        scheduler = None

    # Обновляем конфигурацию
    schedule_config.update({
        "enabled": enabled,
        "day_of_week": day,
        "hour": hour,
        "minute": minute
    })

    day_names = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье']

    if enabled:
        # Создаем НОВЫЙ планировщик
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


@app.get("/create-report")
async def create_report():
    """Перенаправление на внешний дашборд для авторизации через Госуслуги"""
    return RedirectResponse(
        "/egisz/dashboardsViewer?sectionId=27&dashboardId=8d82093225eb470595ae4d49d2edc555&sheetId=75e7b48008b34db482c350b2333e2d45"
    )




@app.get("/autostart.js")
async def autostart_script():
    """JS для автозапуска выгрузки на проксированной странице дашборда"""
    with open("extension/sd.js", "r", encoding="utf-8") as f:
        script = f.read()
    return PlainTextResponse(script, media_type="application/javascript")


async def _proxy_to_external(request: Request, full_path: str, inject_autostart: bool = False):
    query = ("?" + request.url.query) if request.url.query else ""
    target_url = f"{EXTERNAL_BASE_URL}/{full_path}{query}"

    incoming_headers = dict(request.headers)
    incoming_headers.pop("host", None)
    incoming_headers.pop("content-length", None)
    incoming_headers["accept-encoding"] = "identity"

    body = await request.body()

    verify_value = _egisz_ssl_verify_value()
    used_insecure_ssl_fallback = False
    try:
        upstream = requests.request(
            request.method,
            target_url,
            headers=incoming_headers,
            data=body,
            allow_redirects=False,
            timeout=60,
            verify=verify_value,
        )
    except RequestsSSLError as exc:
        if settings.egisz_ssl_allow_insecure_fallback and verify_value is not False:
            upstream = requests.request(
                request.method,
                target_url,
                headers=incoming_headers,
                data=body,
                allow_redirects=False,
                timeout=60,
                verify=False,
            )
            used_insecure_ssl_fallback = True
        else:
            raise HTTPException(
                status_code=502,
                detail=(
                    "Ошибка SSL при подключении к внешнему дашборду. "
                    "Укажите доверенный CA в settings.egisz_ca_bundle "
                    "или временно отключите проверку сертификата settings.egisz_ssl_verify=false. "
                    f"Детали: {exc}"
                ),
            )
    except RequestException as exc:
        raise HTTPException(status_code=502, detail=f"Ошибка запроса к внешнему дашборду: {exc}")

    content = upstream.content
    content_type = upstream.headers.get("content-type", "")

    if "text/html" in content_type:
        html = content.decode("utf-8", errors="ignore")
        html = _rewrite_html_for_proxy(html)
        if inject_autostart:
            html = _inject_autostart_script(html)
        content = html.encode("utf-8")

    excluded = {"content-length", "transfer-encoding", "connection", "content-encoding"}
    headers = {k: v for k, v in upstream.headers.items() if k.lower() not in excluded and k.lower() != "set-cookie"}

    if "location" in upstream.headers:
        headers["location"] = _rewrite_location_header(upstream.headers["location"])

    response = Response(content=content, status_code=upstream.status_code, headers=headers)

    if used_insecure_ssl_fallback:
        response.headers.setdefault("x-egisz-ssl-warning", "insecure-fallback-used")

    cookies = upstream.raw.headers.get_all("Set-Cookie") if hasattr(upstream.raw.headers, "get_all") else []
    for cookie in cookies:
        response.headers.append("set-cookie", cookie.replace("Domain=info-bi-db.egisz.rosminzdrav.ru;", ""))

    return response


@app.api_route("/egisz/{full_path:path}", methods=["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS", "HEAD"])
async def egisz_proxy(full_path: str, request: Request):
    """Прокси внешнего дашборда для работы без браузерного расширения"""
    return await _proxy_to_external(request, full_path, inject_autostart=("dashboardsViewer" in full_path))


@app.api_route("/{prefix}/{full_path:path}", methods=["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS", "HEAD"])
async def external_prefix_proxy(prefix: str, full_path: str, request: Request):
    """Проксирование ключевых путей дашборда без префикса /egisz."""
    allowed_prefixes = {"corelogic", "idsrv", "api", "assets", "content", "scripts", "styles"}
    if prefix not in allowed_prefixes:
        raise HTTPException(status_code=404, detail="Not found")
    target_path = f"{prefix}/{full_path}" if full_path else prefix
    return await _proxy_to_external(request, target_path, inject_autostart=False)




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


@app.get("/{full_path:path}")
async def external_fallback_proxy(full_path: str, request: Request):
    """Fallback для корневых ресурсов внешнего SPA, которые не покрыты локальными роутами."""
    if full_path.startswith("egisz/"):
        raise HTTPException(status_code=404, detail="Not found")

    accept = request.headers.get("accept", "")
    is_static_like = "." in full_path or any(x in accept for x in ["text/css", "application/javascript", "image/", "font/"])
    if not is_static_like:
        raise HTTPException(status_code=404, detail="Not found")

    return await _proxy_to_external(request, full_path, inject_autostart=False)
