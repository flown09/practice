# main.py
import time
import os
import re
from datetime import datetime
from fastapi import FastAPI, BackgroundTasks, HTTPException, Form, Request
from fastapi.responses import FileResponse, HTMLResponse, Response
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles

from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

from utils import main_process, get_last_week_dates
from config import settings

import httpx
from urllib.parse import urlencode

app = FastAPI(title="MIAC Report Service")
templates = Jinja2Templates(directory="templates")

# --- СТАТИКА (для инжекта auto_report.js) ---
app.mount("/static", StaticFiles(directory="static"), name="static")

schedule_config = {
    "enabled": False,
    "day_of_week": 0,
    "hour": 10,
    "minute": 0
}

scheduler = None

def create_scheduler():
    global scheduler
    return BackgroundScheduler()

def scheduled_report():
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Автовыгрузка запущена")
    try:
        duration = main_process()
        print(f"[{(datetime.now().strftime('%H:%M:%S'))}] Автовыгрузка завершена за {duration:.2f}с")
    except Exception as e:
        print(f"[{(datetime.now().strftime('%H:%M:%S'))}] Ошибка автозагрузки: {e}")

def run_miac_direct():
    print(f"[{(datetime.now().strftime('%H:%M:%S'))}] Ручная выгрузка запущена")
    try:
        duration = main_process()
        print(f"[{(datetime.now().strftime('%H:%M:%S'))}] Обработка завершена за {duration:.2f} секунд")
    except Exception as e:
        print(f"[{(datetime.now().strftime('%H:%M:%S'))}] Ошибка обработки: {e}")

@app.on_event("startup")
async def startup():
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
    global scheduler, schedule_config

    if scheduler is not None:
        if scheduler.running:
            scheduler.shutdown(wait=True)
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
    return templates.TemplateResponse("schedule.html", {
        "request": request,
        "config": schedule_config
    })

@app.post("/start-report", response_model=dict)
async def start_report(background_tasks: BackgroundTasks):
    background_tasks.add_task(run_miac_direct)
    task_id = f"direct_{int(time.time())}"
    return {"task_id": task_id, "status": "started", "time": datetime.now().strftime("%H:%M:%S")}

@app.get("/manual")
async def manual_report(background_tasks: BackgroundTasks):
    background_tasks.add_task(run_miac_direct)
    return {"status": "started", "time": datetime.now().strftime("%H:%M:%S")}

@app.get("/download-report")
async def download_report():
    if not os.path.exists(settings.excel_path):
        raise HTTPException(404, "Отчет не готов. Запустите выгрузку!")
    return FileResponse(
        settings.excel_path,
        filename="миац_отчет.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.get("/last-week-dates")
async def dates():
    return {"from": get_last_week_dates()[0], "to": get_last_week_dates()[1]}

@app.get("/docs")
async def docs_redirect():
    return {"docs": "/docs", "ui": "http://localhost:8000/docs"}


# ======================================================================
#  ПРОКСИ ДЛЯ info-bi-db.egisz.rosminzdrav.ru + ИНЖЕКТ auto_report.js
# ======================================================================

UPSTREAM = "https://info-bi-db.egisz.rosminzdrav.ru"
PROXY_PREFIX = "/bi"

def _rewrite_set_cookie(cookie_value: str) -> str:
    # убираем Domain=..., чтобы cookie применялась к вашему домену
    parts = cookie_value.split(";")
    out = []
    for p in parts:
        ps = p.strip()
        if ps.lower().startswith("domain="):
            continue
        out.append(p)
    return ";".join(out)

def _rewrite_location(loc: str) -> str:
    if loc.startswith(UPSTREAM):
        return PROXY_PREFIX + loc[len(UPSTREAM):]
    if loc.startswith("/"):
        return PROXY_PREFIX + loc
    return loc  # внешние редиректы (Госуслуги) не трогаем

def _should_inject(path: str, content_type: str) -> bool:
    return "text/html" in (content_type or "").lower()



def _inject_script_and_base(html: str) -> str:
    """
    1) Добавляем <base href="/bi/"> чтобы /dist/... и /fonts/... начали ходить через /bi/
    2) Добавляем наш скрипт auto_report.js
    """
    base_tag = f'<base href="{PROXY_PREFIX}/">'
    script_tag = '\n<script src="/static/auto_report.js"></script>\n'

    # Вставляем base в <head>
    if "<head" in html.lower():
        # вставим сразу после открывающего <head ...>
        html = re.sub(r"(<head[^>]*>)", r"\1\n" + base_tag + "\n", html, flags=re.IGNORECASE, count=1)
    else:
        # если вдруг head нет — добавим в начало
        html = base_tag + "\n" + html

    # Скрипт — перед </body> если возможно
    if re.search(r"</body>", html, flags=re.IGNORECASE):
        html = re.sub(r"</body>", script_tag + "</body>", html, flags=re.IGNORECASE, count=1)
    else:
        html += script_tag

    return html


def _rewrite_body_text(body: str) -> str:
    """
    Переписываем:
    - абсолютные URL вида https://info-bi-db... -> /bi/...
    - атрибуты src="/..." href="/..." -> src="/bi/..." href="/bi/..."
    - CSS url(/...) -> url(/bi/...)
    """
    # 1) абсолютные ссылки на апстрим
    body = body.replace(UPSTREAM, PROXY_PREFIX)

    # 2) src="/..." или href="/..." -> "/bi/..."
    # важно: не трогать уже /bi/
    body = re.sub(r'''(\b(?:src|href)=["'])/(?!bi/)([^"']+)''',
                  rf"\1{PROXY_PREFIX}/\2",
                  body, flags=re.IGNORECASE)

    # 3) CSS: url(/something) -> url(/bi/something)
    body = re.sub(r'''url\(\s*/(?!bi/)([^)]+)\)''',
                  rf"url({PROXY_PREFIX}/\1)",
                  body, flags=re.IGNORECASE)

    return body

from fastapi.responses import RedirectResponse

@app.get("/dist/{path:path}")
async def dist_redirect(path: str, request: Request):
    # /dist/... -> /bi/dist/...
    qs = f"?{request.url.query}" if request.url.query else ""
    return RedirectResponse(url=f"{PROXY_PREFIX}/dist/{path}{qs}", status_code=307)

@app.get("/fonts/{path:path}")
async def fonts_redirect(path: str, request: Request):
    qs = f"?{request.url.query}" if request.url.query else ""
    return RedirectResponse(url=f"{PROXY_PREFIX}/fonts/{path}{qs}", status_code=307)

from fastapi.responses import RedirectResponse

def _with_qs(request: Request, url: str) -> str:
    return url + (f"?{request.url.query}" if request.url.query else "")

@app.api_route("/admin/{path:path}", methods=["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS", "HEAD"])
async def admin_redirect(path: str, request: Request):
    return RedirectResponse(_with_qs(request, f"{PROXY_PREFIX}/admin/{path}"), status_code=307)

@app.api_route("/idsrv/{path:path}", methods=["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS", "HEAD"])
async def idsrv_redirect(path: str, request: Request):
    return RedirectResponse(_with_qs(request, f"{PROXY_PREFIX}/idsrv/{path}"), status_code=307)

# на всякий случай — BI API часто ходит сюда
@app.api_route("/corelogic/{path:path}", methods=["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS", "HEAD"])
async def corelogic_redirect(path: str, request: Request):
    return RedirectResponse(_with_qs(request, f"{PROXY_PREFIX}/corelogic/{path}"), status_code=307)


@app.get("/favicon.ico")
async def favicon_redirect():
    return RedirectResponse(url=f"{PROXY_PREFIX}/dist/favicon.ico", status_code=307)



@app.api_route(f"{PROXY_PREFIX}" + "/{path:path}", methods=["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS", "HEAD"])
async def bi_proxy(path: str, request: Request):
    upstream_url = f"{UPSTREAM}/{path}"
    if request.query_params:
        upstream_url += "?" + urlencode(list(request.query_params.multi_items()))

    # --- headers ---
    headers = dict(request.headers)
    headers.pop("host", None)
    headers["accept-encoding"] = "identity"  # чтобы можно было менять HTML

    # referer/origin иногда важно
    if "referer" in headers:
        headers["referer"] = headers["referer"].replace(str(request.base_url).rstrip("/"), UPSTREAM).replace(PROXY_PREFIX, "")
    if "origin" in headers:
        headers["origin"] = UPSTREAM

    body = await request.body()

    async with httpx.AsyncClient(
        follow_redirects=False,
        timeout=60.0,
        verify=False,   # ваш вариант A
    ) as client:
        resp = await client.request(
            method=request.method,
            url=upstream_url,
            headers=headers,
            content=body,
            cookies=request.cookies
        )

    content_type = resp.headers.get("content-type", "")

    # --- соберём headers ответа, исключая проблемные ---
    hop_by_hop = {
        "content-encoding", "transfer-encoding", "connection", "keep-alive",
        "proxy-authenticate", "proxy-authorization", "te", "trailers",
        "upgrade"
    }

    out_headers = {}
    for k, v in resp.headers.items():
        lk = k.lower()
        if lk in hop_by_hop:
            continue
        if lk == "content-length":
            # КРИТИЧНО: мы можем менять тело, поэтому длину не прокидываем
            continue
        if lk == "set-cookie":
            # добавим ниже корректно, как множественные заголовки
            continue
        if lk == "location":
            out_headers[k] = _rewrite_location(v)
            continue

        out_headers[k] = v

    # --- тело ответа ---
    is_text = any(t in content_type.lower() for t in [
        "text/html", "application/javascript", "text/javascript", "text/css", "application/json"
    ])

    if is_text:
        text = resp.text
        text = _rewrite_body_text(text)
        if _should_inject(path, content_type):
            text = _inject_script_and_base(text)

        response = Response(
            content=text,
            status_code=resp.status_code,
            headers=out_headers,
            media_type=content_type.split(";")[0] if content_type else None
        )
    else:
        response = Response(
            content=resp.content,
            status_code=resp.status_code,
            headers=out_headers,
            media_type=content_type.split(";")[0] if content_type else None
        )

    # --- Set-Cookie: добавляем как множественные заголовки (ВАЖНО) ---
    # Нельзя join через запятую: cookies могут содержать запятые в Expires=...
    for c in resp.headers.get_list("set-cookie"):
        response.headers.append("set-cookie", _rewrite_set_cookie(c))

    response.headers["x-miac-proxy"] = "1"
    print("PROXY:", request.method, request.url.path, "->", upstream_url, "CT:", content_type, "STATUS:",
          resp.status_code)

    return response