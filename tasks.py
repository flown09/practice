from celery import Celery
from utils import main_process
import time

celery_app = Celery(
    "miac_service",
    broker="memory://",
    backend="cache+memory://",
    include=["tasks"]
)

@celery_app.task(bind=True, max_retries=3)
def run_miac_report(self):
    try:
        duration = main_process()
        return {"status": "success", "duration": duration}
    except Exception as exc:
        print(f"Ошибка в задаче: {exc}")
        raise self.retry(exc=exc, countdown=60)
