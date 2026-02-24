# MIAC Report Service

Подробная документация по сервису формирования отчётов МИАЦ на базе FastAPI.

## 1. Назначение проекта

Сервис решает две основные задачи:

1. **Выгрузка талонов** по списку МО (LPU) из внешней системы и запись агрегатов в `miac_table.xlsx`.
2. **Сборка финального Excel-отчёта** из `parse.json` + шаблона Excel + справочников LPU/LPU2.

Приложение предоставляет HTTP API и простую web-страницу (`/`) для ручного запуска и настройки расписания.

## 2. Ключевые возможности

- API для загрузки шаблона `.xlsx` и проверки его состояния.
- API для загрузки `parse.json` и генерации итогового файла `final_<id>.xlsx`.
- Плановый запуск фоновой выгрузки талонов через APScheduler.
- Скачивание готового отчёта и ZIP-архива browser extension из папки `extension/`.

## 3. Структура проекта

```text
/workspace/practice
├── main.py                 # FastAPI приложение и маршруты
├── utils.py                # Основная бизнес-логика обработки Excel/JSON и выгрузки талонов
├── config.py               # Настройки (Pydantic BaseSettings)
├── tasks.py                # Celery task (альтернативный запуск main_process)
├── templates/
│   └── schedule.html       # UI-страница управления
├── extension/              # Файлы расширения для скачивания
├── reports/                # Генерируемые файлы (final_*.xlsx, last_parse.json)
├── requirements.txt
├── LPU.xlsx                # Справочник МО для выгрузки талонов
├── LPU2.xlsx               # Справочник МО для финального отчёта
├── miac_table.xlsx         # Промежуточная таблица талонов (Q/R/S)
└── sample.xlsx             # Excel-шаблон (с листом "Лист-шаблон")
```

## 4. Архитектура и поток данных

### 4.1 Поток A: загрузка parse.json → финальный Excel

1. Клиент отправляет `POST /upload-parse` с файлом JSON.
2. В threadpool вызывается `build_final_excel_from_parse_bytes(...)`.
3. Обработчик:
   - читает метрики из JSON;
   - сопоставляет МО через `LPU2.xlsx` (и fallback через `LPU.xlsx` + RMIS ID);
   - подтягивает талоны Q/R/S из `miac_table.xlsx`;
   - копирует лист `Лист-шаблон` в новый лист с именем прошлой ISO-недели;
   - заполняет колонки D/F/I/K/M/O, Q/R/S и формулы C/E/G/H/J/L/N/P/T;
   - сохраняет итог в `reports/final_<upload_id>.xlsx`.
4. Клиент скачивает файл по `GET /download-final/{upload_id}`.

### 4.2 Поток B: выгрузка талонов

1. Плановый запуск: настройка через `POST /schedule`.
2. Вызывается `main_process()`:
   - читает `LPU.xlsx`;
   - авторизуется во внешней системе;
   - по каждой МО получает статистику талонов;
   - сохраняет агрегаты в `miac_table.xlsx` (лист `Лист 1`).

## 5. Требования

- Python 3.10+.
- Доступ к внешнему хосту, указанному в `base_url` (для выгрузки талонов).
- Входные Excel/JSON-файлы в ожидаемом формате.

## 6. Установка и запуск

### 6.1 Установка зависимостей

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

> Примечание: `requirements.txt` сохранён в UTF-16 LE. Если установщик не читает файл корректно, перекодируйте его в UTF-8.

### 6.2 Запуск сервера

```bash
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

## 7. Конфигурация

Настройки задаются в `config.py` через `BaseSettings` (могут переопределяться переменными окружения).

Основные параметры:

- `base_url` — базовый URL внешней системы.
- `login` / `password` — учётные данные.
- `report_id` / `employer` — параметры отчётного запроса.
- `excel_path` — путь к Excel-шаблону (`sample.xlsx`).
- `excel_path_ticket` — путь к промежуточной таблице талонов (`miac_table.xlsx`).
- `lpu_path` — путь к `LPU.xlsx`.
- `lpu2_path` — путь к `LPU2.xlsx`.

## 8. Форматы входных файлов

### 8.1 Шаблон Excel (`sample.xlsx`)

Обязателен лист **`Лист-шаблон`**. При загрузке через `/upload-template` сервис проверяет:

- расширение `.xlsx`;
- что файл не пустой;
- что лист `Лист-шаблон` присутствует.

### 8.2 `parse.json`

Сервис ожидает структуру, где:

- в первой колонке — название МО;
- в колонках 1..6 — массивы вида `[..., [число]]`.

Метрики раскладываются так:

- col1 → `chern` (D)
- col2 → `osh_all` (I)
- col3 → `org_osh` (K)
- col4 → `teh_osh` (M)
- col5 → `usp` (F)
- col6 → `fed` (O)

### 8.3 `LPU.xlsx` и `LPU2.xlsx`

Поддерживаются вариации имён колонок, например:

- название МО: `Наменование МО`, `Наименование МО`, `МО`, `name`, `NAME`;
- код МО: `РМИС ID`, `Код МО`, `RMIS ID`, `Код`, `CODE`.

### 8.4 `miac_table.xlsx`

Ожидается лист **`Лист 1`** и структура:

- A: название МО
- B: РМИС ID
- C: `PORTAL_ALL`
- D: `TOTAL`
- E: `TOTAL_ALL`

## 9. API (подробно)

### 9.1 Шаблон

- `GET /template-info` — метаданные текущего шаблона.
- `GET /download-template` — скачать текущий шаблон.
- `POST /upload-template` — загрузить/заменить шаблон.

### 9.2 Parse JSON и итоговый отчёт

- `POST /upload-parse` — загрузка `parse.json` + генерация `final_<id>.xlsx`.
- `POST /upload-parse-raw` — просто сохранить сырой JSON как `reports/last_parse.json`.
- `GET /download-final/{upload_id}` — скачать результат сборки.
- `GET /download-last-parse` — скачать последний JSON.
- `GET /view-last-parse` — посмотреть JSON в браузере.


### 9.3 Планировщик и запуск

- `POST /schedule` (form-data):
  - `enabled` (`true/false`)
  - `day` (0..6, где 0 = понедельник)
  - `hour` (0..23)
  - `minute` (0..59)
- `GET /status` — текущее состояние scheduler + состояние файлов.
- `POST /start-report` — ручной запуск в фоне (API).
- `GET /manual` — ручной запуск (для UI-кнопки).

Пример:

```bash
curl -X POST http://localhost:8000/schedule \
  -F "enabled=true" \
  -F "day=0" \
  -F "hour=10" \
  -F "minute=0"
```

### 9.4 Прочие endpoint’ы

- `GET /download-report` — скачать файл по `settings.excel_path`.
- `GET /download-extension` — скачать ZIP содержимого `extension/`.
- `GET /last-week-dates` — даты прошлой недели.
- `GET /docs` — возвращает ссылку на Swagger UI.

## 10. Логирование и диагностика

В текущей реализации используются `print(...)` в stdout:

- запуск/завершение ручной и плановой задач;
- этапы авторизации/выгрузки;
- ошибки по отдельным МО.

Рекомендуется запускать через процесс-менеджер (systemd/supervisor/docker) с ротацией логов.

## 11. Частые проблемы

1. **`В шаблоне нет листа 'Лист-шаблон'`**
   - Загрузите корректный `.xlsx` через `/upload-template`.

2. **`last_parse.json не найден`**
   - Сначала выполните `POST /upload-parse` или `POST /upload-parse-raw`.

3. **Ошибка доступа к внешнему API**
   - Проверьте `base_url`, сеть, логин/пароль, `report_id`, `employer`.

4. **Неполное сопоставление МО**
   - Проверьте консистентность `LPU.xlsx`, `LPU2.xlsx`, названий МО и RMIS ID.

## 12. Эксплуатационные рекомендации

- Храните шаблон, LPU и отчётные файлы в backup.
- Ограничьте доступ к API (как минимум reverse proxy + auth).
- Вынесите секреты (`login`, `password`) в переменные окружения.
- Для production запускайте без `--reload`.

## 13. Быстрый сценарий использования

1. Запустите сервис.
2. Проверьте шаблон: `GET /template-info`.
3. При необходимости загрузите новый шаблон: `POST /upload-template`.
4. Запустите выгрузку талонов (`/manual` или `/start-report`).
5. Загрузите `parse.json` через `/upload-parse`.
6. Скачайте результат по `download_url` из ответа.

---

Если нужно, могу дополнительно подготовить:

- OpenAPI-клиент (готовые примеры на Python/JS),
- docker-compose для продакшн-развёртывания,
- отдельный `OPERATIONS.md` (резервное копирование, мониторинг, recovery).
