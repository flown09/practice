# MIAC Report Service

Подробная документация по сервису формирования отчётов МИАЦ на базе FastAPI.

## 1. Назначение проекта
Проект автоматизирует формирование еженедельного отчёта по попыткам записи/ошибкам/успешным записям:

1. Сбор данных по талонам из MIAC.
   Скрипт (utils.main_process) авторизуется в системе MIAC по HTTP, по списку МО из LPU.xlsx получает данные отчёта (XML), суммирует и записывает результат в Excel-файл miac_table.xlsx.
2. Сбор агрегатов с веб-дашборда (ЕГИСЗ).
   Chrome-расширение открывает дашборд, берёт токен из sessionStorage, делает серию запросов в API дашборда, собирает parse.json и загружает его на FastAPI-сервис.
3. Сборка финального Excel.
   Серверный метод build_final_excel_from_parse_bytes() берёт:
   - parse.json (из расширения),
   - LPU.xlsx (справочник МО),
   - miac_table.xlsx (таблица с Q/R/S по талонам),
   - Excel-шаблон sample.xlsx
   …и формирует финальный Отчет.xlsx
4. Веб-страница управления (schedule.html).
   Позволяет:
   - включить/выключить автозапуск MIAC-выгрузки по расписанию (APScheduler),
   - загрузить/скачать Excel-шаблон,
   - скачать архив расширения.

## 2. Архитектура и поток данных

### 2.1 Компоненты

- FastAPI приложение (main.py)
  - UI: / (страница schedule.html)
  - API: расписание, загрузка шаблона, генерация финального Excel, скачивания файлов
  - Планировщик: APScheduler (фоновые задания)
- Модуль настроек (config.py)
  - параметры подключения к MIAC (URL, логин, пароль и пр.)
  - пути к файлам: шаблон, таблица MIAC, справочник LPU
- Сбор данных из MIAC (utils.py → main_process(), process_lpu())
  - авторизация, выбор LPU, получение FormCache, запрос XML, запись в miac_table.xlsx
- Сборка финального Excel из parse.json (utils.py → build_final_excel_from_parse_bytes())
  - заполнение шаблона
  - формулы/проценты/итоги
  - добавление колонок Q/R/S из miac_table.xlsx
- Chrome-расширение (extension/)
  - sd.js — выполняется в контексте страницы, собирает данные с дашборда
  - script-loader.js — контент-скрипт, инжектит sd.js, мостит сообщения в background
  - background.js — отправляет parse.json на FastAPI /upload-parse и (опционально) скачивает итоговый Excel
  - manifest.json — MV3 конфигурация
- Celery задачи (tasks.py)
  - заготовка для фонового запуска main_process() через Celery
  - в текущем коде не используется FastAPI приложением

### 2.2 Поток данных (схема)

1. Пользователь открывает дашборд (сайт ЕГИСЗ), расширение автоматически запускается.
2. sd.js:
   - ждёт токен (sessionStorage oidc.user:/idsrv:DashboardsApp)
   - считывает выбранный диапазон дат из iframe
   - получает список организаций (ORG)
   - по каждой ORG выполняет пачку запросов по метрикам (черновики/ошибки/успешные/и т.п.)
   - формирует массив resultAll и отправляет на страницу сообщение UPLOAD_PARSE
3. script-loader.js ловит UPLOAD_PARSE, пересылает в background.js
4. background.js POST-ит файл parse.json на http://<server>:8000/upload-parse
5. FastAPI /upload-parse:
   - вызывает build_final_excel_from_parse_bytes(...)
   - сохраняет reports/final_<uuid>.xlsx
   - возвращает ссылку /download-final/<uuid>
6. background.js при желании сразу скачивает файл (через chrome.downloads).

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

## 4. Требования и зависимости

### 4.1 Python

Рекомендуется Python 3.10+.

### 4.2 Библиотеки

Минимально нужны:
- fastapi, uvicorn
- jinja2
- apscheduler
- pydantic-settings
- requests
- pandas
- openpyxl
- isoweek

## 5. Конфигурация (config.py)

Settings(BaseSettings) поддерживает конфигурацию через переменные окружения.

Параметры:
- base_url — URL MIAC/D3 (пример: http://.../demo/)
- login, password — учётка для MIAC
- report_id, employer — параметры отчётной формы MIAC
- excel_path — путь к Excel-шаблону (по коду это ещё и “главный файл” для скачивания в /download-report)
- excel_path_ticket — путь к miac_table.xlsx (таблица талонов Q/R/S)
- lpu_path — путь к LPU.xlsx

Рекомендация: вынести секреты в .env и не хранить в Git.

## 6. FastAPI сервис (main.py)

### 6.1 Веб-интерфейс

GET /
Отдаёт страницу schedule.html (Jinja2), где можно:
- включить автозапуск (APScheduler)
- выбрать день недели и время
- загрузить Excel-шаблон
- скачать текущий шаблон
- скачать расширение
- открыть дашборд

### 6.2 API эндпоинты (полный список)

**Шаблон Excel**

GET /template-info

Возвращает информацию о текущем шаблоне (settings.excel_path):
- существует ли файл
- имя/размер/время изменения
- названия листов (пытается открыть openpyxl)

GET /download-template

Скачивает текущий шаблон .xlsx.

POST /upload-template

Загружает новый Excel-шаблон:
- проверяет расширение .xlsx
- проверяет, что файл реально открывается
- проверяет, что есть лист "Лист-шаблон" (важно для сборки отчёта)
- сохраняет атомарно (через временный файл) под settings.excel_path
- защищено TEMPLATE_LOCK от гонок

Генерация финального отчёта из parse.json
POST /upload-parse

Принимает файл parse.json (обязателен Content-Type: multipart/form-data):
- создаёт upload_id
- строит reports/final_<upload_id>.xlsx
- возвращает ссылку на скачивание

GET /status

Возвращает:
- включено ли расписание
- запущен ли scheduler
- строка следующего запуска (условная)
- число jobs
- размер/дата файла settings.excel_path
- даты прошлой недели (get_last_week_dates())

GET /download-extension

Архивирует папку extension/ в zip и отдаёт как Dashbord-extension.zip.

GET /last-week-dates

Возвращает даты прошлой недели (для отладки).

GET /docs

Возвращает JSON-подсказку (не редирект), где swagger.

## 7. Логика MIAC выгрузки (utils.py)

### 7.1 main_process()

Назначение: обновить файл miac_table.xlsx (колонки Q/R/S в финальном отчёте).

Шаги:

1. Загружает LPU.xlsx в DataFrame.
2. Нормализует названия колонок (удаляет NBSP, strip).
3. Пытается обеспечить наличие колонки "РМИС ID" (переименованием альтернатив: "Код МО", "RMIS ID", "Код", "CODE").
4. Открывает miac_table.xlsx → лист "Лист 1".
5. Если cancelLPU не пустой — повторяет обработку неудачных МО.
6. Запускает process_lpu(LPU, ...) — основная обработка.
7. Сохраняет miac_table.xlsx, возвращает длительность.

### 7.2 process_lpu(lpu_df, ws, wb, cancel_lpu=None)

На каждую МО:

1. Получает SYS_CACHE_UID из GET {BASE_URL}/d3/~d3api (regex по JS).
2. Авторизуется POST-ом в request.php (System/login), получает Bearer token (если есть), кладёт в headers.
3. Выбирает LPU (System/lpu) с параметрами LPU/EMPLOYER.
4. Получает FormCache заголовок у getform.php (Reports/run).
5. Запрашивает getmultidata.php → получает XML, чистит обёртку <DataSet...>, парсит pandas.read_xml.
6. Ищет строку в листе "Лист 1", где B{k} == РМИС ID, записывает суммы:
   - C = PORTAL_ALL.sum()
   - D = TOTAL.sum()
   - E = TOTAL_ALL.sum()
   - F = текущая дата/время
7. При ошибке добавляет МО в cancelLPU.

При ошибке добавляет МО в cancelLPU.

## 8. Генерация финального отчёта (build_final_excel_from_parse_bytes)

### 8.1 Входные данные

Функция принимает:
- parse_bytes — содержимое parse.json
- lpu_path — путь к LPU.xlsx
- template_xlsx_path — путь к шаблону (должен содержать лист "Лист-шаблон")
- output_xlsx_path — куда сохранить результат
- sheet_name — имя листа-шаблона (по умолчанию "Лист-шаблон")

Дополнительно читает settings.excel_path_ticket (обычно miac_table.xlsx) для Q/R/S.

### 8.2 Формат parse.json (как ожидает ваш код)

Вы строите resultAll как массив, где каждая запись:

[
  "Название МО",
  ["ПОПЫТОК ...", [число]],
  ["НЕУДАЧНЫХ ...", [число]],
  ["Орг. ...", [число]],
  ["Тех. ...", [число]],
  ["УСПЕШНЫХ ...", [число]],
  ["ошибки федерального уровня", [число]]
]

Внутри build_final_excel_from_parse_bytes это читается через pandas.read_json, затем доступ идёт по индексам df.at[row, col], ожидается структура “список из 2 элементов”, где второй элемент — список со значением.

## 9. Веб-страница schedule.html

1. Настройки (collapsible)
   - Расписание (enabled/day/hour/minute)
   - Excel-шаблон:
     - загрузка /upload-template
     - скачивание /download-template
2. Расширение
   - скачать zip: /download-extension
   - открыть страницу расширений (Chrome/Edge/Opera/Firefox) — через helper JS
3. Открыть дашборд
   - ссылка на дашборд ЕГИСЗ

JS-логика:
- POST /schedule — сохранить расписание
- кнопка удаления расписания — отправляет “disabled + дефолтные значения”
- POST /upload-template — загрузка шаблона
- GET /template-info — обновление инфо (в коде элемент templateInfo сейчас не добавлен в HTML, но функция готова)

## 10. Chrome-расширение (extension/)

### 10.1 Как работает
- manifest.json (MV3):
  - content script script-loader.js запускается на https://info-bi-db.egisz.rosminzdrav.ru/*
  - делает web_accessible_resources для sd.js
  - background service worker: background.js
  - права: activeTab, downloads
  - host_permissions для вашего FastAPI сервера
- script-loader.js:
  - инжектит sd.js в контекст страницы
  - слушает window.postMessage и пересылает сообщения в background
- sd.js:
  - ждёт готовность: токен + iframe + поле дат
  - получает список ORG
  - для каждого ORG делает серию fetch в https://info-bi-db.egisz.rosminzdrav.ru/corelogic/api/query
  - собирает resultAll
  - window.postMessage({type:"UPLOAD_PARSE", text: JSON.stringify(resultAll)})
- background.js:
  - принимает UPLOAD_PARSE
  - отправляет parse.json на FastAPI /upload-parse
  - получает ссылку /download-final/<id>
  - (опционально) скачивает итоговый Excel

### 10.2 Важная настройка: адрес сервера

В background.js:

const APP_UPLOAD_URL = "http://172.20.220.254:8000/upload-parse";

## 11. Инструкция по установке и запуску

### 11.1 Запуск сервиса

uvicorn main:app --host 0.0.0.0 --port 8000

### 11.2 Установка расширения

1. В UI нажмите “Скачать расширение (.zip)”
2. Распакуйте архив
3. Откройте страницу расширений:
   - Chrome/Яндекс: chrome://extensions/
   - Edge: edge://extensions/
4. Включите “Режим разработчика”
5. “Загрузить распакованное” → выберите папку расширения
