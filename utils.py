import time
import requests
import re
import json
import xml.etree.ElementTree as ET
from datetime import datetime
import pandas as pd
import warnings
from isoweek import Week
from openpyxl import load_workbook
from config import settings

warnings.filterwarnings('ignore')
cancelLPU = pd.DataFrame()


def get_last_week_dates():
    date = datetime.today()
    lol = date.isocalendar()
    return Week(date.year, (lol[1] - 1)).monday().strftime("%d.%m.%Y"), Week(date.year, (lol[1] - 1)).sunday().strftime(
        "%d.%m.%Y")


def process_lpu(lpu_df: pd.DataFrame, ws, wb):
    BASE_URL = settings.base_url
    LOGIN = settings.login
    PASSWORD = settings.password
    REPORT_ID = settings.report_id
    EMPLOYER = settings.employer
    DATE_FROM, DATE_TO = get_last_week_dates()

    session = requests.Session()

    # ----------------------------
    # 0. Получение SYS_CACHE_UID
    # ----------------------------
    url_js = f"{BASE_URL}/d3/~d3api"
    resp_js = session.get(url_js)
    resp_js.raise_for_status()

    match = re.search(r'D3Api\.SYS_CACHE_UID\s*=\s*"([a-f0-9]+)"', resp_js.text)
    if not match:
        raise ValueError("Не удалось найти SYS_CACHE_UID")
    SYS_CACHE_UID = match.group(1)

    print(f"[INFO] SYS_CACHE_UID = {SYS_CACHE_UID}")

    # ----------------------------
    # 1. Авторизация
    # ----------------------------
    url_login = f"{BASE_URL}/d3/request.php"
    params_login = {
        "cache_enabled": "1",
        "session_cache": "1",
        "modal": "1",
        "Form": "System/login",
        "cache": SYS_CACHE_UID
    }
    data_login = {
        "request": json.dumps({
            "Authorization": {
                "type": "Module",
                "params": {
                    "DBPassword": PASSWORD,
                    "DBLogin": LOGIN
                }
            }
        })
    }
    resp_login = session.post(url_login, params=params_login, data=data_login)
    resp_login.raise_for_status()
    print("[INFO] Авторизация прошла успешно")

    # --- Берем Bearer токен ---
    login_json = resp_login.json()
    bearer_token = login_json.get("Authorization", {}).get("data", {}).get("access_token")
    if bearer_token:
        session.headers.update({"Authorization": f"Bearer {bearer_token}"})
        print(f"[INFO] Bearer token установлен")
    else:
        print("[WARN] Bearer токен не найден в ответе")

    # ----------------------------
    # 2. Выбор LPU и получение данных
    # ----------------------------
    for i in range(len(lpu_df)):  # Используем lpu_df вместо LPU
        LPU_ROW = lpu_df.iloc[i]  # Текущая строка LPU

        url_lpu = f"{BASE_URL}/d3/request.php"
        params_lpu = {
            "cache_enabled": "1",
            "session_cache": "1",
            "modal": "1",
            "Form": "System/lpu",
            "cache": SYS_CACHE_UID
        }
        data_lpu = {
            "request": json.dumps({
                "Authorization": {
                    "type": "Module",
                    "params": {
                        "LPU": str(LPU_ROW["РМИС ID"]),
                        "CABLAB": "",
                        "EMPLOYER": EMPLOYER
                    }
                }
            })
        }
        resp_lpu = session.post(url_lpu, params=params_lpu, data=data_lpu)
        resp_lpu.raise_for_status()
        print(f"[INFO] LPU установлено {LPU_ROW['РМИС ID']}")

        # ----------------------------
        # 3. Получение FormCache
        # ----------------------------
        url_report_form = f"{BASE_URL}/getform.php"
        params_report_form = {
            "Form": "Reports/run",
            "modulename": "7465547992_report",
            "_rep_code": "lpu_talons_pmsp",
            "_exp_id": 0,
            "_fme_doc_id": 0,
            "modal": 1,
            "theme": "bars",
            "cache": SYS_CACHE_UID,
            "cache_enabled": 1,
            "session_cache": 1
        }
        resp_form = session.post(url_report_form, params=params_report_form)
        resp_form.raise_for_status()

        form_cache = resp_form.headers.get("FormCache")
        if not form_cache:
            raise ValueError("Не удалось получить FormCache")
        print(f"[INFO] FormCache = {form_cache}")

        # ----------------------------
        # 4. Получение данных отчёта (XML)
        # ----------------------------
        url_multidata = f"{BASE_URL}/getmultidata.php"
        params_multidata = {
            "Form": "Reports/Statistic/lpu_talons",
            "reportid": REPORT_ID,
            "nooverflow": "true",
            "theme": "bars",
            "cache": SYS_CACHE_UID,
            "FormCache": form_cache
        }
        data_multidata = {
            "_sysid": "1",
            "ds0[Form]": "Reports/Statistic/lpu_talons",
            "ds0[nooverflow]": "true",
            "ds0[DataSet]": "DS_TALONS",
            "ds0[DATE_FROM_g0]": DATE_FROM,
            "ds0[DATE_TO_g1]": DATE_TO,
            "ds0[CMP_REG_TYPE_g2]": 4,
            "ds0[EXCLUDE_TIME_TYPE_g4]": "2;6;27;39;41;96;116;136;149;420;421;422;423;424;74;77;155;400;401;402;403;75;17",
            "ds0[INCLUDE_SPECS_g6]": "1177;70;84;88;159;160;161;162;1179;61;123;117;1281;1282;1291;1292;1293;1294;1295;1300;5003;2155;2154;2153;9015;2489;2493;2499;2504;2511;2854;3104;3117;3127;3065;3077;3079;2501;2858;2860;2991;3096;3100;3111;3115;3124;3125;3132;3078;9014;2851;2861;3093;3105;3108;2498;2510;2815;2859;999;2976;2992;3106;2490;2492;2512;2977;3107;3139;2502;2509;19;2944;2978;3123;3126;3086;3075;2520;2855;2979;3094;3095;3099;3074;3076;78;5004;1372;87;2500;125;90;133;1056;5000;5001;203;2011;201;2103;2104;2515;2513",
            "ds0[_sysid]": "0"
        }
        try:
            resp_data = session.post(url_multidata, params=params_multidata, data=data_multidata)
            resp_data.raise_for_status()
            df = pd.read_xml(
                resp_data.text.replace('<DataSet name="DS_TALONS" sysid="0">', "").replace('</DataSet>', ""))

            for k in range(3, 94):
                if lpu_df["РМИС ID"][i] == ws[f'B{k}'].value:
                    row_index = i + 4
                    ws[f'C{row_index}'] = df['PORTAL_ALL'].sum()
                    ws[f'D{row_index}'] = df['TOTAL'].sum()
                    ws[f'E{row_index}'] = df['TOTAL_ALL'].sum()
                    ws[f'F{row_index}'] = datetime.today()
                    break

            wb.save(settings.excel_path)
            print(f"[INFO] Сохранено для LPU {lpu_df['РМИС ID'][i]}")

        except Exception as e:
            print(f"[ERROR] Ошибка для LPU {lpu_df['РМИС ID'][i]}: {e}")
            if cancelLPU.empty:
                cancelLPU = pd.DataFrame()
            new_row = pd.DataFrame({
                "Наменование МО": [lpu_df["Наменование МО"][i]],
                "РМИС ID": [lpu_df["РМИС ID"][i]]
            })
            cancelLPU = pd.concat([cancelLPU, new_row], ignore_index=True)


def main_process():
    LPU = pd.read_excel(settings.lpu_path)
    wb = load_workbook(settings.excel_path)
    ws = wb["Лист 1"]

    start_time = time.time()

    if not cancelLPU.empty:
        print("[INFO] Обработка неудачных LPU...")
        process_lpu(cancelLPU.drop_duplicates(), ws, wb)

    print("[INFO] Основная обработка...")
    process_lpu(LPU, ws, wb)

    duration = time.time() - start_time
    print(f"[INFO] Обработка завершена за {duration:.2f} сек")
    return duration
