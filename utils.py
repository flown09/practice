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
import io
import warnings
import datetime
import pandas as pd
from openpyxl import load_workbook
from isoweek import Week

warnings.filterwarnings('ignore')
cancelLPU = pd.DataFrame()


def _clean_name(s: str) -> str:
    if s is None:
        return ""
    return ''.join(filter(str.isalnum, str(s))).lower()

def _pick_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

def nz(v):
    return 0 if v is None else v

def ensure_formula(ws, addr, formula):
    v = ws[addr].value
    if not (isinstance(v, str) and v.startswith("=")):
        ws[addr].value = formula

def build_final_excel_from_parse_bytes(
    parse_bytes: bytes,
    lpu_path: str,
    template_xlsx_path: str,
    output_xlsx_path: str,
    sheet_name: str = "Лист 1",
):
    # 1) читаем parse.json
    df = pd.read_json(io.BytesIO(parse_bytes))
    df = df.dropna().reset_index(drop=True)

    # 2) читаем lpu.xlsx
    lpu = pd.read_excel(lpu_path)

    name_col = _pick_col(lpu, ["name", "Наменование МО", "Наименование МО", "МО", "NAME"])
    code_col = _pick_col(lpu, ["Код", "Код МО", "РМИС ID", "RMIS ID", "CODE"])

    if not name_col:
        raise ValueError(f"В lpu.xlsx не найдена колонка с названием МО. Есть колонки: {list(lpu.columns)}")

    # код не обязателен (в твоём шаблоне ты его вообще не записываешь)
    lpu["__name__"] = lpu[name_col].apply(_clean_name)
    if code_col:
        lpu["__code__"] = lpu[code_col]
    else:
        lpu["__code__"] = None

    ticket_by_code, ticket_by_name = load_ticket_maps(settings.excel_path_ticket)

    # 3) чистим имена (как в твоём коде)
    df[0] = df[0].apply(_clean_name)
    lpu["__name__"] = lpu["__name__"].apply(_clean_name)

    # 4) собираем df2 (как у тебя)
    df2 = pd.DataFrame(columns=[
        'Наменование МО','Код МО','Количество попыток записаться через ЕПГУ','Черновики',
        'Успешные записи','Ошибки все','Орг. ошибки','Тех. Ошибки','Ошибки фед. уровня',
        'Общее количество талонов','Общее количество талонов (все интервалы)','Выложено на ЕПГУ',
    ])

    def _val(cell):
        # cell примерно: ["ТЕКСТ МЕТРИКИ", [число]]
        try:
            return cell[1][0]
        except Exception:
            return 0

    # чтобы не делать двойной цикл len(df)*len(lpu), делаем словарь (логика та же)
    lpu_map = {lpu["__name__"][i]: (lpu["__name__"][i], lpu["__code__"][i]) for i in range(len(lpu))}

    for i in range(len(df)):
        key = df[0][i]
        if key in lpu_map:
            name_clean, code = lpu_map[key]

            # индексы колонок 1..6 как в твоём коде
            chern = _val(df[1][i])
            osh_all = _val(df[2][i])
            org_osh = _val(df[3][i])
            teh_osh = _val(df[4][i])
            usp = _val(df[5][i])
            fed = _val(df[6][i])

            tries = chern + osh_all + org_osh + teh_osh + usp

            df2 = pd.concat([df2, pd.DataFrame([{
                'Наменование МО': name_clean,
                'Код МО': code,
                'Количество попыток записаться через ЕПГУ': tries,
                'Черновики': chern,
                'Успешные записи': usp,
                'Ошибки все': osh_all,
                'Орг. ошибки': org_osh,
                'Тех. Ошибки': teh_osh,
                'Ошибки фед. уровня': fed,
            }])], ignore_index=True)

    # 5) период прошлой недели (как у тебя)
    date = datetime.date.today()
    iso = date.isocalendar()
    d = f"{Week(date.year, (iso[1]-1)).monday()} - {Week(date.year, (iso[1]-1)).sunday()}"

    # 6) открываем шаблон и заполняем
    wb = load_workbook(template_xlsx_path)
    ws = wb[sheet_name]
    ws['C1'] = str(d)

    # тот же алгоритм сопоставления (у тебя было 0..90)
    # сделаем аккуратно: до 90 или до конца lpu
    limit = min(90, len(lpu))

    # быстрый доступ: имя -> строка df2
    df2_map = {df2['Наменование МО'][i].lower(): i for i in range(len(df2))}

    for i in range(limit):
        nm = lpu["__name__"][i].lower()
        if nm in df2_map:
            c = df2_map[nm]
            row = i + 4
            ws[f'D{row}'] = df2['Черновики'][c]
            ws[f'F{row}'] = df2['Успешные записи'][c]
            ws[f'I{row}'] = df2['Ошибки все'][c]
            ws[f'K{row}'] = df2['Орг. ошибки'][c]
            ws[f'M{row}'] = df2['Тех. Ошибки'][c]
            ws[f'O{row}'] = df2['Ошибки фед. уровня'][c]

            d = nz(ws[f"D{row}"].value)
            f = nz(ws[f"F{row}"].value)
            i_val = nz(ws[f"I{row}"].value)

            # 1) C = D + F + I
            ws[f"C{row}"] = d + f + i_val

            # 2) проценты (если в шаблоне уже есть — не трогаем; если пусто — проставим)
            ensure_formula(ws, f"E{row}", f"=D{row}/C{row}")
            ensure_formula(ws, f"G{row}", f"=F{row}/C{row}")
            ensure_formula(ws, f"H{row}", f"=IF(F{row}<>0,F{row}/(F{row}+K{row}+M{row}),0)")
            ensure_formula(ws, f"J{row}", f"=I{row}/C{row}")
            ensure_formula(ws, f"L{row}", f"=K{row}/C{row}")
            ensure_formula(ws, f"N{row}", f"=M{row}/C{row}")
            ensure_formula(ws, f"P{row}", f"=O{row}/C{row}")

            code_key = norm_id(lpu["__code__"][i])  # РМИС ID (если есть)
            vals = None

            if code_key:
                vals = ticket_by_code.get(code_key)

            if vals is None:
                # fallback по имени (на случай если кода нет/не совпал формат)
                vals = ticket_by_name.get(nm)

            if vals:
                portal_all, total, total_all = vals
                ws[f"Q{row}"] = nz(portal_all)  # Выложено на ЕПГУ (конкурентные)
                ws[f"R{row}"] = nz(total)  # Общее количество талонов
                ws[f"S{row}"] = nz(total_all)  # Общее количество талонов (все интервалы)
            else:
                # если не нашли — можно оставить как есть, либо поставить 0
                ws[f"Q{row}"] = nz(ws[f"Q{row}"].value)
                ws[f"R{row}"] = nz(ws[f"R{row}"].value)
                ws[f"S{row}"] = nz(ws[f"S{row}"].value)

            # формулы процентов по этим колонкам (обычно в шаблоне уже есть)
            ensure_formula(ws, f"T{row}", f"=Q{row}/R{row}")
            #ensure_formula(ws, f"W{row}", f"=Q{row}/V{row}")

    wb.save(output_xlsx_path)
    return output_xlsx_path

def norm_id(v):
    if v is None:
        return None
    # Excel часто отдаёт числа как float
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()

def load_ticket_maps(ticket_xlsx_path: str):
    """
    Возвращает 2 словаря:
      by_code:  RMIS_ID -> (portal_all, total, total_all)
      by_name:  cleaned_name -> (portal_all, total, total_all)
    """
    wb_t = load_workbook(ticket_xlsx_path, data_only=True)
    ws_t = wb_t["Лист 1"]

    by_code = {}
    by_name = {}

    for r in range(1, ws_t.max_row + 1):
        code = norm_id(ws_t[f"B{r}"].value)            # у тебя там обычно РМИС ID
        name_clean = _clean_name(ws_t[f"A{r}"].value)  # если в A есть название МО
        portal_all = ws_t[f"C{r}"].value               # PORTAL_ALL
        total = ws_t[f"D{r}"].value                    # TOTAL
        total_all = ws_t[f"E{r}"].value                # TOTAL_ALL

        if code:
            by_code[code] = (portal_all, total, total_all)
        if name_clean:
            by_name[name_clean] = (portal_all, total, total_all)

    return by_code, by_name



def get_last_week_dates():
    date = datetime.datetime.today()
    lol = date.isocalendar()
    return Week(date.year, (lol[1] - 1)).monday().strftime("%d.%m.%Y"), Week(date.year, (lol[1] - 1)).sunday().strftime(
        "%d.%m.%Y")


def process_lpu(lpu_df: pd.DataFrame, ws, wb, cancel_lpu: pd.DataFrame = None):
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

            import io
            xml_clean = resp_data.text.replace('<DataSet name="DS_TALONS" sysid="0">', "").replace('</DataSet>', "")
            df = pd.read_xml(io.StringIO(xml_clean))

            print(f"[DEBUG] DataFrame shape: {df.shape}")
            print(f"[DEBUG] Columns: {df.columns.tolist()}")

            for k in range(3, 94):
                if lpu_df["РМИС ID"][i] == ws[f'B{k}'].value:
                    row_index = i + 4
                    ws[f'C{row_index}'] = df['PORTAL_ALL'].sum()
                    ws[f'D{row_index}'] = df['TOTAL'].sum()
                    ws[f'E{row_index}'] = df['TOTAL_ALL'].sum()
                    ws[f'F{row_index}'] = datetime.datetime.today()
                    break

            wb.save(settings.excel_path_ticket)
            print(f"[INFO] Сохранено для LPU {lpu_df['РМИС ID'][i]}")

        except Exception as e:
            print(f"[ERROR] Ошибка для LPU {lpu_df['РМИС ID'][i]}: {e}")
            if cancel_lpu is None:
                cancel_lpu = pd.DataFrame()

            new_row = pd.DataFrame({
                "Наменование МО": [lpu_df["Наменование МО"][i]],
                "РМИС ID": [lpu_df["РМИС ID"][i]]
            })
            global cancelLPU
            cancelLPU = pd.concat([cancelLPU, new_row], ignore_index=True)


def main_process():
    global cancelLPU

    LPU = pd.read_excel(settings.lpu_path)
    wb_obj = load_workbook(settings.excel_path_ticket)
    ws = wb_obj["Лист 1"]

    start_time = time.time()


    if not cancelLPU.empty:
        print("[INFO] Обработка неудачных LPU...")
        process_lpu(cancelLPU.drop_duplicates(), ws, wb_obj, cancelLPU)


    print("[INFO] Основная обработка...")
    process_lpu(LPU, ws, wb_obj, cancelLPU)

    wb_obj.save(settings.excel_path_ticket)
    duration = time.time() - start_time
    print(f"[INFO] Обработка завершена за {duration:.2f} сек")
    return duration
