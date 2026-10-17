

import argparse
import json
import os
import sys
import time
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import requests
from dotenv import load_dotenv

try:
    import gspread
    from google.oauth2.service_account import Credentials
except ImportError:
    print("Не найдены gspread / google-auth. Установи: pip install -r requirements.txt")
    raise


# =========================================================================
# CONFIG
# =========================================================================

load_dotenv()

SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "")

WB_STATS_TOKEN = os.getenv("WB_TOKEN_1")
WB_ADVERT_TOKEN = os.getenv("WB_ADVERT_TOKEN", WB_STATS_TOKEN)
WB_ANALYTICS_TOKEN = os.getenv("WB_ANALYTICS_TOKEN", WB_STATS_TOKEN)
FULLSTATS_IDS_CHUNK_SIZE = 50
FULLSTATS_REQUEST_DELAY_SECONDS = 3  # между пачками одного запуска
MAIN_SHEET_NAME = "Все артикулы API"
SHEET_ORDERS = "orders"
SHEET_SALES = "sales"
SHEET_ANALYTICS = "analyticsHistory"
SHEET_ADS_STATS = "A_S"
SHEET_ADVERT_EXPENSES = "AdvertExpenses"

MSK = ZoneInfo("Europe/Moscow")

DEBUG_DUMP_RAW = True
DRY_RUN = False          # True = посчитать/показать данные, но НЕ писать их в Google Таблицу
REQUEST_TIMEOUT = 60
MAX_RETRIES = 5
RETRY_SLEEP_SECONDS = 20  # WB жёстко лимитирует некоторые методы (1 запрос/мин)

# Метод воронки (sales-funnel/products/history) принимает максимум 20 nmIds
# за один запрос и ограничен 3 запросами в минуту — держим паузу между пачками.
FUNNEL_NM_IDS_CHUNK_SIZE = 20
FUNNEL_REQUEST_DELAY_SECONDS = 21


# =========================================================================
# ВСПОМОГАТЕЛЬНОЕ: HTTP к WB API с ретраями на 429
# =========================================================================

def wb_request(method, url, token, **kwargs):
    headers = kwargs.pop("headers", {})
    headers["Authorization"] = token
    for attempt in range(1, MAX_RETRIES + 1):
        resp = requests.request(method, url, headers=headers, timeout=REQUEST_TIMEOUT, **kwargs)
        if resp.status_code == 429:
            print(f"  429 Too Many Requests на {url}, жду {RETRY_SLEEP_SECONDS}с "
                  f"(попытка {attempt}/{MAX_RETRIES})...")
            time.sleep(RETRY_SLEEP_SECONDS)
            continue
        if resp.status_code >= 400:
            raise RuntimeError(f"WB API {method} {url} -> {resp.status_code}: {resp.text[:500]}")
        if not resp.text:
            return None
        return resp.json()
    raise RuntimeError(f"WB API {method} {url}: не удалось получить ответ после {MAX_RETRIES} попыток (429)")


def wb_get(url, token, params=None):
    return wb_request("GET", url, token, params=params)


def wb_post(url, token, body=None):
    return wb_request("POST", url, token, json=body)


def dump_raw_once(label, data, dumped_labels=set()):
    if DEBUG_DUMP_RAW and label not in dumped_labels:
        dumped_labels.add(label)
        print(f"\n----- RAW DEBUG DUMP: {label} -----")
        try:
            print(json.dumps(data, ensure_ascii=False, indent=2)[:4000])
        except Exception:
            print(repr(data)[:4000])
        print(f"----- /RAW DEBUG DUMP: {label} -----\n")


# =========================================================================
# GOOGLE SHEETS
# =========================================================================

def open_spreadsheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID)


def get_nm_ids(sh):
    """
    Собирает список артикулов WB (nmID) с листа "Все артикулы".
    Лист устроен блоками, точные номера строк внутри блока не хардкодим —
    ищем ячейки с меткой "артикул wb" и берём ближайшее числовое значение
    (6-10 цифр) рядом с ней (в той же ячейке после переноса строки, справа
    или снизу в пределах нескольких ячеек).
    """
    ws = sh.worksheet(MAIN_SHEET_NAME)
    grid = ws.get_all_values()

    nm_ids = []
    seen = set()

    for r, row in enumerate(grid):
        for c, cell in enumerate(row):
            if not cell:
                continue
            if "артикул wb" not in cell.strip().lower():
                continue

            candidates = []
            if "\n" in cell:
                candidates.append(cell.split("\n")[-1])
            for dc in range(0, 4):
                if c + dc < len(row):
                    candidates.append(row[c + dc])
            for dr in range(1, 3):
                if r + dr < len(grid) and c < len(grid[r + dr]):
                    candidates.append(grid[r + dr][c])

            for val in candidates:
                val = (val or "").strip()
                if val.isdigit() and 6 <= len(val) <= 10 and val not in seen:
                    seen.add(val)
                    nm_ids.append(val)
                    break

    print(f"Найдено артикулов WB (nmID) на листе '{MAIN_SHEET_NAME}': {len(nm_ids)}")
    if len(nm_ids) != 63:
        print("  ВНИМАНИЕ: ожидалось 63 блока/артикула (по твоему описанию структуры). "
              "Если число не совпадает — проверь, что метка 'Артикул WB' находится "
              "достаточно близко к самому значению nmID, либо пришли структуру блока "
              "точнее (номера строк/колонок), чтобы поправить get_nm_ids().")
    return nm_ids


def append_records(ws, records, sheet_label):
    """
    records: список плоских dict (ключи = имена полей WB API).
    Пишем по заголовку (row 1) листа: для каждой колонки заголовка ищем
    значение в record по точному имени, иначе по имени без учёта регистра.
    """
    if not records:
        print(f"[{sheet_label}] нет данных для записи — пропуск.")
        return

    header = ws.row_values(1)
    if not header:
        print(f"[{sheet_label}] ПРЕДУПРЕЖДЕНИЕ: в листе нет строки заголовков (row 1). "
              f"Лист пропущен — добавь заголовки вручную (по именам полей WB API) и перезапусти.")
        return

    header_lower = {h.strip().lower(): h for h in header if h}

    all_keys = set()
    for rec in records:
        all_keys.update(rec.keys())
    unmatched = sorted(
        k for k in all_keys
        if k not in header and k.lower() not in header_lower
    )
    if unmatched:
        print(f"[{sheet_label}] поля из ответа WB API, которых нет в заголовке листа "
              f"(будут проигнорированы): {unmatched}")

    rows_to_write = []
    for rec in records:
        row = []
        for col in header:
            if col in rec:
                row.append(rec[col])
            else:
                row.append(rec.get(header_lower.get(col.strip().lower(), ""), ""))
        rows_to_write.append(row)

    ws.append_rows(rows_to_write, value_input_option="USER_ENTERED")
    print(f"[{sheet_label}] добавлено строк: {len(rows_to_write)}")


# =========================================================================
# WB API: Statistics (orders / sales) — схема полей задокументирована и стабильна
# =========================================================================

def fetch_orders(target_date):
    url = "https://statistics-api.wildberries.ru/api/v1/supplier/orders"
    data = wb_get(url, WB_STATS_TOKEN, {"dateFrom": target_date.isoformat(), "flag": 1}) or []
    dump_raw_once("orders", data[:2] if isinstance(data, list) else data)
    return data if isinstance(data, list) else []


def fetch_sales(target_date):
    url = "https://statistics-api.wildberries.ru/api/v1/supplier/sales"
    data = wb_get(url, WB_STATS_TOKEN, {"dateFrom": target_date.isoformat(), "flag": 1}) or []
    dump_raw_once("sales", data[:2] if isinstance(data, list) else data)
    return data if isinstance(data, list) else []


def fetch_funnel(nm_ids, target_date):
    url = "https://seller-analytics-api.wildberries.ru/api/analytics/v3/sales-funnel/products/history"
    date_str = target_date.isoformat()
    nm_ids_int = [int(x) for x in nm_ids]

    chunks = [
        nm_ids_int[i:i + FUNNEL_NM_IDS_CHUNK_SIZE]
        for i in range(0, len(nm_ids_int), FUNNEL_NM_IDS_CHUNK_SIZE)
    ]

    all_flat = []
    for idx, chunk in enumerate(chunks, start=1):
        body = {
            "nmIds": chunk,
            "selectedPeriod": {"start": date_str, "end": date_str},
            "skipDeletedNm": True,
            "aggregationLevel": "day",
        }
        print(f"  [funnel] пачка {idx}/{len(chunks)} ({len(chunk)} nmID)...")
        data = wb_post(url, WB_ANALYTICS_TOKEN, body)
        dump_raw_once("analyticsHistory (funnel)", data)
        all_flat.extend(flatten_funnel(data))
        if idx < len(chunks):
            time.sleep(FUNNEL_REQUEST_DELAY_SECONDS)

    return all_flat


def flatten_funnel(data):
    """
    Форма ответа v3 (по документации dev.wildberries.ru/docs/openapi/analytics):
        {
          "data": [
            {
              "product": {"nmId": 123, "title": "...", "vendorCode": "...",
                           "brandName": "...", "subjectId": 105, "subjectName": "..."},
              "history": [
                {"date": "2026-09-14", "openCount": ..., "cartCount": ...,
                 "orderCount": ..., "orderSum": ..., "buyoutCount": ...,
                 "buyoutSum": ..., "buyoutPercent": ..., "addToCartConversion": ...,
                 "cartToOrderConversion": ..., "addToWishlistCount": ...}
              ],
              "currency": "RUB"
            }
          ]
        }
    Разворачиваем до одной строки на nmId+дата: поля из "product" (кроме
    вложенного объекта) и "history" объединяются в плоский dict.
    ЕСЛИ РЕАЛЬНЫЙ ОТВЕТ ИМЕЕТ ДРУГУЮ ФОРМУ — смотри DEBUG DUMP выше и
    поправь эту функцию (это единственное место, которое их читает).
    """
    if not data:
        return []
    cards = data.get("data") if isinstance(data, dict) else data
    if cards is None:
        return []

    flat = []
    for card in cards:
        product = card.get("product") or {}
        nm_id = product.get("nmId") or product.get("nmID") or card.get("nmID") or card.get("nmId")
        currency = card.get("currency")
        history = card.get("history") or card.get("stats") or []
        if isinstance(history, dict):
            history = [history]
        for point in history:
            row = {"nmID": nm_id}
            row.update({k: v for k, v in product.items() if k not in ("nmId", "nmID")})
            if currency is not None:
                row["currency"] = currency
            row.update(point)
            flat.append(row)
    return flat


# =========================================================================
# WB API: Advert — статистика по кампаниям (A_S) + списания (AdvertExpenses)
# ВНИМАНИЕ: fullstats часто меняет версию (v1/v2/v3) и формат тела запроса —
# сверяй с DEBUG DUMP на первом запуске.
# =========================================================================

def fetch_campaign_ids():
    """Список ID кампаний продавца (нужны для fullstats)."""
    url = "https://advert-api.wildberries.ru/adv/v1/promotion/count"
    data = wb_get(url, WB_ADVERT_TOKEN) or {}
    dump_raw_once("advert campaigns list", data)
    ids = []
    for group in data.get("adverts", []):
        for adv in group.get("advert_list", group.get("advertList", [])):
            adv_id = adv.get("advertId") or adv.get("advert_id")
            if adv_id:
                ids.append(adv_id)
    return ids


def fetch_ads_fullstats(campaign_ids, target_date):
    """
    GET /adv/v3/fullstats принимает максимум 50 ID кампаний за один запрос —
    бьём список на пачки и делаем паузу между ними.
    """
    if not campaign_ids:
        print("[A_S] нет активных/доступных кампаний — пропуск.")
        return []

    url = "https://advert-api.wildberries.ru/adv/v3/fullstats"
    date_str = target_date.isoformat()

    chunks = [
        campaign_ids[i:i + FULLSTATS_IDS_CHUNK_SIZE]
        for i in range(0, len(campaign_ids), FULLSTATS_IDS_CHUNK_SIZE)
    ]

    all_flat = []
    for idx, chunk in enumerate(chunks, start=1):
        params = {
            "ids": ",".join(str(cid) for cid in chunk),
            "beginDate": date_str,
            "endDate": date_str,
        }
        print(f"  [fullstats] пачка {idx}/{len(chunks)} ({len(chunk)} кампаний)...")
        data = wb_get(url, WB_ADVERT_TOKEN, params)
        dump_raw_once("A_S (advert fullstats v3)", data)
        all_flat.extend(flatten_ads_fullstats(data))
        if idx < len(chunks):
            time.sleep(FULLSTATS_REQUEST_DELAY_SECONDS)

    return all_flat

def flatten_ads_fullstats(data):
    """
    Ответ v3 — список объектов по кампаниям:
        [{"advertId": 123, "canceled": 0, ...,
          "days": [{"date": "2026-09-14T00:00:00Z", "canceled": 0, ...,
                     "apps": [{"appType": 32, ...,
                                "nms": [{"nmId": ..., "name": ..., "views": ...,
                                          "clicks": ..., "orders": ..., "sum": ...,
                                          "canceled": ...}]}]}]}]
    Разворачиваем до уровня nmId+дата (появился новый уровень apps->nms
    по сравнению со старой v2-схемой). Если у карточки нет разбивки по nms —
    оставляем строку на уровне appType/дня.
    """
    if not data:
        return []
    flat = []
    for camp in data:
        advert_id = camp.get("advertId") or camp.get("advert_id")
        for day in camp.get("days", []):
            date_str = day.get("date")
            apps = day.get("apps") or []
            if apps:
                for app in apps:
                    nms = app.get("nms") or []
                    if nms:
                        for nm in nms:
                            row = {"advertId": advert_id, "date": date_str,
                                   "appType": app.get("appType")}
                            row.update(nm)
                            flat.append(row)
                    else:
                        row = {"advertId": advert_id, "date": date_str}
                        row.update({k: v for k, v in app.items() if k != "nms"})
                        flat.append(row)
            else:
                row = {"advertId": advert_id, "date": date_str}
                row.update({k: v for k, v in day.items() if k not in ("date", "apps")})
                flat.append(row)
    return flat

def fetch_advert_expenses(target_date):
    url = "https://advert-api.wildberries.ru/adv/v1/upd"
    data = wb_get(url, WB_ADVERT_TOKEN, {
        "from": target_date.isoformat(),
        "to": target_date.isoformat(),
    }) or []
    dump_raw_once("AdvertExpenses (upd)", data[:2] if isinstance(data, list) else data)
    return data if isinstance(data, list) else []


# =========================================================================
# MAIN
# =========================================================================

def parse_args():
    p = argparse.ArgumentParser(description="Дозаполнить сырые листы Google Sheets данными WB за день.")
    p.add_argument("--date", help="Дата в формате YYYY-MM-DD (по умолчанию — вчера по МСК).")
    return p.parse_args()


def main():
    args = parse_args()

    if args.date:
        target_date = datetime.strptime(args.date, "%Y-%m-%d").date()
    else:
        target_date = (datetime.now(MSK) - timedelta(days=3)).date()

    print(f"Целевая дата: {target_date.isoformat()}")
    if DRY_RUN:
        print("Режим DRY_RUN включён (переменная DRY_RUN=True в CONFIG) — в таблицу ничего писаться не будет.")

    if not WB_STATS_TOKEN:
        print("ОШИБКА: WB_TOKEN_1 не найден в .env")
        sys.exit(1)
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        print(f"ОШИБКА: не найден файл сервис-аккаунта: {SERVICE_ACCOUNT_FILE}")
        sys.exit(1)

    print("Подключаюсь к Google Sheets...")
    sh = open_spreadsheet()

    nm_ids = get_nm_ids(sh)

    print("\n== Statistics API: orders ==")
    orders = fetch_orders(target_date)
    print(f"  получено записей: {len(orders)}")

    print("\n== Statistics API: sales ==")
    sales = fetch_sales(target_date)
    print(f"  получено записей: {len(sales)}")

    print("\n== Seller Analytics API: воронка ==")
    funnel = fetch_funnel(nm_ids, target_date) if nm_ids else []
    print(f"  получено записей: {len(funnel)}")

    print("\n== Advert API: кампании + fullstats ==")
    campaign_ids = fetch_campaign_ids()
    print(f"  найдено кампаний: {len(campaign_ids)}")
    ads_stats = fetch_ads_fullstats(campaign_ids, target_date)
    print(f"  получено записей: {len(ads_stats)}")

    print("\n== Advert API: списания (upd) ==")
    advert_expenses = fetch_advert_expenses(target_date)
    print(f"  получено записей: {len(advert_expenses)}")

    if DRY_RUN:
        print("\nDRY_RUN=True: в таблицу ничего не пишу.")
        return
    from write_all_articles import write_to_all_articles
    write_to_all_articles(sh, target_date, orders, sales, funnel, ads_stats, dry_run=DRY_RUN)
    print("\n== Запись в Google Sheets ==")
    append_records(sh.worksheet(SHEET_ORDERS), orders, SHEET_ORDERS)
    append_records(sh.worksheet(SHEET_SALES), sales, SHEET_SALES)
    append_records(sh.worksheet(SHEET_ANALYTICS), funnel, SHEET_ANALYTICS)
    append_records(sh.worksheet(SHEET_ADS_STATS), ads_stats, SHEET_ADS_STATS)
    append_records(sh.worksheet(SHEET_ADVERT_EXPENSES), advert_expenses, SHEET_ADVERT_EXPENSES)
    from wb_new_fetches import fetch_stocks, fetch_finance_report, fetch_paid_storage
    from write_all_articles import write_to_all_articles

    print("\n== Аналитика: остатки ==")
    stocks = fetch_stocks()
    print(f"  получено записей: {len(stocks)}")

    print("\n== Finance: логистика/комиссия/штрафы ==")
    finance_rows = fetch_finance_report(target_date, target_date)
    print(f"  получено записей: {len(finance_rows)}")

    print("\n== Аналитика: платное хранение ==")
    storage_rows = fetch_paid_storage(target_date, target_date)
    print(f"  получено записей: {len(storage_rows)}")

    write_to_all_articles(
        sh, target_date, orders, sales, funnel, ads_stats,
        stocks=stocks, finance_rows=finance_rows, storage_rows=storage_rows,
        dry_run=DRY_RUN,
    )

    print("\nГотово.")

if __name__ == "__main__":
    main()