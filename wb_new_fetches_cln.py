

import os
import time
from collections import defaultdict

import requests

# Переиспользуем то, что уже есть в wb_daily_update.py — импортируй оттуда
# или скопируй wb_get/wb_post/dump_raw_once/REQUEST_TIMEOUT/MAX_RETRIES при интеграции.
from wb_daily_update import (
    wb_get, wb_post, dump_raw_once, WB_STATS_TOKEN, WB_ANALYTICS_TOKEN, DEBUG_DUMP_RAW,
)

WB_FINANCE_TOKEN = os.getenv("WB_FINANCE_TOKEN", WB_STATS_TOKEN)

PAID_STORAGE_POLL_DELAY_SECONDS = 5
PAID_STORAGE_MAX_POLLS = 24  # 24 * 5с = до 2 минут ожидания готовности отчёта


# =========================================================================
# 1) ОСТАТКИ
# =========================================================================

def fetch_stocks():
    """
    Возвращает список сырых строк остатков (что бы ни вернул WB —
    aggregate_stocks() сама разберётся с полями, но их точные имена
    нужно свериться по DEBUG_DUMP_RAW).
    """
    url = "https://seller-analytics-api.wildberries.ru/api/analytics/v1/stocks-report/wb-warehouses"
    data = wb_post(url, WB_ANALYTICS_TOKEN, {})
    dump_raw_once("stocks (wb-warehouses, сырой ответ)", data)

    # Метод может оказаться синхронным (сразу данные) ИЛИ асинхронным
    # (вернёт taskId, который нужно опрашивать) — обрабатываем оба варианта.
    task_id = None
    if isinstance(data, dict):
        task_id = (data.get("data") or {}).get("taskId") or data.get("taskId")

    if task_id:
        print(f"  [stocks] метод асинхронный, taskId={task_id}, опрашиваю статус...")
        for attempt in range(PAID_STORAGE_MAX_POLLS):
            status_resp = wb_get(f"{url}/tasks/{task_id}/status", WB_ANALYTICS_TOKEN)
            status = (status_resp or {}).get("data", {}).get("status") or (status_resp or {}).get("status")
            if status in ("done", "DONE", "success", "SUCCESS"):
                data = wb_get(f"{url}/tasks/{task_id}/download", WB_ANALYTICS_TOKEN)
                dump_raw_once("stocks (wb-warehouses, download)", data)
                break
            time.sleep(PAID_STORAGE_POLL_DELAY_SECONDS)
        else:
            print("  [stocks] отчёт не готов вовремя, пропускаю остатки за этот запуск.")
            return []

    rows = data.get("data") if isinstance(data, dict) else data
    if isinstance(rows, dict):

        list_candidates = [v for v in rows.values() if isinstance(v, list)]
        if list_candidates:
            rows = list_candidates[0]
        else:
            print(f"  [stocks] НЕ РАСПОЗНАЛ структуру ответа (ожидал список, "
                  f"получил словарь с ключами {list(rows.keys())}). "
                  f"Смотри RAW DEBUG DUMP выше и пришли его — поправим разбор.")
            return []
    return rows or []


def aggregate_stocks(stocks_raw):

    agg = defaultdict(lambda: {"on_warehouse": 0, "to_client": 0, "from_client": 0, "warehouses": set()})
    for r in stocks_raw:
        nm = str(r.get("nmID") or r.get("nmId") or "").strip()
        if not nm:
            continue
        a = agg[nm]
        wh_name = (r.get("warehouseName") or "").strip().lower()
        qty = r.get("quantity", 0) or 0
        if "к клиент" in wh_name or wh_name == "to the client":
            a["to_client"] += qty
        elif "от клиент" in wh_name or wh_name == "from the client":
            a["from_client"] += qty
        else:
            a["on_warehouse"] += qty
            if qty:
                a["warehouses"].add(wh_name)
        # если метод всё ещё отдаёт отдельные поля — берём их приоритетно
        if r.get("inWayToClient") is not None:
            a["to_client"] += r.get("inWayToClient", 0) or 0
        if r.get("inWayFromClient") is not None:
            a["from_client"] += r.get("inWayFromClient", 0) or 0

    result = {}
    for nm, a in agg.items():
        result[nm] = {
            "stock_on_warehouse": a["on_warehouse"],
            "stock_to_client": a["to_client"],
            "stock_from_client": a["from_client"],
            "stock_warehouses_count": len(a["warehouses"]),
        }
    return result


# =========================================================================
# 2) ЛОГИСТИКА / КОМИССИЯ / ШТРАФЫ (новый Finance-метод)
# =========================================================================

def fetch_finance_report(date_from, date_to):
    """
    Реальная схема метода (по официальной документации):
    - ответ — ПРОСТОЙ МАССИВ [...], без обёртки "data"
    - обязательна пагинация через rrdId: повторяем запрос с rrdId из
      последней строки предыдущего ответа, пока не придёт 204 (пустой ответ)
    - лимит: 1 запрос в минуту — при пагинации обязательна пауза
    - нужно явно передавать period="daily", иначе по умолчанию "weekly"
    - поля camelCase: nmId, docTypeName, deliveryService (логистика),
      rebillLogisticCost (перевыставленная логистика), ppvzSalesCommission
      (комиссия), penalty + additionalPayment (штрафы), paidStorage (хранение)
    """
    url = "https://finance-api.wildberries.ru/api/finance/v1/sales-reports/detailed"
    fields = [
        "rrdId", "nmId", "docTypeName",
        "deliveryService", "rebillLogisticCost",
        "ppvzSalesCommission", "penalty", "additionalPayment", "paidStorage",
    ]
    all_rows = []
    rrd_id = 0
    page = 0
    while True:
        body = {
            "dateFrom": date_from.isoformat(),
            "dateTo": date_to.isoformat(),
            "limit": 100000,
            "rrdId": rrd_id,
            "period": "daily",
            "fields": fields,
        }
        data = wb_post(url, WB_FINANCE_TOKEN, body)
        dump_raw_once("finance report (sales-reports/detailed)",
                       data[:3] if isinstance(data, list) else data)
        rows = data if isinstance(data, list) else (data or {}).get("data", [])
        if not rows:
            break  # 204 / пустой ответ — отчёт закончился
        all_rows.extend(rows)
        page += 1
        if len(rows) < body["limit"]:
            break  # последняя (неполная) страница
        rrd_id = rows[-1].get("rrdId", 0)
        print(f"  [finance] страница {page} получена ({len(rows)} строк), жду 61с "
              f"(лимит метода — 1 запрос/мин)...")
        time.sleep(61)

    doc_type_counts = defaultdict(int)
    doc_type_commission_sum = defaultdict(float)
    for r in all_rows:
        dt = r.get("docTypeName") or "(пусто)"
        doc_type_counts[dt] += 1
        try:
            doc_type_commission_sum[dt] += float(r.get("ppvzSalesCommission") or 0)
        except (TypeError, ValueError):
            pass
    print(f"  [finance] всего строк: {len(all_rows)}. Разбивка по docTypeName "
          f"(кол-во строк / сумма ppvzSalesCommission):")
    for dt, cnt in sorted(doc_type_counts.items(), key=lambda x: -x[1]):
        print(f"    '{dt}': {cnt} строк, комиссия={round(doc_type_commission_sum[dt], 2)}")

    return all_rows


def aggregate_finance(finance_rows):
    """
    nmID(str) -> Логистика/Штрафы/Хранение за период (в рублях).
    """
    def to_float(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0.0

    agg = defaultdict(lambda: {"logistics": 0.0, "penalty": 0.0, "storage": 0.0})
    for r in finance_rows:
        nm = str(r.get("nmId") or r.get("nmID") or "").strip()
        if not nm:
            continue
        a = agg[nm]
        a["logistics"] += to_float(r.get("deliveryService")) + to_float(r.get("rebillLogisticCost"))
        a["penalty"] += to_float(r.get("penalty")) + to_float(r.get("additionalPayment"))
        a["storage"] += to_float(r.get("paidStorage"))

    return {
        nm: {
            # Пишем ОТРИЦАТЕЛЬНЫМИ числами — так они хранятся на листе клиента.
            "logistics_rub": round(-a["logistics"], 2),
            "penalty_rub": round(-a["penalty"], 2),
        }
        for nm, a in agg.items()
    }


# =========================================================================
# 3) ХРАНЕНИЕ (paid_storage, паттерн "задача")
# =========================================================================

def fetch_paid_storage(date_from, date_to):
    base = "https://seller-analytics-api.wildberries.ru/api/v1/paid_storage"
    create_resp = wb_get(base, WB_ANALYTICS_TOKEN, {
        "dateFrom": date_from.isoformat(),
        "dateTo": date_to.isoformat(),
    })
    dump_raw_once("paid_storage (создание задачи)", create_resp)
    task_id = (create_resp or {}).get("data", {}).get("taskId")
    if not task_id:
        print("  [paid_storage] не получил taskId — пропускаю хранение.")
        return []

    for attempt in range(PAID_STORAGE_MAX_POLLS):
        status_resp = wb_get(f"{base}/tasks/{task_id}/status", WB_ANALYTICS_TOKEN)
        status = (status_resp or {}).get("data", {}).get("status")
        if status == "done":
            break
        time.sleep(PAID_STORAGE_POLL_DELAY_SECONDS)
    else:
        print("  [paid_storage] отчёт не готов вовремя, пропускаю хранение за этот запуск.")
        return []

    data = wb_get(f"{base}/tasks/{task_id}/download", WB_ANALYTICS_TOKEN)
    dump_raw_once("paid_storage (download)", data[:3] if isinstance(data, list) else data)
    return data if isinstance(data, list) else []


def aggregate_paid_storage(storage_rows):
    """nmID(str) -> Хранение, руб (сумма warehousePrice за период, ОТРИЦАТЕЛЬНЫМ числом)."""
    agg = defaultdict(float)
    for r in storage_rows:
        nm = str(r.get("nmId") or r.get("nmID") or "").strip()
        if not nm:
            continue
        agg[nm] += float(r.get("warehousePrice", 0) or 0)
    return {nm: {"storage_rub": round(-v, 2)} for nm, v in agg.items()}