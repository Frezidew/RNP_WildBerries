

import os
from collections import defaultdict


TARGET_SHEET_NAME = os.getenv("TARGET_SHEET_NAME", "Все артикулы API")

ARTICLE_LABEL_COL_IDX = 6


METRIC_OFFSETS = {
    "orders_qty": 2,             # Заказы, шт
    "orders_rub_before_spp": 3,  # Заказы, руб до СПП
    "orders_rub_after_spp": 4,   # Заказы, руб после СПП
    "sales_qty": 5,              # Продажи, шт
    "sales_rub": 6,              # Продажи, руб
    "cancels_qty": 7,            # Отмены, шт
    "returns_qty": 8,            # Возвраты, шт (пишем ОТРИЦАТЕЛЬНЫМ числом)
    "returns_rub": 9,            # Возвраты, руб (пишем ОТРИЦАТЕЛЬНЫМ числом)
    "price": 10,                 # Цена, руб
    "discount_pct": 11,          # Скидка, %
    "price_with_discount": 12,   # Цена со скидкой, руб
    "avg_check_before_spp": 13,  # Средний чек до СПП, руб
    "avg_check_after_spp": 14,   # Средний чек после СПП, руб
    "avg_spp_pct": 15,           # Средний СПП, %
    "ads_views": 25,             # Просмотры
    "ads_clicks": 26,            # Клики
    "ads_ctr": 27,               # CTR
    "ads_cost": 28,              # Затраты (реклама), руб (пишем ОТРИЦАТЕЛЬНЫМ числом)
    "ads_cpc": 29,               # CPC
    "ads_carts": 30,             # Корзины (реклама), шт
    "ads_orders_qty": 31,        # Заказы (реклама), шт
    "ads_orders_rub": 32,        # Заказы (реклама), руб
    "ads_drr_pct": 33,           # ДРР к рекламным заказам, %
    "ads_drr_all_orders_pct": 34,  # ДРР ко всем заказам (до СПП), %
    "ads_drr_all_sales_pct": 35,   # ДРР ко всем продажам, %
    "funnel_visits": 38,         # Переходы
    "funnel_carts": 39,          # Корзины (воронка)
    "funnel_cr_cart": 40,        # CR в корзину, %
    "funnel_cr_order": 41,       # CR в заказ, %
    "funnel_buyout_pct": 42,     # Выкуп, %

    # Остатки (offset 22 "Оборачиваемость" сознательно не пишем — это
    # производная метрика, скорее всего уже формула на листе)
    "stock_on_warehouse": 18,    # На складах, шт
    "stock_to_client": 19,       # К клиенту, шт
    "stock_from_client": 20,     # От клиента, шт
    "stock_warehouses_count": 21,  # Складов, шт

    # Затраты. Все пишем ОТРИЦАТЕЛЬНЫМИ числами — так они хранятся на листе клиента.
    "logistics_rub": 45,         # Логистика, руб
    "storage_rub": 46,           # Хранение, руб
    "cost_total_rub": 47,        # Себестоимость, руб (за период = цена/шт х Продажи,шт), из листа "Шаблон для загрузки"
    "commission_rub": 48,        # Комиссия, руб
    "penalty_rub": 49,           # Штрафы, руб

    # Итого/Маржа
    "margin_rub": 52,            # Маржа, руб
    "margin_pct": 53,            # Маржа, %
}

COST_REFERENCE_SHEET_NAME = "Шаблон для загрузки"
COST_REFERENCE_SKU_COL_IDX = 1        # столбец B — SKU (nmID)
COST_REFERENCE_COST_COL_IDX = 3       # столбец D — Себестоимость, шт
COST_REFERENCE_COMMISSION_COL_IDX = 4  # столбец E — Комиссия, %


def _parse_ru_number(raw):
    """'1 159' -> 1159.0, '48,00%' -> 48.0. Возвращает None, если не число."""
    if not raw:
        return None
    cleaned = raw.strip().replace("\xa0", "").replace(" ", "").replace("%", "").replace(",", ".")
    if not cleaned:
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def fetch_cost_reference(sh, sheet_name=COST_REFERENCE_SHEET_NAME):
    """
    nmID(str) -> {"cost_per_unit": себестоимость за штуку в руб (столбец D),
                  "commission_pct": комиссия в % (столбец E)}
    из листа "Шаблон для загрузки". Столбец B — SKU/nmID.
    """
    ws = sh.worksheet(sheet_name)
    grid = ws.get_all_values()
    result = {}
    for row in grid[1:]:  # первая строка — заголовки
        if len(row) <= COST_REFERENCE_COMMISSION_COL_IDX:
            continue
        nm = row[COST_REFERENCE_SKU_COL_IDX].strip()
        if not nm.isdigit():
            continue
        cost = _parse_ru_number(row[COST_REFERENCE_COST_COL_IDX])
        commission_pct = _parse_ru_number(row[COST_REFERENCE_COMMISSION_COL_IDX])
        result[nm] = {"cost_per_unit": cost, "commission_pct": commission_pct}
    return result


def col_letter(col_idx_0based):
    letters = ""
    n = col_idx_0based + 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def find_product_blocks(grid):
    """nmID(str) -> label_row (0-based индекс строки с меткой 'Артикул WB')."""
    blocks = {}
    for r, row in enumerate(grid):
        if len(row) <= ARTICLE_LABEL_COL_IDX:
            continue
        if row[ARTICLE_LABEL_COL_IDX].strip().lower() != "артикул wb":
            continue
        if r + 1 >= len(grid) or len(grid[r + 1]) <= ARTICLE_LABEL_COL_IDX:
            continue
        val = grid[r + 1][ARTICLE_LABEL_COL_IDX].strip()
        if val.isdigit() and 6 <= len(val) <= 10:
            blocks[val] = r
    return blocks


def find_date_column(grid, target_date, header_search_rows=range(0, 15)):
    """Ищет столбец с датой target_date (формат дд.мм) в первых строках-шапках."""
    date_str = target_date.strftime("%d.%m")
    for r in header_search_rows:
        if r >= len(grid):
            continue
        for c, cell in enumerate(grid[r]):
            if cell.strip() == date_str:
                return c
    return None


def aggregate_orders(orders_raw):
    """
    nmID(str) -> заказы/отмены/цена/скидка/спп/чек за день (Statistics API orders).

    "до СПП" = priceWithDisc (цена со скидкой продавца, без скидки WB/кошелька).
    "после СПП" = finishedPrice (то, что реально заплатил покупатель).
    """
    agg = defaultdict(lambda: {
        "qty": 0, "cancels": 0,
        "price_sum": 0.0, "discount_sum": 0.0,
        "pwd_sum": 0.0, "finished_sum": 0.0, "spp_sum": 0.0,
        "n": 0,
    })
    for o in orders_raw:
        nm = str(o.get("nmId") or o.get("nmID") or "").strip()
        if not nm:
            continue
        a = agg[nm]
        # "Заказы, шт" — ВСЕ заказы (включая отменённые), т.к. "Заказы, руб"
        # ниже тоже считается по всем строкам без исключения отмен.
        # "Отмены" — отдельная информационная метрика, не вычитается из Заказов.
        a["qty"] += 1
        if o.get("isCancel"):
            a["cancels"] += 1
        price = o.get("totalPrice")
        if price is not None:
            a["price_sum"] += float(price)
            a["n"] += 1
            discount = o.get("discountPercent")
            pwd = o.get("priceWithDisc")
            finished = o.get("finishedPrice")
            spp = o.get("spp")
            if discount is not None:
                a["discount_sum"] += float(discount)
            if pwd is not None:
                a["pwd_sum"] += float(pwd)
            if finished is not None:
                a["finished_sum"] += float(finished)
            if spp is not None:
                a["spp_sum"] += float(spp)

    result = {}
    for nm, a in agg.items():
        n = a["n"] or 1
        result[nm] = {
            "orders_qty": a["qty"],
            "cancels_qty": a["cancels"],
            "price": round(a["price_sum"] / n, 2) if a["n"] else "",
            "discount_pct": round(a["discount_sum"] / n, 2) if a["n"] else "",
            "price_with_discount": round(a["pwd_sum"] / n, 2) if a["n"] else "",
            "avg_spp_pct": round(a["spp_sum"] / n, 2) if a["n"] else "",
            "orders_rub_before_spp": round(a["pwd_sum"], 2),
            "orders_rub_after_spp": round(a["finished_sum"], 2),
            # "Средний чек" — ДВЕ строки (до/после СПП), как и "Заказы, руб".
            "avg_check_before_spp": round(a["pwd_sum"] / n, 2) if a["n"] else "",
            "avg_check_after_spp": round(a["finished_sum"] / n, 2) if a["n"] else "",
        }
    return result


def aggregate_sales(sales_raw):

    agg = defaultdict(lambda: {"sales_qty": 0, "sales_rub": 0.0, "returns_qty": 0, "returns_rub": 0.0})
    for s in sales_raw:
        nm = str(s.get("nmId") or s.get("nmID") or "").strip()
        if not nm:
            continue
        a = agg[nm]

        sale_id = str(s.get("saleID") or "")
        if sale_id.upper().startswith("R"):
            a["returns_qty"] += 1
            finished = s.get("finishedPrice")
            if finished is not None:
                a["returns_rub"] += float(finished)
        else:
            a["sales_qty"] += 1
            pwd = s.get("priceWithDisc")
            if pwd is not None:
                a["sales_rub"] += float(pwd)

    return {
        nm: {
            "sales_qty": a["sales_qty"],
            "sales_rub": round(a["sales_rub"], 2),
            "returns_qty": -a["returns_qty"],

            "returns_rub": round(a["returns_rub"], 2),
        }
        for nm, a in agg.items()
    }


def aggregate_ads(ads_rows):
    """nmID(str) -> просмотры/клики/CTR/затраты/CPC/корзины/заказы шт+руб (A_S fullstats v3, уже сплющенные)."""
    agg = defaultdict(lambda: {"views": 0, "clicks": 0, "cost": 0.0, "carts": 0, "orders": 0, "orders_rub": 0.0})
    for r in ads_rows:
        nm = str(r.get("nmId") or "").strip()
        if not nm:
            continue
        a = agg[nm]
        a["views"] += r.get("views", 0) or 0
        a["clicks"] += r.get("clicks", 0) or 0
        a["cost"] += r.get("sum", 0) or 0
        a["carts"] += r.get("atbs", 0) or 0
        a["orders"] += r.get("orders", 0) or 0
        a["orders_rub"] += r.get("sum_price", 0) or 0

    result = {}
    for nm, a in agg.items():
        ctr = round(a["clicks"] / a["views"] * 100, 2) if a["views"] else 0
        cpc = round(a["cost"] / a["clicks"], 2) if a["clicks"] else 0
        result[nm] = {
            "ads_views": a["views"],
            "ads_clicks": a["clicks"],
            "ads_ctr": ctr,
            "ads_cost": round(-a["cost"], 2),
            "ads_cpc": cpc,
            "ads_carts": a["carts"],
            "ads_orders_qty": a["orders"],
            "ads_orders_rub": round(a["orders_rub"], 2),
        }
    return result


def aggregate_funnel(funnel_rows):
    """nmID(str) -> переходы/корзины/CR/выкуп (funnel v3, уже сплющенные, одна строка = один nmID+дата)."""
    result = {}
    for r in funnel_rows:
        nm = str(r.get("nmID") or r.get("nmId") or "").strip()
        if not nm:
            continue
        result[nm] = {
            "funnel_visits": r.get("openCount", 0) or 0,
            "funnel_carts": r.get("cartCount", 0) or 0,
            "funnel_cr_cart": r.get("addToCartConversion", 0) or 0,
            "funnel_cr_order": r.get("cartToOrderConversion", 0) or 0,
            "funnel_buyout_pct": r.get("buyoutPercent", 0) or 0,
        }
    return result


def write_to_all_articles(sh, target_date, orders, sales, funnel, ads_stats,
                           stocks=None, finance_rows=None, storage_rows=None,
                           dry_run=True):

    ws = sh.worksheet(TARGET_SHEET_NAME)
    grid = ws.get_all_values()

    blocks = find_product_blocks(grid)
    print(f"[Все артикулы] Найдено активных блоков товаров: {len(blocks)}")

    date_col0 = find_date_column(grid, target_date)
    if date_col0 is None:
        print(f"[Все артикулы] ОШИБКА: не найден столбец для даты "
              f"{target_date.strftime('%d.%m.%Y')} в шапке листа '{TARGET_SHEET_NAME}'. "
              f"Проверь, что дропдаун месяца/года на листе выставлен на нужный месяц. "
              f"Запись пропущена.")
        return

    orders_agg = aggregate_orders(orders)
    sales_agg = aggregate_sales(sales)
    ads_agg = aggregate_ads(ads_stats)
    funnel_agg = aggregate_funnel(funnel)
    cost_reference = fetch_cost_reference(sh)  # nmID -> {cost_per_unit, commission_pct}

    stocks_agg = {}
    finance_agg = {}
    storage_agg = {}
    if stocks is not None:
        from wb_new_fetches import aggregate_stocks
        stocks_agg = aggregate_stocks(stocks)
    if finance_rows is not None:
        from wb_new_fetches import aggregate_finance
        finance_agg = aggregate_finance(finance_rows)
    if storage_rows is not None:
        from wb_new_fetches import aggregate_paid_storage
        storage_agg = aggregate_paid_storage(storage_rows)

    updates = []
    written_nm = 0
    for nm, label_row0 in blocks.items():
        row_data = {}
        row_data.update(orders_agg.get(nm, {}))
        row_data.update(sales_agg.get(nm, {}))
        row_data.update(ads_agg.get(nm, {}))
        row_data.update(funnel_agg.get(nm, {}))
        row_data.update(stocks_agg.get(nm, {}))
        row_data.update(finance_agg.get(nm, {}))
        row_data.update(storage_agg.get(nm, {}))
        if not row_data:
            continue


        ads_cost_abs = abs(row_data.get("ads_cost", 0) or 0)
        ads_orders_rub = row_data.get("ads_orders_rub", 0) or 0
        orders_rub_before_spp = row_data.get("orders_rub_before_spp", 0) or 0
        sales_rub = row_data.get("sales_rub", 0) or 0
        row_data["ads_drr_pct"] = round(ads_cost_abs / ads_orders_rub * 100, 2) if ads_orders_rub else 0
        row_data["ads_drr_all_orders_pct"] = round(ads_cost_abs / orders_rub_before_spp * 100, 2) if orders_rub_before_spp else 0
        row_data["ads_drr_all_sales_pct"] = round(ads_cost_abs / sales_rub * 100, 2) if sales_rub else 0

        ref = cost_reference.get(nm, {})
        cost_per_unit = ref.get("cost_per_unit")
        commission_pct = ref.get("commission_pct")
        sales_qty = row_data.get("sales_qty", 0) or 0

        if commission_pct is not None:
            row_data["commission_rub"] = round(-(sales_rub * commission_pct / 100), 2)

        if cost_per_unit is not None:
            row_data["cost_total_rub"] = round(-(cost_per_unit * sales_qty), 2)

        if cost_per_unit is not None and commission_pct is not None:
            margin_rub = (
                sales_rub
                + (row_data.get("returns_rub", 0) or 0)
                + (row_data.get("ads_cost", 0) or 0)
                + (row_data.get("logistics_rub", 0) or 0)
                + (row_data.get("storage_rub", 0) or 0)
                + row_data["cost_total_rub"]
                + row_data["commission_rub"]
            )
            row_data["margin_rub"] = round(margin_rub, 2)
            row_data["margin_pct"] = round(margin_rub / sales_rub * 100, 2) if sales_rub else 0

        for key, value in row_data.items():
            offset = METRIC_OFFSETS.get(key)
            if offset is None:
                continue
            target_row0 = label_row0 + offset
            updates.append({
                "range": f"{col_letter(date_col0)}{target_row0 + 1}",
                "values": [[value]],
            })
        written_nm += 1

    if not updates:
        print(f"[Все артикулы] Нет данных для записи за {target_date.strftime('%d.%m.%Y')}.")
        return

    if dry_run:
        print(f"[Все артикулы] DRY_RUN: подготовлено {len(updates)} ячеек для "
              f"{written_nm} товаров (столбец {col_letter(date_col0)}, "
              f"дата {target_date.strftime('%d.%m')}). Ничего не записываю.")
        for u in updates[:10]:
            print(f"    {u['range']} = {u['values'][0][0]}")
        if len(updates) > 10:
            print(f"    ... и ещё {len(updates) - 10} ячеек")
        return

    BATCH_SIZE = 500
    for i in range(0, len(updates), BATCH_SIZE):
        chunk = updates[i:i + BATCH_SIZE]
        ws.batch_update(chunk, value_input_option="USER_ENTERED")

    print(f"[Все артикулы] Записано {len(updates)} ячеек для {written_nm} товаров "
          f"(столбец {col_letter(date_col0)}, дата {target_date.strftime('%d.%m')}).")