from typing import List
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import pandas as pd
import io
import re
import xlsxwriter
import asyncio

app = FastAPI(title="Fin Analysis Generator (Improved)")

from fastapi.middleware.cors import CORSMiddleware

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Разрешает запросы с любых сайтов/HTML-файлов
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

CONFIRM_SET = {"да", "yes", "+", "compliant", "согласен", "согласна", "agree", "agreed", "ok", "ок"}


def clean(x):
    return "" if pd.isna(x) else str(x).strip()


def clean_lower(x):
    return clean(x).lower()


def is_confirmed(x) -> bool:
    return clean_lower(x) in CONFIRM_SET


def find_row(df, col_idx, pattern: str):
    idx = df[
        df[col_idx].astype(str).str.contains(pattern, na=False, case=False, regex=True)
    ].index
    return int(idx[0]) if len(idx) else None


def detect_currency(df) -> str:
    text = "\n".join([clean(x) for x in df.iloc[:60, :].astype(str).values.flatten()])
    m = re.search(r"\b(USD|UZS|EUR|RUB|CNY|GBP)\b", text, flags=re.IGNORECASE)
    return m.group(1).upper() if m else ""


def parse_kp(file_bytes: bytes):
    # Читаем сразу ВСЕ листы
    all_sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=None)

    # Умный поиск нужного листа
    best_sheet_name = list(all_sheets.keys())[0]
    if len(all_sheets) > 1:
        max_score = -1
        for sheet_name, df_sheet in all_sheets.items():
            # Берем первые 50 строк для анализа
            text_sample = ' '.join([str(val).lower() for val in df_sheet.head(50).values.flatten()])

            score = 0
            # Обязательные поля - даем им огромный вес!
            if re.search(r"закупочной процедуры|procurement|tender", text_sample):
                score += 100
            if re.search(r"участник|participant|bidder|vendor", text_sample):
                score += 100

            # Второстепенные слова
            anchors = [
                "наименование",
                "описание",
                "кол-во",
                "количество",
                "цена",
                "сумма",
                "итого",
                "total",
                "price",
                "qty",
            ]
            score += sum(5 for anchor in anchors if anchor in text_sample)

            if score > max_score:
                max_score = score
                best_sheet_name = sheet_name

    df = all_sheets[best_sheet_name]
    currency = detect_currency(df)

    # Функция для гибкого поиска (ищет в первых 4-х колонках, а не только жестко в 0-й)
    def find_field(pattern):
        for c in range(min(4, len(df.columns))):
            idx = df[
                df[c]
                .astype(str)
                .str.contains(pattern, na=False, case=False, regex=True)
            ].index
            if len(idx):
                r = int(idx[0])
                # Значение ищем в ячейках правее
                for val_col in range(c + 1, len(df.columns)):
                    val = clean(df.iloc[r, val_col])
                    if val and val != ":":
                        return r, val
                # Фолбэк, если правее почему-то пусто
                return r, clean(df.iloc[r, 2])
        return None, None

    zp_row, zp_name = find_field(
        r"Название Закупочной процедуры|Name of.*[Pp]rocurement|Tender Name|Procurement procedure name"
    )
    if zp_row is None:
        raise ValueError(
            f"В КП (на листе '{best_sheet_name}') не найдено Название процедуры."
        )

    p_row, participant = find_field(
        r"^Участник$|^Participant$|^Competition participant$|Участник Закупочной процедуры|Bidder|Vendor"
    )
    if p_row is None:
        raise ValueError(
            f"В КП (на листе '{best_sheet_name}') не найдено имя Участника."
        )

    # Ищем шапку (Наименование/Name) в первых 4 колонках
    # Ищем шапку таблицы (строку, где есть хотя бы 2 ключевых слова)
    header_row = None
    for r in range(min(50, len(df))):
        row_text = ' '.join([str(val).lower() for val in df.iloc[r].values])
        matches = 0
        if re.search(
            r"наименование|name|description|описание|товар|услуг|item", row_text
        ):
            matches += 1
        if re.search(r"кол-во|количество|qty|quantity|объем", row_text):
            matches += 1
        if re.search(r"цена|price|расценка", row_text):
            matches += 1

        if matches >= 2:
            header_row = r
            break

    if header_row is None:
        raise ValueError(
            f"В КП (на листе '{best_sheet_name}') не найдена шапка таблицы"
        )

    headers = df.iloc[header_row].astype(str).str.lower()
    col_name = headers[
        headers.str.contains(
            r"наименование|описание|name|description|услуг|товар|item",
            na=False,
            regex=True,
        )
    ].index.min()
    col_uom = headers[
        headers.str.contains(
            r"ед\.? изм|измерения|uom|unit|measure", na=False, regex=True
        )
    ].index.min()
    col_qty = headers[
        headers.str.contains(
            r"кол-во|количество|qty|quantity|объем|q-ty|q'ty", na=False, regex=True
        )
    ].index.min()
    col_price = headers[
        headers.str.contains(r"цена|price|расценка|unit price", na=False, regex=True)
    ].index.min()

    col_name = int(col_name) if pd.notna(col_name) else 1
    col_uom = int(col_uom) if pd.notna(col_uom) else None  # нет колонки — пишем пусто
    if pd.isna(col_qty):
        raise ValueError(
            f"В КП (лист '{best_sheet_name}') не найдена колонка «Количество». "
            f"Убедитесь что в шапке таблицы есть заголовок: Quantity / Кол-во / Количество / Qty"
        )
    col_qty = int(col_qty)
    if pd.isna(col_price):
        raise ValueError(
            f"В КП (лист '{best_sheet_name}') не найдена колонка «Цена». "
            f"Убедитесь что в шапке таблицы есть заголовок: Price / Цена / Unit price"
        )
    col_price = int(col_price)

    goods = []
    has_lots = False
    i = header_row + 1

    stop_phrases = [
        "всего",
        "общая стоимость",
        "срок поставки",
        "условия оплаты",
        "гарантия",
        "заверяется",
        "условия поставки",
        "примечание",
        "total",
        "delivery time",
        "payment terms",
        "terms of payment",
        "warranty",
        "delivery terms",
        "terms of delivery",
        "note",
        "remarks",
        "certified",
        "signature",
        "lead time",
    ]

    while i < len(df):
        name = clean(df.iloc[i, col_name])
        no_raw = df.iloc[i, 0]
        no = clean(no_raw)

        display_name = name if name else no
        name_lower = display_name.lower()

        if not display_name and pd.isna(
            pd.to_numeric(df.iloc[i, col_qty], errors="coerce")
        ):
            i += 1
            continue

        is_total = "итого" in name_lower or "total" in name_lower

        if any(phrase in name_lower for phrase in stop_phrases) and not is_total:
            break

        if is_total:
            if has_lots:
                goods.append({"type": "lot_total", "name": display_name})
                i += 1
                continue
            else:
                break

        qty_val = clean(df.iloc[i, col_qty])
        is_qty_empty = qty_val == "" or pd.isna(pd.to_numeric(qty_val, errors="coerce"))
        is_no_empty_or_text = not str(no_raw).strip().isdigit()

        if display_name and is_qty_empty and is_no_empty_or_text:
            is_explicit_lot = "лот" in name_lower or "lot" in name_lower
            has_items = any(g["type"] == "item" for g in goods)

            if is_explicit_lot:
                # Если предыдущий лот есть и у него нет lot_total — добавляем авто-итог
                if has_lots and has_items:
                    # Проверяем: последний элемент — не lot_total?
                    last_non_empty = next(
                        (g for g in reversed(goods) if g["type"] in ("item", "lot_total", "lot_header")),
                        None
                    )
                    if last_non_empty and last_non_empty["type"] != "lot_total":
                        goods.append({"type": "lot_total", "name": "Итого", "auto": True})
                has_lots = True
                goods.append({"type": "lot_header", "name": display_name})
            elif not has_items and len(display_name) < 100:
                has_lots = True
                goods.append({"type": "lot_header", "name": display_name})
            else:
                break

            i += 1
            continue

        if no or name:
            uom = clean(df.iloc[i, col_uom]) if col_uom is not None else ""
            qty = pd.to_numeric(df.iloc[i, col_qty], errors="coerce")
            price = pd.to_numeric(df.iloc[i, col_price], errors="coerce")
            goods.append(
                {
                    "type": "item",
                    "name": display_name,
                    "uom": uom,
                    "qty": qty,
                    "price": price,
                    "no": no,
                }
            )

        i += 1

    if not goods:
        raise ValueError("В КП не найдено товарных строк")

    # Авто-итог для последнего лота если нет явного lot_total
    if has_lots:
        last_non_empty = next(
            (g for g in reversed(goods) if g["type"] in ("item", "lot_total", "lot_header")),
            None
        )
        if last_non_empty and last_non_empty["type"] == "item":
            goods.append({"type": "lot_total", "name": "Итого", "auto": True})

    # --- ДИНАМИЧЕСКИЙ ПАРСИНГ УСЛОВИЙ ---
    # Стоп-строки которые НЕ являются условиями
    COND_STOP = {
        "positions marked", "full name", "signature", "stamped",
        "note:", "примечание", "заверяется", "подпись"
    }
    # Известные метки → русский перевод для заголовка в отчёте
    KNOWN_LABELS = {
        r"срок поставки|delivery time|delivery period|lead time": "Срок поставки",
        r"условия оплаты|payment terms|terms of payment": "Условия оплаты",
        r"гарантия|warranty|guarantee": "Гарантия",
        r"условия поставки|delivery terms|terms of delivery|incoterms": "Условия поставки",
    }

    cond_final = {}  # {ru_label: value}

    # Сканируем строки ПОСЛЕ товарной таблицы
    for idx in range(i, len(df)):
        # Ищем ярлык в первых 3 колонках
        c_label = None
        label_text_raw = ""
        for c in range(min(3, len(df.columns))):
            val = clean(df.iloc[idx, c])
            if not val:
                continue
            val_lower = val.lower()
            # Пропускаем стоп-строки
            if any(stop in val_lower for stop in COND_STOP):
                break
            # Это похоже на ярлык условия если:
            # - в строке есть хоть одно непустое значение правее
            # - сама ячейка не выглядит как длинный текст (< 60 символов)
            if len(val) < 80:
                row_vals_check = [
                    clean(x) for x in df.iloc[idx].values[c + 1:] if clean(x) != ""
                ]
                if row_vals_check:
                    c_label = c
                    label_text_raw = val
                    break

        if c_label is None:
            continue

        label_lower = label_text_raw.lower()

        # Определяем русское название
        ru_lab = None
        for pattern, ru_name in KNOWN_LABELS.items():
            if re.search(pattern, label_lower):
                ru_lab = ru_name
                break
        # Если не распознали — используем оригинальный текст ярлыка как есть
        if ru_lab is None:
            ru_lab = label_text_raw.strip()

        # Пропускаем дубли
        if ru_lab in cond_final:
            continue

        # Берём значения правее ярлыка
        row_vals = [
            clean(x) for x in df.iloc[idx].values[c_label + 1:] if clean(x) != ""
        ]

        if not row_vals:
            continue

        if len(row_vals) == 1:
            if is_confirmed(row_vals[0]):
                if ":" in label_text_raw:
                    cond_final[ru_lab] = label_text_raw.split(":", 1)[1].strip()
                else:
                    cond_final[ru_lab] = "Согласен (требование не указано)"
            else:
                cond_final[ru_lab] = row_vals[0]
        else:
            customer_text = row_vals[0]
            participant_text = row_vals[-1]
            cond_final[ru_lab] = (
                customer_text if is_confirmed(participant_text) else participant_text
            )

    return zp_name, participant, goods, cond_final, currency


def rank_colors_with_ties(prices):
    vals = [
        (i, float(p)) for i, p in enumerate(prices) if p is not None and pd.notna(p)
    ]
    if not vals:
        return [0] * len(prices)

    uniq = sorted(set(v for _, v in vals))
    level = {val: rank + 1 for rank, val in enumerate(uniq[:3])}

    out = [0] * len(prices)
    for i, p in vals:
        out[i] = level.get(p, 0)
    return out


def generate_excel(parsed, check_anomaly=False, anomaly_percent=50):
    zp_name = parsed[0]["zp"]
    # ... дальше без изменений до проверки цен
    n_part = len(parsed)
    n_cols_total = n_part + 1  # Участники + колонка Медианы
    goods_struct = parsed[0]["goods"]

    uom = next(
        (g["uom"] for g in goods_struct if g["type"] == "item" and g.get("uom")), "Шт"
    )
    currency = next((p["currency"] for p in parsed if p["currency"]), "USD")

    out = io.BytesIO()
    wb = xlsxwriter.Workbook(out, {"in_memory": True})
    ws = wb.add_worksheet("Фин анализ")

    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)

    # 3. Шрифт Arial Narrow везде
    font = "Arial Narrow"
    blue = "#9FD5E4"
    med_bg = "#EAEAEA"  # 2. Цвет для медианы (светло-серый)

    fmt_title = wb.add_format(
        {
            "font_name": font,
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "font_size": 14,
        }
    )
    fmt_sub = wb.add_format({"font_name": font, "bold": True, "font_size": 11})
    fmt_hdr = wb.add_format(
        {
            "font_name": font,
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": blue,
            "border": 1,
            "text_wrap": True,
            "font_size": 10,
        }
    )
    fmt_hdr_small = wb.add_format(
        {
            "font_name": font,
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": blue,
            "border": 1,
            "text_wrap": True,
            "font_size": 9,
        }
    )
    fmt_cell = wb.add_format({"font_name": font, "border": 1, "font_size": 9})
    fmt_cell_c = wb.add_format(
        {
            "font_name": font,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "font_size": 9,
        }
    )
    fmt_money_c = wb.add_format(
        {
            "font_name": font,
            "border": 1,
            "num_format": "#,##0.00",
            "align": "center",
            "valign": "vcenter",
            "font_size": 9,
        }
    )
    fmt_total = wb.add_format(
        {
            "font_name": font,
            "border": 1,
            "bold": True,
            "font_size": 9,
            "align": "right",
            "num_format": "#,##0.00",
        }
    )
    fmt_lot_hdr = wb.add_format(
        {
            "font_name": font,
            "bold": True,
            "bg_color": "#D9D9D9",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "font_size": 10,
        }
    )
    fmt_cond_lbl = wb.add_format(
        {
            "font_name": font,
            "border": 1,
            "bold": True,
            "font_size": 10,
            "align": "center",
            "valign": "vcenter",
        }
    )
    fmt_cond_txt = wb.add_format(
        {
            "font_name": font,
            "border": 1,
            "font_size": 9,
            "text_wrap": True,
            "valign": "top",
        }
    )
    fmt_min_title = wb.add_format(
        {"font_name": font, "align": "center", "valign": "vcenter", "font_size": 12}
    )

    ranks_colors = ["#C6EFCE", "#FFEB9C", "#F8CBAD"]
    fmt_ranks = [
        wb.add_format(
            {
                "font_name": font,
                "border": 1,
                "num_format": "#,##0.00",
                "align": "center",
                "valign": "vcenter",
                "font_size": 9,
                "bg_color": c,
            }
        )
        for c in ranks_colors
    ]

    # Специальные форматы для колонки "Медиана"
    fmt_hdr_med = wb.add_format(
        {
            "font_name": font,
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": med_bg,
            "border": 1,
            "text_wrap": True,
            "font_size": 9,
        }
    )
    fmt_cell_c_med = wb.add_format(
        {
            "font_name": font,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "font_size": 9,
            "bg_color": med_bg,
        }
    )
    fmt_money_c_med = wb.add_format(
        {
            "font_name": font,
            "border": 1,
            "num_format": "#,##0.00",
            "align": "center",
            "valign": "vcenter",
            "font_size": 9,
            "bg_color": med_bg,
        }
    )
    fmt_total_med = wb.add_format(
        {
            "font_name": font,
            "border": 1,
            "bold": True,
            "font_size": 9,
            "align": "right",
            "bg_color": med_bg,
            "num_format": "#,##0.00",
        }
    )
    fmt_cond_txt_med = wb.add_format(
        {
            "font_name": font,
            "border": 1,
            "font_size": 9,
            "text_wrap": True,
            "valign": "top",
            "bg_color": med_bg,
        }
    )
    fmt_anomaly = wb.add_format(
        {
            "font_name": font,
            "border": 1,
            "num_format": "#,##0.00",
            "align": "center",
            "valign": "vcenter",
            "font_size": 9,
            "bg_color": "#FFC7CE",
            "font_color": "#9C0006",
        }
    )

    col_no, col_name = 0, 1
    first_block, step = 2, 4

    ws.set_column(col_no, col_no, 4)
    ws.set_column(col_name, col_name, 30)
    for i in range(n_cols_total):
        s = first_block + i * step
        ws.set_column(s, s, 12)
        ws.set_column(s + 1, s + 1, 10)
        ws.set_column(s + 2, s + 2, 12)
        if i < n_cols_total - 1:
            ws.set_column(s + 3, s + 3, 2)

    last_col = first_block + (n_cols_total - 1) * step + 2

    ws.merge_range(0, 0, 0, last_col, "Результаты Финансовой оценки", fmt_title)
    ws.write(2, 0, f"Название Закупочной процедуры: {zp_name}", fmt_sub)
    ws.merge_range(5, 0, 5, 1, "Этап 1", fmt_hdr)
    ws.merge_range(6, 0, 6, 1, "Участник Запроса цен", fmt_hdr_small)

    for i in range(n_cols_total):
        s = first_block + i * step
        if i < n_part:
            title = parsed[i]["participant"]
            ws.merge_range(6, s, 6, s + 2, title, fmt_hdr_small)
            ws.write(7, s, f"Количество\nПродукции, {uom}", fmt_hdr_small)
            ws.write(7, s + 1, f"Цена за ед.,\n({currency}),\nбез НДС", fmt_hdr_small)
            ws.write(
                7, s + 2, f"Стоимость\nВСЕГО, ({currency}) без\nНДС *", fmt_hdr_small
            )
        else:
            title = "Медиана"
            ws.merge_range(6, s, 6, s + 2, title, fmt_hdr_med)
            ws.write(7, s, f"Количество\nПродукции, {uom}", fmt_hdr_med)
            ws.write(7, s + 1, f"Медиана\n({currency}),\nбез НДС", fmt_hdr_med)
            ws.write(
                7, s + 2, f"Стоимость\nВСЕГО, ({currency}) без\nНДС *", fmt_hdr_med
            )

    ws.write(7, col_no, "№", fmt_hdr)
    ws.write(7, col_name, "Наименование Продукции", fmt_hdr)

    data_start = 8
    current_lot_start = None
    all_totals_rows = []

    for r, row_data in enumerate(goods_struct):
        rr = data_start + r

        if row_data["type"] == "lot_header":
            ws.merge_range(rr, 0, rr, last_col, row_data["name"], fmt_lot_hdr)
            current_lot_start = rr + 1

        elif row_data["type"] == "lot_total":
            ws.merge_range(rr, col_no, rr, col_name, row_data["name"], fmt_total)
            all_totals_rows.append(rr)
            for i in range(n_cols_total):
                s = first_block + i * step
                c_fmt_tot = fmt_total if i < n_part else fmt_total_med
                ws.write(rr, s, "", c_fmt_tot)
                ws.write(rr, s + 1, "", c_fmt_tot)
                col_total = s + 2
                if current_lot_start:
                    ws.write_formula(
                        rr,
                        col_total,
                        f"=SUM({xlsxwriter.utility.xl_col_to_name(col_total)}{current_lot_start+1}:{xlsxwriter.utility.xl_col_to_name(col_total)}{rr})",
                        c_fmt_tot,
                    )
                else:
                    ws.write(rr, col_total, "", c_fmt_tot)

        elif row_data["type"] == "item":
            ws.write(rr, col_no, row_data.get("no", ""), fmt_cell_c)
            ws.write(rr, col_name, row_data["name"], fmt_cell)

            prices = [parsed[i]["goods"][r]["price"] for i in range(n_part)]
            ranks = rank_colors_with_ties(prices)

            # --- ВЫЧИСЛЯЕМ МЕДИАНУ ДЛЯ ПОИСКА АНОМАЛИЙ (50%) ---
            # --- ВЫЧИСЛЯЕМ МЕДИАНУ ДЛЯ ПОИСКА АНОМАЛИЙ ---
            valid_prices = [float(p) for p in prices if pd.notna(p) and p != ""]
            med_val = pd.Series(valid_prices).median() if valid_prices else 0

            # Считаем границы на основе переданного процента
            lower_bound = 1 - (anomaly_percent / 100)
            upper_bound = 1 + (anomaly_percent / 100)

            for i in range(n_part):
                s = first_block + i * step
                qty = parsed[i]["goods"][r]["qty"]
                price = parsed[i]["goods"][r]["price"]

                if pd.notna(qty) and qty != "":
                    ws.write_number(rr, s, float(qty), fmt_cell_c)
                else:
                    ws.write_blank(rr, s, None, fmt_cell_c)

                if pd.notna(price) and price != "":
                    num_price = float(price)

                    # ПРОВЕРКА НА АНОМАЛИЮ (если стоит галочка check_anomaly)
                    if (
                        check_anomaly
                        and med_val > 0
                        and (
                            num_price > med_val * upper_bound
                            or num_price < med_val * lower_bound
                        )
                    ):
                        pf = fmt_anomaly
                    else:
                        pf = (
                            fmt_ranks[ranks[i] - 1]
                            if ranks[i] in [1, 2, 3]
                            else fmt_money_c
                        )

                    ws.write_number(rr, s + 1, num_price, pf)
                    ws.write_formula(
                        rr,
                        s + 2,
                        f"={xlsxwriter.utility.xl_col_to_name(s)}{rr+1}*{xlsxwriter.utility.xl_col_to_name(s+1)}{rr+1}",
                        fmt_money_c,
                    )
                else:
                    ws.write_blank(rr, s + 1, None, fmt_money_c)
                    ws.write_blank(rr, s + 2, None, fmt_money_c)

            # Медиана
            s_med = first_block + n_part * step
            qty_med = parsed[0]["goods"][r]["qty"]

            if pd.notna(qty_med) and qty_med != "":
                ws.write_number(rr, s_med, float(qty_med), fmt_cell_c_med)
            else:
                ws.write_blank(rr, s_med, None, fmt_cell_c_med)

            price_cells = [
                f"{xlsxwriter.utility.xl_col_to_name(first_block + i * step + 1)}{rr+1}"
                for i in range(n_part)
            ]
            ws.write_formula(
                rr, s_med + 1, f"=MEDIAN({','.join(price_cells)})", fmt_money_c_med
            )
            ws.write_formula(
                rr,
                s_med + 2,
                f"={xlsxwriter.utility.xl_col_to_name(s_med)}{rr+1}*{xlsxwriter.utility.xl_col_to_name(s_med+1)}{rr+1}",
                fmt_money_c_med,
            )

    total_row = data_start + len(goods_struct)
    ws.merge_range(total_row, col_no, total_row, col_name, "ВСЕГО", fmt_total)

    for i in range(n_cols_total):
        s = first_block + i * step
        col_total = s + 2
        c_fmt_tot = fmt_total if i < n_part else fmt_total_med

        ws.write(total_row, s, "", c_fmt_tot)
        ws.write(total_row, s + 1, "", c_fmt_tot)
        if all_totals_rows:
            formula = "=" + "+".join(
                f"{xlsxwriter.utility.xl_col_to_name(col_total)}{tr+1}"
                for tr in all_totals_rows
            )
        else:
            formula = f"=SUM({xlsxwriter.utility.xl_col_to_name(col_total)}{data_start+1}:{xlsxwriter.utility.xl_col_to_name(col_total)}{total_row})"
        ws.write_formula(total_row, col_total, formula, c_fmt_tot)

    has_conditions = any(v for p in parsed for v in p["cond_final"].values())
    cond_offset = 0

    if has_conditions:
        cond_start = total_row + 3
        # Собираем все уникальные ярлыки условий из всех участников (сохраняя порядок)
        seen = {}
        for p in parsed:
            for lab in p["cond_final"].keys():
                if lab not in seen:
                    seen[lab] = True
        labels = list(seen.keys())

        for j, lab in enumerate(labels):
            rr = cond_start + j
            ws.set_row(rr, 50)
            ws.merge_range(rr, 0, rr, 1, lab, fmt_cond_lbl)
            for i in range(n_part):
                s = first_block + i * step
                ws.merge_range(
                    rr, s, rr, s + 2, parsed[i]["cond_final"].get(lab, ""), fmt_cond_txt
                )

            s_med = first_block + n_part * step
            ws.merge_range(rr, s_med, rr, s_med + 2, "-", fmt_cond_txt_med)

        cond_offset = len(labels) + 4
    else:
        cond_offset = 3

    min_title_row = total_row + cond_offset
    ws.merge_range(
        min_title_row,
        0,
        min_title_row,
        10,
        "Сводка: Победитель и Резервный поставщик",
        fmt_min_title,
    )

    min_hdr = min_title_row + 1
    ws.set_row(min_hdr, 35)
    ws.write(min_hdr, 0, "№", fmt_hdr)
    ws.write(min_hdr, 1, "Наименование Продукции", fmt_hdr)
    ws.write(min_hdr, 2, f"Количество\nПродукции, {uom}", fmt_hdr_small)

    ws.write(min_hdr, 3, f"Цена\nпобедителя, ({currency}),\nбез НДС", fmt_hdr_small)
    ws.write(min_hdr, 4, f"Стоимость\nВСЕГО, ({currency}) без\nНДС *", fmt_hdr_small)
    ws.merge_range(min_hdr, 5, min_hdr, 6, "Победители", fmt_hdr)

    ws.write(
        min_hdr,
        7,
        f"Цена Резервного\nпоставщика, ({currency}),\nбез НДС",
        fmt_hdr_small,
    )
    ws.write(min_hdr, 8, f"Стоимость\nВСЕГО, ({currency}) без\nНДС *", fmt_hdr_small)
    ws.merge_range(min_hdr, 9, min_hdr, 10, "Резервный\nпоставщик", fmt_hdr)

    min_start = min_hdr + 1
    min_lot_start = None
    min_totals_rows = []

    for r, row_data in enumerate(goods_struct):
        rr = min_start + r
        ws.set_row(rr, 18)

        if row_data["type"] == "lot_header":
            ws.merge_range(rr, 0, rr, 10, row_data["name"], fmt_lot_hdr)
            min_lot_start = rr + 1

        elif row_data["type"] == "lot_total":
            ws.merge_range(rr, 0, rr, 1, row_data["name"], fmt_total)
            min_totals_rows.append(rr)
            ws.write(rr, 2, "", fmt_total)
            ws.write(rr, 3, "", fmt_total)

            if min_lot_start:
                ws.write_formula(rr, 4, f"=SUM(E{min_lot_start+1}:E{rr})", fmt_total)
            else:
                ws.write(rr, 4, "", fmt_total)

            ws.merge_range(rr, 5, rr, 6, "", fmt_total)
            ws.write(rr, 7, "", fmt_total)

            if min_lot_start:
                ws.write_formula(rr, 8, f"=SUM(I{min_lot_start+1}:I{rr})", fmt_total)
            else:
                ws.write(rr, 8, "", fmt_total)

            ws.merge_range(rr, 9, rr, 10, "", fmt_total)

        elif row_data["type"] == "item":
            ws.write(rr, 0, row_data.get("no", ""), fmt_cell_c)
            ws.write(rr, 1, row_data["name"], fmt_cell)

            qty0 = parsed[0]["goods"][r]["qty"]

            if pd.notna(qty0) and qty0 != "":
                ws.write_number(rr, 2, float(qty0), fmt_cell_c)
            else:
                ws.write_blank(rr, 2, None, fmt_cell_c)

            valid = [
                (parsed[i]["participant"], float(parsed[i]["goods"][r]["price"]))
                for i in range(n_part)
                if pd.notna(parsed[i]["goods"][r]["price"])
                and parsed[i]["goods"][r]["price"] != ""
            ]
            valid.sort(key=lambda x: x[1])

            if len(valid) > 0:
                winner, minp = valid[0]
                ws.write_number(rr, 3, minp, fmt_money_c)
                ws.write_formula(rr, 4, f"=C{rr+1}*D{rr+1}", fmt_money_c)
                ws.merge_range(rr, 5, rr, 6, winner, fmt_cell_c)
            else:
                ws.write_blank(rr, 3, None, fmt_cell)
                ws.write_blank(rr, 4, None, fmt_cell)
                ws.merge_range(rr, 5, rr, 6, "", fmt_cell)

            if len(valid) > 1:
                reserve, resp = valid[1]
                ws.write_number(rr, 7, resp, fmt_money_c)
                ws.write_formula(rr, 8, f"=C{rr+1}*H{rr+1}", fmt_money_c)
                ws.merge_range(rr, 9, rr, 10, reserve, fmt_cell_c)
            else:
                ws.write_blank(rr, 7, None, fmt_cell)
                ws.write_blank(rr, 8, None, fmt_cell)
                ws.merge_range(rr, 9, rr, 10, "", fmt_cell)

    min_total_row = min_start + len(goods_struct)
    ws.set_row(min_total_row, 18)
    ws.merge_range(min_total_row, 0, min_total_row, 1, "ВСЕГО", fmt_total)
    ws.write(min_total_row, 2, "", fmt_total)
    ws.write(min_total_row, 3, "", fmt_total)

    if min_totals_rows:
        formula_win = "=" + "+".join(f"E{tr+1}" for tr in min_totals_rows)
        formula_res = "=" + "+".join(f"I{tr+1}" for tr in min_totals_rows)
    else:
        formula_win = f"=SUM(E{min_start+1}:E{min_total_row})"
        formula_res = f"=SUM(I{min_start+1}:I{min_total_row})"

    ws.write_formula(min_total_row, 4, formula_win, fmt_total)
    ws.merge_range(min_total_row, 5, min_total_row, 6, "", fmt_total)
    ws.write(min_total_row, 7, "", fmt_total)
    ws.write_formula(min_total_row, 8, formula_res, fmt_total)
    ws.merge_range(min_total_row, 9, min_total_row, 10, "", fmt_total)

    # 1. Легенда цветов
    legend_row = min_total_row + 2
    fmt_leg_title = wb.add_format({"font_name": font, "bold": True, "font_size": 10})
    fmt_leg_text = wb.add_format({"font_name": font, "font_size": 10})

    ws.write(legend_row, 0, "Обозначения цветов:", fmt_leg_title)

    ws.write(legend_row + 1, 0, "", fmt_ranks[0])
    ws.write(legend_row + 1, 1, "- 1 место по минимальной цене", fmt_leg_text)

    ws.write(legend_row + 2, 0, "", fmt_ranks[1])
    ws.write(legend_row + 2, 1, "- 2 место по минимальной цене", fmt_leg_text)

    ws.write(legend_row + 3, 0, "", fmt_ranks[2])
    ws.write(legend_row + 3, 1, "- 3 место по минимальной цене", fmt_leg_text)

    if check_anomaly:
        ws.write(legend_row + 4, 0, "", fmt_anomaly)
        ws.write(
            legend_row + 4,
            1,
            f"- Отклонение от медианы более чем на {anomaly_percent}% (Аномалия)",
            fmt_leg_text,
        )

    wb.close()
    out.seek(0)
    return out


def process_file(file_bytes: bytes):
    return parse_kp(file_bytes)


@app.post("/analyze")
async def analyze(
    files: List[UploadFile] = File(...),
    check_anomaly: bool = Form(False),
    anomaly_percent: int = Form(50),
):
    if not files:
        raise HTTPException(status_code=400, detail="Загрузи минимум 1 КП")

    parsed = []
    for f in files:
        if not f.filename.lower().endswith((".xls", ".xlsx")):
            raise HTTPException(
                status_code=400,
                detail=f"Файл {f.filename} не является Excel-документом",
            )

        content = await f.read()

        try:
            zp, participant, goods, cond_final, curr = await asyncio.to_thread(
                process_file, content
            )
        except Exception as e:
            raise HTTPException(
                status_code=400, detail=f"Ошибка в файле {f.filename}: {str(e)}"
            )

        parsed.append(
            {
                "zp": zp,
                "participant": participant,
                "goods": goods,
                "cond_final": cond_final,
                "currency": curr,
            }
        )

    for p in parsed[1:]:
        if len(p["goods"]) != len(parsed[0]["goods"]):
            raise HTTPException(
                status_code=400, detail="У участников разное количество позиций."
            )

    # Передаем галочку и процент в функцию генерации
    out = await asyncio.to_thread(
        generate_excel, parsed, check_anomaly, anomaly_percent
    )

    return StreamingResponse(
        out,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=final_analysis.xlsx"},
    )
