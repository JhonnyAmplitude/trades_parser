import pandas as pd
import re
import json
from OperationDTO import OperationDTO

REQUIRED_COLUMNS = ["дата", "номер", "время", "курс сделки", "объём в валюте", "объём в сопряж"]

def to_num(x):
    try:
        return float(str(x).replace(",", ".")) if pd.notna(x) else 0
    except:
        return 0

def find_column_index(headers, *keywords):
    for i, h in enumerate(headers):
        if all(k in h for k in keywords):
            return i
    return None

def parse_date_cell(cell):
    try:
        return pd.to_datetime(str(cell).strip(), format="%d.%m.%y")
    except:
        return pd.NaT

def parse_forex_trades(file_path):
    df = pd.read_excel(file_path, header=None)
    results = []

    # 1) начало блока
    start_idx = None
    for i, row in df.iterrows():
        txt = " ".join(str(c).lower() for c in row if pd.notna(c))
        if "иностранная валюта" in txt:
            start_idx = i + 1
            break
    if start_idx is None:
        return []

    df_block = df.iloc[start_idx:].reset_index(drop=True)

    # 2) строка заголовков
    header_row = None
    for i, row in df_block.iterrows():
        txt = " ".join(str(c).lower() for c in row if pd.notna(c))
        if all(key in txt for key in REQUIRED_COLUMNS):
            header_row = i
            break
    if header_row is None:
        return []

    headers = df_block.iloc[header_row].astype(str).str.lower()

    # 3) индексы колонок
    col_number     = find_column_index(headers, "номер")
    col_buy_price  = find_column_index(headers, "курс", "покупка")
    col_sell_price = find_column_index(headers, "курс", "продажа")
    qty_cols       = [i for i,h in enumerate(headers) if "объём в валюте" in h]
    sum_cols       = [i for i,h in enumerate(headers) if "объём в сопряж" in h]
    col_buy_qty, col_sell_qty = (qty_cols + [None, None])[:2]
    col_buy_sum, col_sell_sum = (sum_cols + [None, None])[:2]
    col_exec_date  = find_column_index(headers, "дата", "соверш")
    col_exec_time  = find_column_index(headers, "время", "соверш")

    current_ticker  = ""
    current_currency = ""
    i = header_row + 1

    while i < len(df_block):
        row = df_block.iloc[i]

        # ► Пропускаем любые строки с "итого"
        if row.astype(str).str.lower().str.contains("итого").any():
            i += 1
            continue

        # ► Строка с новым тикером?
        ticker_cell = None
        for j, cell in enumerate(row):
            if isinstance(cell, str) and re.fullmatch(r"[A-Z]{6,}_TOM", cell.strip()):
                ticker_cell = (j, cell.strip())
                break
        if ticker_cell:
            raw = ticker_cell[1]
            current_ticker = raw.split("+")[0][:6]
            info = df_block.iloc[i].astype(str)
            idx = next((k for k,v in enumerate(info) if "сопряж. валюта:" in v.lower()), None)
            current_currency = info.iat[idx+2].strip() if idx and idx+2 < len(info) else ""
            i += 1
            continue

        # ► Пустая строка = конец
        if all(pd.isna(c) or str(c).strip()=="" for c in row):
            break

        # ► Обработка операции
        raw_date = row.iat[col_exec_date] if col_exec_date is not None else None
        raw_time = row.iat[col_exec_time] if col_exec_time is not None else None

        date_obj = parse_date_cell(raw_date)
        time_str = str(raw_time).strip() if pd.notna(raw_time) else "00:00:00"

        # поле `date_obj` гарантированно валидно (итого-строки мы уже отбросили)
        dt = pd.to_datetime(f"{date_obj.strftime('%Y-%m-%d')} {time_str}")
        date_str = dt.strftime("%Y-%m-%d %H:%M:%S")

        number = str(row.iat[col_number]).strip() if col_number is not None else ""

        bp = row.iat[col_buy_price]  if col_buy_price  is not None else None
        sp = row.iat[col_sell_price] if col_sell_price is not None else None

        if pd.notna(bp):
            op    = "currency_buy"
            price = to_num(bp)
            qty   = row.iat[col_buy_qty]
            sm    = row.iat[col_buy_sum]
        else:
            op    = "currency_sale"
            price = to_num(sp)
            qty   = row.iat[col_sell_qty]
            sm    = row.iat[col_sell_sum]

        dto = OperationDTO(
            date=date_str,
            operation_type=op,
            payment_sum=to_num(sm),
            currency=current_currency,
            ticker=current_ticker,
            isin="",
            price=price,
            quantity=to_num(qty),
            aci=0.0,
            comment="",
            operation_id=number
        )
        results.append(dto.to_dict())
        i += 1

    return results

if __name__ == "__main__":
    trades = parse_forex_trades("pensil.XLSX")
    print(json.dumps(trades, ensure_ascii=False, indent=2))
