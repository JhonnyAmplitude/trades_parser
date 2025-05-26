import json

import pandas as pd
import re
from OperationDTO import OperationDTO
from constants import (
    CURRENCY_DICT,
    SKIP_OPERATIONS,
    VALID_OPERATIONS, is_nonzero,
)
from utils import parse_date, parse_header_data, to_num, extract_isin, detect_operation_type


def parse_financial_operations(file_path: str) -> dict:
    # 1) Читаем весь лист как строки
    df = pd.read_excel(file_path, header=None, dtype=str)

    # 2) Склеиваем каждую строку для удобного поиска
    df['_row_txt'] = df.fillna('').agg(' '.join, axis=1).str.strip()

    # 3) Находим и «разливаем» валюту
    mask_cur = df['_row_txt'].isin(k for k in CURRENCY_DICT)
    df.loc[mask_cur, '_currency'] = df.loc[mask_cur, '_row_txt'].map(
        lambda x: CURRENCY_DICT.get(x, x)
    )
    df['_currency'] = df['_currency'].ffill().fillna('RUB')

    # 4) Ищем индекс строки-заголовка таблицы операций
    header_mask = (
        df['_row_txt'].str.lower().str.contains('дата') &
        df['_row_txt'].str.lower().str.contains('операция') &
        df['_row_txt'].str.lower().str.contains('зачислен')
    )
    hdr_idxs = df.index[header_mask]
    if hdr_idxs.empty:
        return {"account_id": None, "account_date_start": None, "date_start": None, "date_end": None, "operations": []}
    hdr_i = hdr_idxs[0]

    # 5) Парсим header_data из строк до hdr_i
    header_data = {
        "account_id": None,
        "account_date_start": None,
        "date_start": None,
        "date_end": None,
        "unknown_operations": []
    }
    for _, raw_row in df.iloc[:hdr_i].iterrows():
        # отбрасываем наши технические колонки
        # и склеиваем всё обратно в единый текст
        vals = [c for c in raw_row[:-2] if pd.notna(c)]
        row_str = " ".join(str(c).strip() for c in vals)
        parse_header_data(row_str, header_data)

    # 6) Собираем сам блок операций
    block = df.iloc[hdr_i:].reset_index(drop=True)
    headers = block.iloc[0].fillna('').astype(str).str.lower()

    def col(*keys):
        for i, h in enumerate(headers):
            if all(k in h for k in keys):
                return i
        return None

    ci = {
        "date":    col("дата"),
        "op":      col("операция"),
        "inc":     col("зачислен"),
        "exp":     col("списани"),
        "comment": col("примеч") or col("коммент")
    }

    data = block.iloc[1:].copy()
    data['_txt'] = data.fillna('').agg(' '.join, axis=1).str.lower()
    data = data[~data['_txt'].str.contains('итого')]

    # 7) Проходим по строкам и собираем DTO
    operations = []
    for _, row in data.iterrows():
        op_raw = str(row.iat[ci["op"]] or "").strip()
        if not op_raw or op_raw in SKIP_OPERATIONS or op_raw not in VALID_OPERATIONS:
            header_data["unknown_operations"].append(op_raw)
            continue

        dt = parse_date(row.iat[ci["date"]])
        if not dt:
            continue

        inc = row.iat[ci["inc"]] or ""
        exp = row.iat[ci["exp"]] or ""
        payment = to_num(inc if is_nonzero(inc) else exp)

        raw_comment = row.iat[ci["comment"]]
        comment = "" if pd.isna(raw_comment) or str(raw_comment).lower() == "nan" else str(raw_comment).strip()
        isin = extract_isin(comment)

        op_type = detect_operation_type(op_raw, inc, exp)
        currency = row['_currency']

        dto = OperationDTO(
            date=dt,
            operation_type=op_type,
            payment_sum=payment,
            currency=currency,
            ticker="",
            isin=isin,
            price=0.0,
            quantity=0,
            aci=0.0,
            comment=comment,
            operation_id=""
        )
        operations.append(dto)

    # 8) Формируем итоговый словарь
    return {
        "account_id": header_data.get("account_id"),
        "account_date_start": header_data.get("account_date_start"),
        "date_start": header_data.get("date_start"),
        "date_end": header_data.get("date_end"),
        "operations": [o.to_dict() for o in operations],
    }


