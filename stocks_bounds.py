import pandas as pd
import re
import json
from datetime import datetime
from typing import List, Any, Optional, Union

from OperationDTO import OperationDTO
from utils import to_num, find_column_index


# --- Парсинг тикера и ISIN из одной строки ---
def parse_ticker_and_isin_row(row: List[Any]) -> tuple[str, str]:
    """
    Извлекает ticker и ISIN из строки над сделками, например:
    """
    text = " ".join(str(c).strip() for c in row if pd.notna(c))
    match = re.search(r"^(?P<ticker>\S+).*?ISIN[:\s]*(?P<isin>[A-Z0-9]{12})", text)
    if match:
        return match.group("ticker"), match.group("isin")
    return "", ""


def is_ticker_row(row: List[Any]) -> bool:
    return any(
        isinstance(cell, str) and any(keyword in cell.lower() for keyword in ['номер рег', 'isin'])
        for cell in row
    )


def find_block_start(df: pd.DataFrame, keyword: str) -> Optional[int]:
    for i, row in df.iterrows():
        if keyword in ' '.join(str(c).lower() for c in row if pd.notna(c)):
            return i + 1
    return None


def find_header_row(df: pd.DataFrame, required: List[str]) -> Optional[int]:
    for i, row in df.iterrows():
        cells = [str(c).lower() for c in row if pd.notna(c)]
        if all(any(req in cell for cell in cells) for req in required):
            return i
    return None



def parse_stock_section(block: pd.DataFrame) -> List[dict]:
    results: List[dict] = []
    required = ['дата', 'номер', 'куплено', 'продано', 'сумма', 'валюта', 'дата соверш', 'время соверш']
    hdr_idx = find_header_row(block, required)
    if hdr_idx is None:
        return results
    headers = block.iloc[hdr_idx].astype(str).str.lower().tolist()

    idx = {
        'num':     find_column_index(headers, 'номер'),
        'buy_qty': find_column_index(headers, 'куплено'),
        'sell_qty':find_column_index(headers, 'продано'),
        'buy_sum': find_column_index(headers, 'сумма', 'платеж'),
        'sell_sum':find_column_index(headers, 'сумма', 'выруч'),
        'currency':find_column_index(headers, 'валюта'),
        'date':    find_column_index(headers, 'дата соверш'),
        'time':    find_column_index(headers, 'время соверш'),
    }
    price_cols = [i for i, h in enumerate(headers) if 'цена' in h]
    idx['buy_pr'], idx['sell_pr'] = (price_cols + [None, None])[:2]

    curr_ticker, curr_isin = '', ''
    for row in block.iloc[hdr_idx+1:].itertuples(index=False):
        cells = list(row)
        if all(pd.isna(c) or (isinstance(c, str) and not c.strip()) for c in cells):
            break
        if any(isinstance(c, str) and c.lower().startswith('итого') for c in cells):
            continue
        if is_ticker_row(cells):
            curr_ticker, curr_isin = parse_ticker_and_isin_row(cells)
            continue
        dt = pd.to_datetime(
            f"{cells[idx['date']]} {cells[idx['time']]}", dayfirst=True, errors='coerce'
        )
        if pd.isna(dt):
            continue

        buy = to_num(cells[idx['buy_qty']])
        sell = to_num(cells[idx['sell_qty']])
        if buy > 0:
            op, qty, pr, total, aci = 'buy', buy, to_num(cells[idx['buy_pr']]), to_num(cells[idx['buy_sum']]), 0.0
        elif sell > 0:
            op, qty, pr, total, aci = 'sell', sell, to_num(cells[idx['sell_pr']]), to_num(cells[idx['sell_sum']]), 0.0
        else:
            continue

        results.append(
            OperationDTO(
                date=dt.strftime('%Y-%m-%d %H:%M:%S'),
                operation_type=op,
                payment_sum=total,
                currency=str(cells[idx['currency']]).strip(),
                ticker=curr_ticker,
                isin=curr_isin,
                price=pr,
                quantity=int(qty),
                aci=aci,
                comment='',
                operation_id=str(cells[idx['num']]).strip()
            ).to_dict()
        )
    return results


def parse_bond_section(block: pd.DataFrame) -> List[dict]:
    results: List[dict] = []
    required = ['совершена', 'номер', 'куплено', 'продано', 'сумма', 'валюта', 'нкд покупки', 'нкд продажи']
    hdr_idx = find_header_row(block, required)
    if hdr_idx is None:
        return results
    headers = block.iloc[hdr_idx].astype(str).str.lower().tolist()

    idx = {
        'num':     find_column_index(headers, 'номер'),
        'buy_qty': find_column_index(headers, 'куплено'),
        'sell_qty':find_column_index(headers, 'продано'),
        'buy_sum': find_column_index(headers, 'сумма', 'платеж'),
        'sell_sum':find_column_index(headers, 'сумма', 'выруч'),
        'currency':find_column_index(headers, 'валюта'),
        'date':    find_column_index(headers, 'совершена'),
        'aci_buy': find_column_index(headers, 'нкд', 'покупки'),
        'aci_sell':find_column_index(headers, 'нкд', 'продажи'),
    }
    price_cols = [i for i, h in enumerate(headers) if 'цена' in h]
    idx['buy_pr'], idx['sell_pr'] = (price_cols + [None, None])[:2]

    curr_ticker, curr_isin = '', ''
    for row in block.iloc[hdr_idx+1:].itertuples(index=False):
        cells = list(row)
        if all(pd.isna(c) or (isinstance(c, str) and not c.strip()) for c in cells):
            break
        if any(isinstance(c, str) and c.lower().startswith('итого') for c in cells):
            continue
        if is_ticker_row(cells):
            curr_ticker, curr_isin = parse_ticker_and_isin_row(cells)
            continue
        dt = pd.to_datetime(f"{cells[idx['date']]} 00:00:00", dayfirst=True, errors='coerce')
        if pd.isna(dt):
            continue
        buy = to_num(cells[idx['buy_qty']])
        sell = to_num(cells[idx['sell_qty']])
        if buy > 0:
            op, qty, pr, total, aci = 'buy', buy, to_num(cells[idx['buy_pr']]), to_num(cells[idx['buy_sum']]), to_num(cells[idx['aci_buy']])
        elif sell > 0:
            op, qty, pr, total, aci = 'sell', sell, to_num(cells[idx['sell_pr']]), to_num(cells[idx['sell_sum']]), to_num(cells[idx['aci_sell']])
        else:
            continue
        results.append(
            OperationDTO(
                date=dt.strftime('%Y-%m-%d %H:%M:%S'),
                operation_type=op,
                payment_sum=total,
                currency=str(cells[idx['currency']]).strip(),
                ticker=curr_ticker,
                isin=curr_isin,
                price=pr,
                quantity=int(qty),
                aci=aci,
                comment='',
                operation_id=str(cells[idx['num']]).strip()
            ).to_dict()
        )
    return results


def parse_stock_bond_trades(file_path: Union[str, Any]) -> List[dict]:
    df = pd.read_excel(file_path, header=None)
    start_idx = find_block_start(df, '2.1. сделки')
    if start_idx is None:
        return []
    block = df.iloc[start_idx:].reset_index(drop=True)

    section_indices, section_types = [], []
    for i, row in block.iterrows():
        text = ' '.join(str(c).lower() for c in row if pd.notna(c))
        if 'акция' in text or 'адр' in text:
            section_indices.append(i)
            section_types.append('stock')
        elif 'облигация' in text:
            section_indices.append(i)
            section_types.append('bond')
    section_indices.append(len(block))

    all_results: List[dict] = []
    for (s, e), t in zip(zip(section_indices, section_indices[1:]), section_types):
        sect = block.iloc[s+1:e].reset_index(drop=True)
        if t == 'stock':
            all_results.extend(parse_stock_section(sect))
        else:
            all_results.extend(parse_bond_section(sect))

    return sorted(all_results, key=lambda x: x['date'])



