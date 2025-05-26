# full_statement.py

import json
from typing import List, Dict, Any

from constants import CURRENCY_DICT
from fin_operations import parse_financial_operations
from forex_trades    import parse_forex_trades
from stocks_bounds   import parse_stock_bond_trades

def normalize_currency(op: Dict[str, Any]) -> None:
    """
    Заменяет op['currency'] на нормализованное значение из CURRENCY_DICT,
    если оно там есть.
    """
    cur = op.get("currency", "")
    op["currency"] = CURRENCY_DICT.get(cur, cur)

def parse_full_statement(file_path: str) -> Dict[str, Any]:
    """
    Собирает:
      1) Финансовые операции по счёту
      2) Сделки с иностранной валютой
      3) Сделки с акциями и облигациями
    Приводит currency через CURRENCY_DICT и возвращает единый словарь
    с метаданными и отсортированным по дате списком операций.
    """

    # 1) Финансовые операции по счёту
    fin = parse_financial_operations(file_path)
    header_data = {
        "account_id":         fin.get("account_id"),
        "account_date_start": fin.get("account_date_start"),
        "date_start":         fin.get("date_start"),
        "date_end":           fin.get("date_end"),
    }
    fin_ops = fin.get("operations", [])

    # 2) Сделки по иностранной валюте
    forex_ops = parse_forex_trades(file_path)

    # 3) Сделки с акциями и облигациями
    stockbond_ops = parse_stock_bond_trades(file_path)

    # 4) Объединяем все операции
    all_ops: List[Dict[str, Any]] = fin_ops + forex_ops + stockbond_ops

    # 5) Нормализуем currency для каждой операции
    for op in all_ops:
        normalize_currency(op)

    # 6) Сортируем по дате
    all_ops.sort(key=lambda op: op.get("date", ""))

    return {
        **header_data,
        "operations": all_ops
    }

if __name__ == "__main__":
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else "4.xls"
    result = parse_full_statement(path)
    print(json.dumps(result, ensure_ascii=False, indent=2))
