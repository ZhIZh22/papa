import os
import re
from datetime import date, datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

EXCEL_FILE = os.getenv("EXCEL_FILE", "!!!Мастерская.xlsx")


def _get_sheet_name(year: int) -> str:
    """Возвращает имя листа для года (например, 2026 -> '26')."""
    return str(year)[-2:]


def _format_money(value) -> str:
    """Форматирует число в строку с пробелами между разрядами."""
    if value is None:
        return ""
    try:
        n = int(value)
    except (ValueError, TypeError):
        return str(value)
    # Форматируем с пробелами: 1 234 567
    result = []
    s = str(abs(n))
    for i, ch in enumerate(reversed(s)):
        if i > 0 and i % 3 == 0:
            result.append(" ")
        result.append(ch)
    formatted = "".join(reversed(result))
    return ("-" + formatted) if n < 0 else formatted


def _parse_money(s: str):
    """Парсит строку с деньгами в int."""
    s = str(s).replace(" ", "").replace("\xa0", "")
    try:
        return int(s)
    except ValueError:
        return None


def _ensure_sheet(wb, sheet_name: str):
    """Создаёт лист года если его нет, с нужными заголовками и формулами."""
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(title=sheet_name)
        # Строка 1 — заголовки столбцов
        ws["A1"] = "Дата"
        ws["B1"] = "Приход"
        ws["C1"] = "Расход"
        ws["D1"] = "Комментарий"
        # Строка 2 — суммы
        ws["A2"] = "Итого:"
        ws["B2"] = 0  # сумма доходов — будем обновлять вручную
        ws["C2"] = 0  # сумма расходов
        # Строка 3 — остаток
        ws["A3"] = "Остаток:"
        ws["B3"] = 0  # остаток
    return wb[sheet_name]


def _recalc_totals(ws):
    """Пересчитывает суммы в строках 2 и 3."""
    total_income = 0
    total_expense = 0
    for row in ws.iter_rows(min_row=4, values_only=True):
        income = _parse_money(row[1]) if row[1] else 0
        expense = _parse_money(row[2]) if row[2] else 0
        if income:
            total_income += income
        if expense:
            total_expense += expense
    ws["B2"] = _format_money(total_income)
    ws["C2"] = _format_money(total_expense)
    ws["B3"] = _format_money(total_income - total_expense)


def _find_or_create_date_row(ws, target_date: date) -> int:
    """
    Находит строку с нужной датой или создаёт новую строку ниже последней.
    Возвращает номер строки.
    """
    date_str = target_date.strftime("%d.%m.%y")

    # Ищем существующую строку с этой датой
    for row_idx in range(4, ws.max_row + 1):
        cell_val = ws.cell(row=row_idx, column=1).value
        if cell_val and str(cell_val).strip() == date_str:
            return row_idx

    # Находим последнюю заполненную строку
    last_row = 3
    for row_idx in range(4, ws.max_row + 1):
        if ws.cell(row=row_idx, column=1).value:
            last_row = row_idx

    new_row = last_row + 1
    ws.cell(row=new_row, column=1).value = date_str
    return new_row


def add_transaction(target_date: date, amount: int, comment: str = "") -> dict:
    """
    Добавляет транзакцию в Excel.
    amount > 0 — доход, amount < 0 — расход.
    Возвращает словарь с итогами листа.
    """
    wb = load_workbook(EXCEL_FILE)
    sheet_name = _get_sheet_name(target_date.year)
    ws = _ensure_sheet(wb, sheet_name)

    row_idx = _find_or_create_date_row(ws, target_date)

    if amount > 0:
        # Доход — столбец B
        existing = _parse_money(ws.cell(row=row_idx, column=2).value) or 0
        ws.cell(row=row_idx, column=2).value = _format_money(existing + amount)
    else:
        # Расход — столбец C (сохраняем как положительное число)
        existing = _parse_money(ws.cell(row=row_idx, column=3).value) or 0
        ws.cell(row=row_idx, column=3).value = _format_money(existing + abs(amount))

    # Комментарий
    if comment:
        existing_comment = ws.cell(row=row_idx, column=4).value or ""
        if existing_comment:
            ws.cell(row=row_idx, column=4).value = existing_comment + ", " + comment
        else:
            ws.cell(row=row_idx, column=4).value = comment

    _recalc_totals(ws)
    wb.save(EXCEL_FILE)

    return get_totals(target_date.year)


def add_comment_to_date(target_date: date, comment: str) -> bool:
    """Добавляет комментарий к существующей строке даты."""
    wb = load_workbook(EXCEL_FILE)
    sheet_name = _get_sheet_name(target_date.year)
    if sheet_name not in wb.sheetnames:
        return False
    ws = wb[sheet_name]

    date_str = target_date.strftime("%d.%m.%y")
    for row_idx in range(4, ws.max_row + 1):
        cell_val = ws.cell(row=row_idx, column=1).value
        if cell_val and str(cell_val).strip() == date_str:
            existing_comment = ws.cell(row=row_idx, column=4).value or ""
            if existing_comment:
                ws.cell(row=row_idx, column=4).value = existing_comment + ", " + comment
            else:
                ws.cell(row=row_idx, column=4).value = comment
            wb.save(EXCEL_FILE)
            return True
    return False


def get_totals(year: int) -> dict:
    """Возвращает итоги за год: доход, расход, остаток."""
    wb = load_workbook(EXCEL_FILE)
    sheet_name = _get_sheet_name(year)
    if sheet_name not in wb.sheetnames:
        return {"income": 0, "expense": 0, "balance": 0}
    ws = wb[sheet_name]
    income = _parse_money(ws["B2"].value) or 0
    expense = _parse_money(ws["C2"].value) or 0
    balance = _parse_money(ws["B3"].value) or (income - expense)
    return {"income": income, "expense": expense, "balance": balance}


def get_day_info(target_date: date) -> dict:
    """Возвращает данные за конкретный день."""
    wb = load_workbook(EXCEL_FILE)
    sheet_name = _get_sheet_name(target_date.year)
    if sheet_name not in wb.sheetnames:
        return {"date": target_date, "income": 0, "expense": 0, "comment": ""}
    ws = wb[sheet_name]
    date_str = target_date.strftime("%d.%m.%y")
    for row_idx in range(4, ws.max_row + 1):
        cell_val = ws.cell(row=row_idx, column=1).value
        if cell_val and str(cell_val).strip() == date_str:
            income = _parse_money(ws.cell(row=row_idx, column=2).value) or 0
            expense = _parse_money(ws.cell(row=row_idx, column=3).value) or 0
            comment = ws.cell(row=row_idx, column=4).value or ""
            return {"date": target_date, "income": income, "expense": expense, "comment": comment}
    return {"date": target_date, "income": 0, "expense": 0, "comment": ""}


def replace_excel_file(new_file_path: str):
    """Заменяет рабочий Excel файл новым."""
    import shutil
    shutil.copy2(new_file_path, EXCEL_FILE)
