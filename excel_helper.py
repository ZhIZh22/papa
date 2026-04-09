import os
import shutil
from datetime import date, datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side

EXCEL_FILE = os.getenv("EXCEL_FILE", "!!!Мастерская.xlsx")

_THIN = Side(style="thin")
_BORDER = Border(left=_THIN, right=_THIN)
_ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
_ALIGN_LEFT   = Alignment(horizontal="left",   vertical="center")


def _apply_row_style(ws, row_idx: int):
    for col in range(1, 5):
        cell = ws.cell(row=row_idx, column=col)
        cell.border = _BORDER
        cell.alignment = _ALIGN_LEFT if col == 4 else _ALIGN_CENTER


def _get_sheet_name(year: int) -> str:
    return str(year)[-2:]


def _parse_money(s):
    if s is None:
        return None
    s = str(s).replace(" ", "").replace("\xa0", "")
    try:
        return int(s)
    except ValueError:
        return None


def _parse_date(s: str) -> date | None:
    for fmt in ("%d.%m.%y", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _ensure_sheet(wb, sheet_name: str):
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(title=sheet_name)
        ws["A1"] = "Дата"
        ws["B1"] = "Приход"
        ws["C1"] = "Расход"
        ws["D1"] = "Комментарий"
        ws["A2"] = "Итого:"
        ws["A3"] = "Остаток:"
    return wb[sheet_name]


def _last_used_row(ws) -> int:
    """Последняя строка где есть хоть что-то в любом из 4 столбцов (>= 4), или 3."""
    last = 3
    for row_idx in range(4, ws.max_row + 1):
        for col in range(1, 5):
            if ws.cell(row=row_idx, column=col).value not in (None, ""):
                last = row_idx
                break
    return last


def _rows_with_dates(ws) -> dict:
    """Возвращает {row_idx: date} для всех строк с датой в столбце A."""
    result = {}
    for row_idx in range(4, ws.max_row + 1):
        val = ws.cell(row=row_idx, column=1).value
        if val:
            d = _parse_date(str(val).strip())
            if d:
                result[row_idx] = d
    return result


def _find_row_by_date(ws, target_date: date) -> int | None:
    date_str = target_date.strftime("%d.%m.%y")
    for row_idx in range(4, ws.max_row + 1):
        val = ws.cell(row=row_idx, column=1).value
        if val and str(val).strip() == date_str:
            return row_idx
    return None


def _insert_new_row_for_date(ws, target_date: date) -> int:
    """
    Вставляет новую пустую строку в правильное место по дате.
    Возвращает номер новой строки.
    """
    dated = _rows_with_dates(ws)

    if not dated:
        row = 4
        ws.cell(row=row, column=1).value = target_date.strftime("%d.%m.%y")
        _apply_row_style(ws, row)
        return row

    last_date = max(dated.values())

    if target_date > last_date:
        # Добавляем в конец после последней использованной строки
        row = _last_used_row(ws) + 1
        ws.cell(row=row, column=1).value = target_date.strftime("%d.%m.%y")
        _apply_row_style(ws, row)
        return row

    # Вставляем в нужное место: после последней строки с датой < target_date
    insert_after = 3
    for row_idx in sorted(dated.keys()):
        if dated[row_idx] < target_date:
            insert_after = row_idx
        elif dated[row_idx] == target_date:
            # Дата уже есть — найдём последнюю строку этого дня (без даты в A)
            # и вставим после неё
            insert_after = row_idx
            # Идём вниз пока строки без даты относятся к этому дню
            next_row = row_idx + 1
            while next_row <= ws.max_row:
                next_val = ws.cell(row=next_row, column=1).value
                if next_val:  # следующая дата
                    break
                # Есть ли данные в строке?
                has_data = any(
                    ws.cell(row=next_row, column=c).value not in (None, "")
                    for c in range(2, 5)
                )
                if has_data:
                    insert_after = next_row
                next_row += 1
            break

    insert_at = insert_after + 1
    ws.insert_rows(insert_at)
    ws.cell(row=insert_at, column=1).value = target_date.strftime("%d.%m.%y")
    _apply_row_style(ws, insert_at)
    return insert_at


def add_transaction(target_date: date, amount: int, comment: str = "") -> dict:
    """
    Каждая транзакция — отдельная строка.
    Первая транзакция за день: дата в столбце A.
    Последующие в тот же день: дата не пишется, строка добавляется после последней строки этого дня.
    Запись на прошлое число: вставляется новая строка рядом с нужной датой.
    """
    wb = load_workbook(EXCEL_FILE)
    sheet_name = _get_sheet_name(target_date.year)
    ws = _ensure_sheet(wb, sheet_name)

    first_row = _find_row_by_date(ws, target_date)

    if first_row is None:
        # Даты нет вообще — вставляем новую строку с датой
        row_idx = _insert_new_row_for_date(ws, target_date)
    else:
        # Дата есть — ищем последнюю строку этого дня
        last_row_of_day = first_row
        next_row = first_row + 1
        while next_row <= ws.max_row:
            next_val = ws.cell(row=next_row, column=1).value
            if next_val:  # начался следующий день
                break
            has_data = any(
                ws.cell(row=next_row, column=c).value not in (None, "")
                for c in range(2, 5)
            )
            if has_data:
                last_row_of_day = next_row
            next_row += 1

        # Вставляем новую строку после последней строки этого дня (без даты)
        insert_at = last_row_of_day + 1
        ws.insert_rows(insert_at)
        _apply_row_style(ws, insert_at)
        row_idx = insert_at

    # Записываем сумму
    if amount > 0:
        ws.cell(row=row_idx, column=2).value = amount
    else:
        ws.cell(row=row_idx, column=3).value = abs(amount)

    # Записываем комментарий
    if comment:
        ws.cell(row=row_idx, column=4).value = comment

    wb.save(EXCEL_FILE)
    return get_totals(target_date.year)


def get_totals(year: int) -> dict:
    wb = load_workbook(EXCEL_FILE)
    sheet_name = _get_sheet_name(year)
    if sheet_name not in wb.sheetnames:
        return {"income": 0, "expense": 0, "balance": 0}
    ws = wb[sheet_name]
    total_income = 0
    total_expense = 0
    for row in ws.iter_rows(min_row=4, values_only=True):
        income  = _parse_money(row[1]) if row[1] else 0
        expense = _parse_money(row[2]) if row[2] else 0
        if income:
            total_income  += income
        if expense:
            total_expense += expense
    return {"income": total_income, "expense": total_expense, "balance": total_income - total_expense}


def get_day_info(target_date: date) -> dict:
    """Есть ли уже записи за этот день."""
    wb = load_workbook(EXCEL_FILE)
    sheet_name = _get_sheet_name(target_date.year)
    if sheet_name not in wb.sheetnames:
        return {"date": target_date, "has_records": False}
    ws = wb[sheet_name]
    row_idx = _find_row_by_date(ws, target_date)
    return {"date": target_date, "has_records": row_idx is not None}


def replace_excel_file(new_file_path: str):
    shutil.copy2(new_file_path, EXCEL_FILE)
