import os
import shutil
from datetime import date, datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side
from openpyxl.styles.numbers import FORMAT_DATE_DDMMYY

EXCEL_FILE = os.getenv("EXCEL_FILE", "!!!Мастерская.xlsx")

_THIN = Side(style="thin")
_BORDER = Border(left=_THIN, right=_THIN)
_ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
_ALIGN_LEFT   = Alignment(horizontal="left",   vertical="center")

# Формат даты как в оригинальной таблице: ДД.ММ.ГГ
DATE_FORMAT = "DD.MM.YY"


def _apply_row_style(ws, row_idx: int):
    for col in range(1, 5):
        cell = ws.cell(row=row_idx, column=col)
        cell.border = _BORDER
        cell.alignment = _ALIGN_LEFT if col == 4 else _ALIGN_CENTER


def _get_sheet_name(year: int) -> str:
    return str(year)[-2:]


def _parse_money(s) -> int:
    if s is None:
        return 0
    if isinstance(s, (int, float)):
        return int(s)
    s = str(s).replace(" ", "").replace("\xa0", "").replace(",", "")
    try:
        return int(float(s))
    except ValueError:
        return 0


def _cell_as_date(cell) -> date | None:
    """Читает дату из ячейки — поддерживает datetime, date и строки."""
    val = cell.value
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    # Строка — пробуем распарсить
    s = str(val).strip()
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d"):
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
        ws["A3"] = "остаток"
    return wb[sheet_name]


def _last_used_row(ws) -> int:
    """Последняя строка с данными в любом из столбцов A-D (>= 4), иначе 3."""
    for row_idx in range(ws.max_row, 3, -1):
        for col in range(1, 5):
            v = ws.cell(row=row_idx, column=col).value
            if v is not None and str(v).strip() != "":
                return row_idx
    return 3


def _find_row_by_date(ws, target_date: date) -> int | None:
    """Находит первую строку с нужной датой в столбце A."""
    for row_idx in range(4, ws.max_row + 1):
        d = _cell_as_date(ws.cell(row=row_idx, column=1))
        if d == target_date:
            return row_idx
    return None


def _last_row_of_day(ws, first_row: int) -> int:
    """
    Возвращает последнюю строку принадлежащую дню начатому в first_row.
    Строки без даты в A считаются продолжением предыдущего дня.
    """
    last = first_row
    row_idx = first_row + 1
    while row_idx <= ws.max_row:
        # Если в A есть дата — начался новый день
        if _cell_as_date(ws.cell(row=row_idx, column=1)) is not None:
            break
        # Если есть данные в B/C/D — это строка нашего дня
        has_data = any(
            ws.cell(row=row_idx, column=c).value not in (None, "")
            for c in range(2, 5)
        )
        if has_data:
            last = row_idx
        row_idx += 1
    return last


def _write_date_cell(cell, d: date):
    """Записывает дату в ячейку как datetime с форматированием ДД.ММ.ГГ."""
    cell.value = datetime(d.year, d.month, d.day)
    cell.number_format = DATE_FORMAT
    cell.alignment = _ALIGN_CENTER
    cell.border = _BORDER


def add_transaction(target_date: date, amount: int, comment: str = "") -> dict:
    """
    Каждая транзакция — отдельная строка.
    Первая за день: дата в A, потом сумма и комментарий.
    Следующие за тот же день: без даты, строка добавляется после последней строки дня.
    Прошлое число: вставляется строка через insert_rows рядом с нужной датой.
    """
    wb = load_workbook(EXCEL_FILE)
    sheet_name = _get_sheet_name(target_date.year)
    ws = _ensure_sheet(wb, sheet_name)

    first_row = _find_row_by_date(ws, target_date)

    if first_row is None:
        # Даты нет в таблице
        # Находим все даты чтобы понять куда вставить
        dated_rows = {}  # row_idx -> date
        for r in range(4, ws.max_row + 1):
            d = _cell_as_date(ws.cell(row=r, column=1))
            if d is not None:
                dated_rows[r] = d

        if not dated_rows or target_date >= max(dated_rows.values()):
            # Нет данных вообще или дата >= последней — добавляем в конец
            row_idx = _last_used_row(ws) + 1
            _apply_row_style(ws, row_idx)
            _write_date_cell(ws.cell(row=row_idx, column=1), target_date)
        else:
            # Дата в прошлом — вставляем insert_rows после нужного места
            insert_after = 3
            for r in sorted(dated_rows.keys()):
                if dated_rows[r] < target_date:
                    # Учитываем все строки этого дня
                    insert_after = _last_row_of_day(ws, r)

            insert_at = insert_after + 1
            ws.insert_rows(insert_at)
            _apply_row_style(ws, insert_at)
            _write_date_cell(ws.cell(row=insert_at, column=1), target_date)
            row_idx = insert_at
    else:
        # Дата уже есть — добавляем строку после последней строки этого дня
        last_day_row = _last_row_of_day(ws, first_row)

        # Определяем: этот день последний в таблице?
        dated_rows = {}
        for r in range(4, ws.max_row + 1):
            d = _cell_as_date(ws.cell(row=r, column=1))
            if d is not None:
                dated_rows[r] = d

        last_date = max(dated_rows.values()) if dated_rows else target_date

        if target_date >= last_date:
            # Последний день — просто добавляем в конец
            row_idx = _last_used_row(ws) + 1
            _apply_row_style(ws, row_idx)
        else:
            # День в середине — insert_rows
            insert_at = last_day_row + 1
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
        total_income  += _parse_money(row[1])
        total_expense += _parse_money(row[2])
    return {
        "income": total_income,
        "expense": total_expense,
        "balance": total_income - total_expense,
    }


def get_day_info(target_date: date) -> dict:
    wb = load_workbook(EXCEL_FILE)
    sheet_name = _get_sheet_name(target_date.year)
    if sheet_name not in wb.sheetnames:
        return {"date": target_date, "has_records": False}
    ws = wb[sheet_name]
    has = _find_row_by_date(ws, target_date) is not None
    return {"date": target_date, "has_records": has}


def replace_excel_file(new_file_path: str):
    shutil.copy2(new_file_path, EXCEL_FILE)
