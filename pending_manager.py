"""
Менеджер будущих транзакций.
Хранит в pending.json список записей вида:
  {"date": "YYYY-MM-DD", "amount": int, "comment": str}
"""
import json
import os
from datetime import date

PENDING_FILE = "pending.json"


def _load() -> list:
    if not os.path.exists(PENDING_FILE):
        return []
    with open(PENDING_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def _save(data: list):
    with open(PENDING_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def add(target_date: date, amount: int, comment: str):
    data = _load()
    data.append({
        "date": target_date.isoformat(),
        "amount": amount,
        "comment": comment,
    })
    _save(data)


def get_all() -> list:
    """Возвращает все pending-записи как список dict с полем date: date."""
    data = _load()
    result = []
    for item in data:
        result.append({
            "date": date.fromisoformat(item["date"]),
            "amount": item["amount"],
            "comment": item["comment"],
        })
    return result


def get_due(today: date) -> list:
    """Возвращает записи, дата которых <= сегодня."""
    return [item for item in get_all() if item["date"] <= today]


def remove_by_index(idx: int) -> bool:
    """Удаляет запись по индексу (0-based). Возвращает True если удалено."""
    data = _load()
    if 0 <= idx < len(data):
        data.pop(idx)
        _save(data)
        return True
    return False


def remove_due(today: date):
    """Удаляет все записи с датой <= сегодня."""
    data = _load()
    data = [item for item in data if date.fromisoformat(item["date"]) > today]
    _save(data)
