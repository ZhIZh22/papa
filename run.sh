#!/bin/bash
# Запуск бота
cd "$(dirname "$0")"

# Установка зависимостей (первый раз)
if [ ! -d "venv" ]; then
    echo "Создаём виртуальное окружение..."
    python3 -m venv venv
fi

source venv/bin/activate
pip install -q -r requirements.txt

echo "Запускаем бота..."
python bot.py
