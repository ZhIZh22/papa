import os
import re
import logging
from datetime import date, datetime, timedelta

from dotenv import load_dotenv
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, Document
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ConversationHandler,
    ContextTypes,
    filters,
)

load_dotenv()
import excel_helper  # после load_dotenv, чтобы EXCEL_FILE подхватился

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN")
ALLOWED_USERS = set(int(x) for x in os.getenv("ALLOWED_USERS", "").split(",") if x.strip())
PROXY_URL = os.getenv("PROXY_URL", "")

# Состояния ConversationHandler
WAIT_COMMENT = 1
WAIT_DATE_CHOICE = 2
WAIT_CUSTOM_DATE = 3
WAIT_AMOUNT_FOR_DATE = 4
WAIT_COMMENT_FOR_DATE = 5

RUSSIAN_MONTHS = {
    1: "января", 2: "февраля", 3: "марта", 4: "апреля",
    5: "мая", 6: "июня", 7: "июля", 8: "августа",
    9: "сентября", 10: "октября", 11: "ноября", 12: "декабря",
}


def fmt_date_ru(d: date) -> str:
    return f"{d.day} {RUSSIAN_MONTHS[d.month]} {d.year}"


def fmt_money(n: int) -> str:
    sign = "-" if n < 0 else "+"
    s = str(abs(n))
    result = []
    for i, ch in enumerate(reversed(s)):
        if i > 0 and i % 3 == 0:
            result.append(" ")
        result.append(ch)
    return sign + "".join(reversed(result)) + " ₽"


def parse_amount(text: str):
    """Парсит сумму из текста. Возвращает int или None."""
    text = text.strip().replace(" ", "").replace("\xa0", "")
    m = re.match(r"^([+-]?\d+)(.*)$", text)
    if not m:
        return None, None
    amount = int(m.group(1))
    rest = m.group(2).strip().lstrip(",").strip()
    return amount, rest if rest else None


def totals_text(year: int) -> str:
    t = excel_helper.get_totals(year)
    return (
        f"📊 Итоги {year}:\n"
        f"  Доход:   {fmt_money(t['income'])[1:]}\n"
        f"  Расход:  {fmt_money(t['expense'])[1:]}\n"
        f"  Остаток: {fmt_money(t['balance'])}"
    )


# ── Авторизация ───────────────────────────────────────────────────────────────

def is_allowed(update: Update) -> bool:
    return update.effective_user.id in ALLOWED_USERS


async def deny(update: Update):
    await update.effective_message.reply_text("⛔ Нет доступа.")


# ── /start ────────────────────────────────────────────────────────────────────

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return
    await update.message.reply_text(
        "Привет! 👋\n\n"
        "Отправь сумму с плюсом или минусом, и я запишу её в таблицу:\n"
        "  <code>+5000 зарплата</code> — доход\n"
        "  <code>-1200 продукты</code> — расход\n\n"
        "Комментарий необязателен — если не напишешь, я попрошу.\n\n"
        "Команды:\n"
        "  /add — добавить запись на другую дату\n"
        "  /download — скачать таблицу\n"
        "  /upload — загрузить новую таблицу\n"
        "  /totals — итоги за текущий год",
        parse_mode="HTML",
    )


# ── /totals ───────────────────────────────────────────────────────────────────

async def cmd_totals(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return
    year = date.today().year
    await update.message.reply_text(totals_text(year))


# ── /download ─────────────────────────────────────────────────────────────────

async def cmd_download(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return
    excel_path = os.path.abspath(excel_helper.EXCEL_FILE)
    if not os.path.exists(excel_path):
        await update.message.reply_text("Файл таблицы не найден.")
        return
    with open(excel_path, "rb") as f:
        await update.message.reply_document(
            document=f,
            filename=os.path.basename(excel_path),
            caption="Вот текущая таблица 📎",
        )


# ── /upload ───────────────────────────────────────────────────────────────────

async def cmd_upload(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return
    await update.message.reply_text(
        "Отправь Excel файл (.xlsx) — я заменю им текущую таблицу."
    )
    context.user_data["waiting_upload"] = True


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return
    if not context.user_data.get("waiting_upload"):
        await update.message.reply_text("Не жду файл. Используй /upload чтобы загрузить таблицу.")
        return

    doc: Document = update.message.document
    if not doc.file_name.endswith(".xlsx"):
        await update.message.reply_text("Нужен файл формата .xlsx")
        return

    file = await doc.get_file()
    tmp_path = "tmp_upload.xlsx"
    await file.download_to_drive(tmp_path)
    excel_helper.replace_excel_file(tmp_path)
    os.remove(tmp_path)
    context.user_data["waiting_upload"] = False
    await update.message.reply_text("✅ Таблица обновлена!")


# ── Быстрый ввод (сообщение с суммой) ────────────────────────────────────────

async def handle_amount_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает сообщения вида '+5000 комментарий' или '-1200'."""
    if not is_allowed(update):
        await deny(update)
        return

    text = update.message.text.strip()
    amount, comment = parse_amount(text)

    if amount is None or amount == 0:
        await update.message.reply_text(
            "Не понял. Отправь сумму с + или −, например:\n"
            "<code>+5000 зарплата</code>\n"
            "<code>-800 кафе</code>",
            parse_mode="HTML",
        )
        return

    today = date.today()

    if comment:
        # Всё есть — сразу пишем
        result = excel_helper.add_transaction(today, amount, comment)
        direction = "Доход" if amount > 0 else "Расход"
        await update.message.reply_text(
            f"✅ Записано на {fmt_date_ru(today)}\n"
            f"{direction}: {fmt_money(amount)}\n"
            f"Комментарий: {comment}\n\n"
            + totals_text(today.year)
        )
        return

    # Нет комментария — запоминаем и спрашиваем
    context.user_data["pending_amount"] = amount
    context.user_data["pending_date"] = today

    # Проверяем: был ли уже сегодня другой платёж (тогда кнопка "пропустить")
    day_info = excel_helper.get_day_info(today)
    has_previous = day_info["income"] > 0 or day_info["expense"] > 0

    direction = "доход" if amount > 0 else "расход"
    text_ask = (
        f"Записать {direction} {fmt_money(amount)} на {fmt_date_ru(today)}.\n"
        f"Напиши комментарий:"
    )

    if has_previous:
        keyboard = [[InlineKeyboardButton("Пропустить комментарий", callback_data="skip_comment")]]
        await update.message.reply_text(
            text_ask,
            reply_markup=InlineKeyboardMarkup(keyboard),
        )
    else:
        await update.message.reply_text(text_ask)

    return WAIT_COMMENT


async def receive_comment(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Получает комментарий к отложенной транзакции."""
    if not is_allowed(update):
        await deny(update)
        return ConversationHandler.END

    comment = update.message.text.strip()
    amount = context.user_data.pop("pending_amount", None)
    target_date = context.user_data.pop("pending_date", date.today())

    if amount is None:
        return ConversationHandler.END

    result = excel_helper.add_transaction(target_date, amount, comment)
    direction = "Доход" if amount > 0 else "Расход"
    await update.message.reply_text(
        f"✅ Записано на {fmt_date_ru(target_date)}\n"
        f"{direction}: {fmt_money(amount)}\n"
        f"Комментарий: {comment}\n\n"
        + totals_text(target_date.year)
    )
    return ConversationHandler.END


async def skip_comment_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Нажата кнопка 'Пропустить комментарий'."""
    query = update.callback_query
    await query.answer()

    amount = context.user_data.pop("pending_amount", None)
    target_date = context.user_data.pop("pending_date", date.today())

    if amount is None:
        await query.edit_message_text("Нет ожидающей транзакции.")
        return ConversationHandler.END

    result = excel_helper.add_transaction(target_date, amount, "")
    direction = "Доход" if amount > 0 else "Расход"
    await query.edit_message_text(
        f"✅ Записано на {fmt_date_ru(target_date)}\n"
        f"{direction}: {fmt_money(amount)}\n"
        f"Без комментария.\n\n"
        + totals_text(target_date.year)
    )
    return ConversationHandler.END


# ── /add — выбор даты и ввод ──────────────────────────────────────────────────

def date_picker_keyboard(center_date: date) -> InlineKeyboardMarkup:
    """Клавиатура для выбора даты: ±7 дней от центра."""
    today = date.today()
    buttons = []
    row = []
    for delta in range(-7, 8):
        d = today + timedelta(days=delta)
        label = d.strftime("%d.%m")
        if d == today:
            label = f"[{label}]"
        row.append(InlineKeyboardButton(label, callback_data=f"date_{d.isoformat()}"))
        if len(row) == 4:
            buttons.append(row)
            row = []
    if row:
        buttons.append(row)
    buttons.append([InlineKeyboardButton("Ввести дату вручную", callback_data="date_manual")])
    return InlineKeyboardMarkup(buttons)


async def cmd_add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return
    await update.message.reply_text(
        "Выбери дату для записи:",
        reply_markup=date_picker_keyboard(date.today()),
    )
    return WAIT_DATE_CHOICE


async def date_choice_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    if data == "date_manual":
        await query.edit_message_text("Введи дату в формате ДД.ММ.ГГ:")
        return WAIT_CUSTOM_DATE

    chosen_date = date.fromisoformat(data.replace("date_", ""))
    context.user_data["add_date"] = chosen_date
    await query.edit_message_text(
        f"Дата: {fmt_date_ru(chosen_date)}\n\n"
        "Отправь сумму с + или −, и необязательно комментарий:\n"
        "<code>+5000 аванс</code>\n"
        "<code>-1200</code>",
        parse_mode="HTML",
    )
    return WAIT_AMOUNT_FOR_DATE


async def receive_custom_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    try:
        chosen_date = datetime.strptime(text, "%d.%m.%y").date()
    except ValueError:
        try:
            chosen_date = datetime.strptime(text, "%d.%m.%Y").date()
        except ValueError:
            await update.message.reply_text("Не понял дату. Введи в формате ДД.ММ.ГГ, например: 05.03.26")
            return WAIT_CUSTOM_DATE

    context.user_data["add_date"] = chosen_date
    await update.message.reply_text(
        f"Дата: {fmt_date_ru(chosen_date)}\n\n"
        "Отправь сумму с + или −, и необязательно комментарий:\n"
        "<code>+5000 аванс</code>\n"
        "<code>-1200</code>",
        parse_mode="HTML",
    )
    return WAIT_AMOUNT_FOR_DATE


async def receive_amount_for_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return ConversationHandler.END

    text = update.message.text.strip()
    amount, comment = parse_amount(text)

    if amount is None or amount == 0:
        await update.message.reply_text(
            "Не понял. Отправь сумму с + или −, например:\n"
            "<code>+5000 аванс</code>\n"
            "<code>-800</code>",
            parse_mode="HTML",
        )
        return WAIT_AMOUNT_FOR_DATE

    target_date = context.user_data.get("add_date", date.today())

    if comment:
        result = excel_helper.add_transaction(target_date, amount, comment)
        direction = "Доход" if amount > 0 else "Расход"
        await update.message.reply_text(
            f"✅ Записано на {fmt_date_ru(target_date)}\n"
            f"{direction}: {fmt_money(amount)}\n"
            f"Комментарий: {comment}\n\n"
            + totals_text(target_date.year)
        )
        return ConversationHandler.END

    # Нет комментария
    context.user_data["pending_amount"] = amount
    context.user_data["pending_date"] = target_date

    day_info = excel_helper.get_day_info(target_date)
    has_previous = day_info["income"] > 0 or day_info["expense"] > 0

    direction = "доход" if amount > 0 else "расход"
    text_ask = (
        f"Записать {direction} {fmt_money(amount)} на {fmt_date_ru(target_date)}.\n"
        f"Напиши комментарий:"
    )

    if has_previous:
        keyboard = [[InlineKeyboardButton("Пропустить комментарий", callback_data="skip_comment")]]
        await update.message.reply_text(
            text_ask,
            reply_markup=InlineKeyboardMarkup(keyboard),
        )
    else:
        await update.message.reply_text(text_ask)

    return WAIT_COMMENT_FOR_DATE


async def receive_comment_for_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return ConversationHandler.END

    comment = update.message.text.strip()
    amount = context.user_data.pop("pending_amount", None)
    target_date = context.user_data.pop("pending_date", date.today())

    if amount is None:
        return ConversationHandler.END

    result = excel_helper.add_transaction(target_date, amount, comment)
    direction = "Доход" if amount > 0 else "Расход"
    await update.message.reply_text(
        f"✅ Записано на {fmt_date_ru(target_date)}\n"
        f"{direction}: {fmt_money(amount)}\n"
        f"Комментарий: {comment}\n\n"
        + totals_text(target_date.year)
    )
    return ConversationHandler.END


# ── Запуск ────────────────────────────────────────────────────────────────────

def main():
    builder = Application.builder().token(BOT_TOKEN)
    if PROXY_URL:
        builder = builder.proxy(PROXY_URL).get_updates_proxy(PROXY_URL)
    app = builder.build()

    # ConversationHandler для быстрого ввода (без /add)
    quick_conv = ConversationHandler(
        entry_points=[
            MessageHandler(filters.TEXT & filters.Regex(r"^[+-]\d+") & ~filters.COMMAND, handle_amount_message)
        ],
        states={
            WAIT_COMMENT: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_comment),
                CallbackQueryHandler(skip_comment_callback, pattern="^skip_comment$"),
            ],
        },
        fallbacks=[],
        per_message=False,
    )

    # ConversationHandler для /add
    add_conv = ConversationHandler(
        entry_points=[CommandHandler("add", cmd_add)],
        states={
            WAIT_DATE_CHOICE: [
                CallbackQueryHandler(date_choice_callback, pattern="^date_"),
            ],
            WAIT_CUSTOM_DATE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_custom_date),
            ],
            WAIT_AMOUNT_FOR_DATE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_amount_for_date),
            ],
            WAIT_COMMENT_FOR_DATE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_comment_for_date),
                CallbackQueryHandler(skip_comment_callback, pattern="^skip_comment$"),
            ],
        },
        fallbacks=[],
        per_message=False,
    )

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("totals", cmd_totals))
    app.add_handler(CommandHandler("download", cmd_download))
    app.add_handler(CommandHandler("upload", cmd_upload))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(quick_conv)
    app.add_handler(add_conv)

    logger.info("Бот запущен.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
