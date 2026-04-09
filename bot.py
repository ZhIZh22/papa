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
    ContextTypes,
    filters,
)

load_dotenv()
import excel_helper

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN")
ALLOWED_USERS = set(int(x) for x in os.getenv("ALLOWED_USERS", "").split(",") if x.strip())
PROXY_URL = os.getenv("PROXY_URL", "")

RUSSIAN_MONTHS = {
    1: "января", 2: "февраля", 3: "марта", 4: "апреля",
    5: "мая", 6: "июня", 7: "июля", 8: "августа",
    9: "сентября", 10: "октября", 11: "ноября", 12: "декабря",
}

# Режимы ожидания в user_data["mode"]
MODE_WAIT_COMMENT = "wait_comment"
MODE_WAIT_AMOUNT_FOR_DATE = "wait_amount_for_date"
MODE_WAIT_COMMENT_FOR_DATE = "wait_comment_for_date"
MODE_WAIT_CUSTOM_DATE = "wait_custom_date"
MODE_WAIT_UPLOAD = "wait_upload"


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
    text = text.strip().replace("\xa0", "")
    # Убираем пробелы только внутри числа, но сохраняем текст после
    m = re.match(r"^([+-]?\s*[\d\s]+?)([^\d\s].*)$", text)
    if m:
        num_part = m.group(1).replace(" ", "")
        rest = m.group(2).strip()
    else:
        # Только число
        num_part = text.replace(" ", "")
        rest = ""
    try:
        amount = int(num_part)
    except ValueError:
        return None, None
    return amount, rest if rest else None


def totals_text(year: int) -> str:
    t = excel_helper.get_totals(year)
    return (
        f"📊 Итоги {year}:\n"
        f"  Доход:   {fmt_money(t['income'])[1:]} ₽\n"
        f"  Расход:  {fmt_money(t['expense'])[1:]} ₽\n"
        f"  Остаток: {fmt_money(t['balance'])} ₽"
    )


def is_allowed(update: Update) -> bool:
    return update.effective_user.id in ALLOWED_USERS


async def deny(update: Update):
    await update.effective_message.reply_text("⛔ Нет доступа.")


def clear_mode(context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("mode", None)
    context.user_data.pop("pending_amount", None)
    context.user_data.pop("pending_date", None)
    context.user_data.pop("add_date", None)
    context.user_data.pop("waiting_upload", None)


# ── /start ────────────────────────────────────────────────────────────────────

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return
    clear_mode(context)
    await update.message.reply_text(
        "Привет! 👋\n\n"
        "Отправь сумму с плюсом или минусом:\n"
        "  <code>+5000 зарплата</code> — доход\n"
        "  <code>-1200 продукты</code> — расход\n\n"
        "Комментарий необязателен.\n\n"
        "Команды:\n"
        "  /add — запись на другую дату\n"
        "  /download — скачать таблицу\n"
        "  /upload — загрузить таблицу\n"
        "  /totals — итоги за год",
        parse_mode="HTML",
    )


# ── /totals ───────────────────────────────────────────────────────────────────

async def cmd_totals(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return
    await update.message.reply_text(totals_text(date.today().year))


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
    clear_mode(context)
    context.user_data["mode"] = MODE_WAIT_UPLOAD
    await update.message.reply_text("Отправь Excel файл (.xlsx) — заменю текущую таблицу.")


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return
    if context.user_data.get("mode") != MODE_WAIT_UPLOAD:
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
    clear_mode(context)
    await update.message.reply_text("✅ Таблица обновлена!")


# ── /add — выбор даты ─────────────────────────────────────────────────────────

def date_picker_keyboard() -> InlineKeyboardMarkup:
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
    clear_mode(context)
    await update.message.reply_text(
        "Выбери дату для записи:",
        reply_markup=date_picker_keyboard(),
    )


async def callback_date_pick(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if not is_allowed(update):
        await query.edit_message_text("⛔ Нет доступа.")
        return

    data = query.data

    if data == "date_manual":
        clear_mode(context)
        context.user_data["mode"] = MODE_WAIT_CUSTOM_DATE
        await query.edit_message_text("Введи дату в формате ДД.ММ.ГГ (например: 05.03.26):")
        return

    chosen_date = date.fromisoformat(data.replace("date_", ""))
    clear_mode(context)
    context.user_data["mode"] = MODE_WAIT_AMOUNT_FOR_DATE
    context.user_data["add_date"] = chosen_date
    await query.edit_message_text(
        f"Дата: {fmt_date_ru(chosen_date)}\n\n"
        "Отправь сумму с + или −:\n"
        "<code>+5000 аванс</code>\n"
        "<code>-1200</code>",
        parse_mode="HTML",
    )


async def callback_skip_comment(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if not is_allowed(update):
        await query.edit_message_text("⛔ Нет доступа.")
        return

    amount = context.user_data.pop("pending_amount", None)
    target_date = context.user_data.pop("pending_date", date.today())
    clear_mode(context)

    if amount is None:
        await query.edit_message_text("Нет ожидающей транзакции.")
        return

    excel_helper.add_transaction(target_date, amount, "")
    direction = "Доход" if amount > 0 else "Расход"
    await query.edit_message_text(
        f"✅ Записано на {fmt_date_ru(target_date)}\n"
        f"{direction}: {fmt_money(amount)}\n"
        f"Без комментария.\n\n"
        + totals_text(target_date.year)
    )


# ── Основной обработчик текстовых сообщений ──────────────────────────────────

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return

    text = update.message.text.strip()
    mode = context.user_data.get("mode")

    # Ждём дату вручную
    if mode == MODE_WAIT_CUSTOM_DATE:
        try:
            chosen_date = datetime.strptime(text, "%d.%m.%y").date()
        except ValueError:
            try:
                chosen_date = datetime.strptime(text, "%d.%m.%Y").date()
            except ValueError:
                await update.message.reply_text("Не понял дату. Введи в формате ДД.ММ.ГГ, например: 05.03.26")
                return
        context.user_data["mode"] = MODE_WAIT_AMOUNT_FOR_DATE
        context.user_data["add_date"] = chosen_date
        await update.message.reply_text(
            f"Дата: {fmt_date_ru(chosen_date)}\n\n"
            "Отправь сумму с + или −:\n"
            "<code>+5000 аванс</code>\n"
            "<code>-1200</code>",
            parse_mode="HTML",
        )
        return

    # Ждём сумму для выбранной даты
    if mode == MODE_WAIT_AMOUNT_FOR_DATE:
        amount, comment = parse_amount(text)
        if amount is None or amount == 0:
            await update.message.reply_text(
                "Не понял. Отправь сумму с + или −, например:\n"
                "<code>+5000</code> или <code>-800 кафе</code>",
                parse_mode="HTML",
            )
            return
        target_date = context.user_data.get("add_date", date.today())
        if comment:
            excel_helper.add_transaction(target_date, amount, comment)
            clear_mode(context)
            direction = "Доход" if amount > 0 else "Расход"
            await update.message.reply_text(
                f"✅ Записано на {fmt_date_ru(target_date)}\n"
                f"{direction}: {fmt_money(amount)}\n"
                f"Комментарий: {comment}\n\n"
                + totals_text(target_date.year)
            )
        else:
            context.user_data["mode"] = MODE_WAIT_COMMENT_FOR_DATE
            context.user_data["pending_amount"] = amount
            context.user_data["pending_date"] = target_date
            day_info = excel_helper.get_day_info(target_date)
            has_previous = day_info["income"] > 0 or day_info["expense"] > 0
            direction = "доход" if amount > 0 else "расход"
            text_ask = f"Записать {direction} {fmt_money(amount)} на {fmt_date_ru(target_date)}.\nНапиши комментарий:"
            if has_previous:
                kb = [[InlineKeyboardButton("Пропустить комментарий", callback_data="skip_comment")]]
                await update.message.reply_text(text_ask, reply_markup=InlineKeyboardMarkup(kb))
            else:
                await update.message.reply_text(text_ask)
        return

    # Ждём комментарий для выбранной даты
    if mode == MODE_WAIT_COMMENT_FOR_DATE:
        comment = text
        amount = context.user_data.pop("pending_amount", None)
        target_date = context.user_data.pop("pending_date", date.today())
        clear_mode(context)
        if amount is not None:
            excel_helper.add_transaction(target_date, amount, comment)
            direction = "Доход" if amount > 0 else "Расход"
            await update.message.reply_text(
                f"✅ Записано на {fmt_date_ru(target_date)}\n"
                f"{direction}: {fmt_money(amount)}\n"
                f"Комментарий: {comment}\n\n"
                + totals_text(target_date.year)
            )
        return

    # Ждём комментарий для быстрого ввода (сегодня)
    if mode == MODE_WAIT_COMMENT:
        comment = text
        amount = context.user_data.pop("pending_amount", None)
        target_date = context.user_data.pop("pending_date", date.today())
        clear_mode(context)
        if amount is not None:
            excel_helper.add_transaction(target_date, amount, comment)
            direction = "Доход" if amount > 0 else "Расход"
            await update.message.reply_text(
                f"✅ Записано на {fmt_date_ru(target_date)}\n"
                f"{direction}: {fmt_money(amount)}\n"
                f"Комментарий: {comment}\n\n"
                + totals_text(target_date.year)
            )
        return

    # Быстрый ввод суммы (сегодня)
    if re.match(r"^[+-]", text):
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
            excel_helper.add_transaction(today, amount, comment)
            clear_mode(context)
            direction = "Доход" if amount > 0 else "Расход"
            await update.message.reply_text(
                f"✅ Записано на {fmt_date_ru(today)}\n"
                f"{direction}: {fmt_money(amount)}\n"
                f"Комментарий: {comment}\n\n"
                + totals_text(today.year)
            )
        else:
            context.user_data["mode"] = MODE_WAIT_COMMENT
            context.user_data["pending_amount"] = amount
            context.user_data["pending_date"] = today
            day_info = excel_helper.get_day_info(today)
            has_previous = day_info["income"] > 0 or day_info["expense"] > 0
            direction = "доход" if amount > 0 else "расход"
            text_ask = f"Записать {direction} {fmt_money(amount)} на {fmt_date_ru(today)}.\nНапиши комментарий:"
            if has_previous:
                kb = [[InlineKeyboardButton("Пропустить комментарий", callback_data="skip_comment")]]
                await update.message.reply_text(text_ask, reply_markup=InlineKeyboardMarkup(kb))
            else:
                await update.message.reply_text(text_ask)
        return

    await update.message.reply_text(
        "Не понял. Отправь сумму с + или −, например:\n"
        "<code>+5000 зарплата</code>\n"
        "<code>-800 кафе</code>\n\n"
        "Или используй /add для записи на другую дату.",
        parse_mode="HTML",
    )


# ── Запуск ────────────────────────────────────────────────────────────────────

def main():
    builder = Application.builder().token(BOT_TOKEN)
    if PROXY_URL:
        builder = builder.proxy(PROXY_URL).get_updates_proxy(PROXY_URL)
    app = builder.build()

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("totals", cmd_totals))
    app.add_handler(CommandHandler("download", cmd_download))
    app.add_handler(CommandHandler("upload", cmd_upload))
    app.add_handler(CommandHandler("add", cmd_add))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(CallbackQueryHandler(callback_date_pick, pattern="^date_"))
    app.add_handler(CallbackQueryHandler(callback_skip_comment, pattern="^skip_comment$"))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    logger.info("Бот запущен.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
