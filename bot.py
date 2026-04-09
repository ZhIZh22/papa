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
import pending_manager

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

MODE_WAIT_COMMENT        = "wait_comment"
MODE_WAIT_AMOUNT_FOR_DATE = "wait_amount_for_date"
MODE_WAIT_COMMENT_FOR_DATE = "wait_comment_for_date"
MODE_WAIT_CUSTOM_DATE    = "wait_custom_date"
MODE_WAIT_UPLOAD         = "wait_upload"


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
    m = re.match(r"^([+-]?\s*[\d\s]+?)([^\d\s].*)$", text)
    if m:
        num_part = m.group(1).replace(" ", "")
        rest = m.group(2).strip()
    else:
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
        f"  Доход:   {fmt_money(t['income'])[1:]}\n"
        f"  Расход:  {fmt_money(t['expense'])[1:]}\n"
        f"  Остаток: {fmt_money(t['balance'])}"
    )


def is_allowed(update: Update) -> bool:
    return update.effective_user.id in ALLOWED_USERS


async def deny(update: Update):
    await update.effective_message.reply_text("⛔ Нет доступа.")


def clear_mode(context: ContextTypes.DEFAULT_TYPE):
    for key in ("mode", "pending_amount", "pending_date", "add_date"):
        context.user_data.pop(key, None)


# ── Flush pending при старте ──────────────────────────────────────────────────

async def flush_pending(app: Application):
    """Записывает в Excel все pending-транзакции с датой <= сегодня."""
    today = date.today()
    due = pending_manager.get_due(today)
    if not due:
        return

    # Сортируем по дате чтобы вставлять по порядку
    due_sorted = sorted(due, key=lambda x: x["date"])
    for item in due_sorted:
        excel_helper.add_transaction(item["date"], item["amount"], item["comment"])
        logger.info(f"Flush pending: {item['date']} {item['amount']} {item['comment']}")

    pending_manager.remove_due(today)

    # Уведомляем всех разрешённых пользователей
    lines = [f"📥 Записаны отложенные транзакции ({len(due_sorted)} шт.):"]
    for item in due_sorted:
        direction = "Доход" if item["amount"] > 0 else "Расход"
        lines.append(f"  {fmt_date_ru(item['date'])} | {direction}: {fmt_money(item['amount'])}"
                     + (f" | {item['comment']}" if item["comment"] else ""))
    text = "\n".join(lines)

    for user_id in ALLOWED_USERS:
        try:
            await app.bot.send_message(chat_id=user_id, text=text)
        except Exception as e:
            logger.warning(f"Не удалось отправить уведомление {user_id}: {e}")


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
        "  /pending — будущие транзакции\n"
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


# ── /pending ──────────────────────────────────────────────────────────────────

async def cmd_pending(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return
    clear_mode(context)
    items = pending_manager.get_all()
    if not items:
        await update.message.reply_text("Нет запланированных транзакций.")
        return

    items_sorted = sorted(items, key=lambda x: x["date"])
    lines = ["📋 Запланированные транзакции:\n"]
    for i, item in enumerate(items_sorted):
        direction = "Доход" if item["amount"] > 0 else "Расход"
        comment_part = f" | {item['comment']}" if item["comment"] else ""
        lines.append(f"{i+1}. {fmt_date_ru(item['date'])} | {direction}: {fmt_money(item['amount'])}{comment_part}")

    keyboard = []
    for i in range(len(items_sorted)):
        keyboard.append([InlineKeyboardButton(f"Удалить #{i+1}", callback_data=f"del_pending_{i}")])

    await update.message.reply_text(
        "\n".join(lines),
        reply_markup=InlineKeyboardMarkup(keyboard),
    )


async def callback_del_pending(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if not is_allowed(update):
        await query.edit_message_text("⛔ Нет доступа.")
        return

    idx = int(query.data.replace("del_pending_", ""))
    items = sorted(pending_manager.get_all(), key=lambda x: x["date"])
    if idx >= len(items):
        await query.edit_message_text("Транзакция не найдена (уже удалена?).")
        return

    # Найдём реальный индекс в несортированном списке
    target = items[idx]
    all_items = pending_manager.get_all()
    for real_idx, item in enumerate(all_items):
        if (item["date"] == target["date"] and
                item["amount"] == target["amount"] and
                item["comment"] == target["comment"]):
            pending_manager.remove_by_index(real_idx)
            break

    direction = "Доход" if target["amount"] > 0 else "Расход"
    await query.edit_message_text(
        f"🗑 Удалено: {fmt_date_ru(target['date'])} | {direction}: {fmt_money(target['amount'])}"
        + (f" | {target['comment']}" if target["comment"] else "")
    )


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

    future_mark = " (будет записано в день наступления)" if chosen_date > date.today() else ""
    await query.edit_message_text(
        f"Дата: {fmt_date_ru(chosen_date)}{future_mark}\n\n"
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

    today = date.today()
    if target_date > today:
        pending_manager.add(target_date, amount, "")
        direction = "Доход" if amount > 0 else "Расход"
        await query.edit_message_text(
            f"🕐 Запланировано на {fmt_date_ru(target_date)}\n"
            f"{direction}: {fmt_money(amount)}\n"
            f"Без комментария."
        )
    else:
        excel_helper.add_transaction(target_date, amount, "")
        direction = "Доход" if amount > 0 else "Расход"
        await query.edit_message_text(
            f"✅ Записано на {fmt_date_ru(target_date)}\n"
            f"{direction}: {fmt_money(amount)}\n"
            f"Без комментария.\n\n"
            + totals_text(target_date.year)
        )


# ── Основной обработчик текста ────────────────────────────────────────────────

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return

    text = update.message.text.strip()
    mode = context.user_data.get("mode")
    today = date.today()

    # ── Ждём дату вручную
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
        future_mark = " (будет записано в день наступления)" if chosen_date > today else ""
        await update.message.reply_text(
            f"Дата: {fmt_date_ru(chosen_date)}{future_mark}\n\n"
            "Отправь сумму с + или −:\n"
            "<code>+5000 аванс</code>\n"
            "<code>-1200</code>",
            parse_mode="HTML",
        )
        return

    # ── Ждём сумму для выбранной даты
    if mode == MODE_WAIT_AMOUNT_FOR_DATE:
        amount, comment = parse_amount(text)
        if amount is None or amount == 0:
            await update.message.reply_text(
                "Не понял. Отправь сумму с + или −, например:\n"
                "<code>+5000</code> или <code>-800 кафе</code>",
                parse_mode="HTML",
            )
            return
        target_date = context.user_data.get("add_date", today)
        if comment:
            await _save_transaction(update, context, target_date, amount, comment)
        else:
            context.user_data["mode"] = MODE_WAIT_COMMENT_FOR_DATE
            context.user_data["pending_amount"] = amount
            context.user_data["pending_date"] = target_date
            day_info = excel_helper.get_day_info(target_date)
            direction = "доход" if amount > 0 else "расход"
            text_ask = f"Записать {direction} {fmt_money(amount)} на {fmt_date_ru(target_date)}.\nНапиши комментарий:"
            if day_info["has_records"] and target_date <= today:
                kb = [[InlineKeyboardButton("Пропустить комментарий", callback_data="skip_comment")]]
                await update.message.reply_text(text_ask, reply_markup=InlineKeyboardMarkup(kb))
            else:
                await update.message.reply_text(text_ask)
        return

    # ── Ждём комментарий для выбранной даты
    if mode == MODE_WAIT_COMMENT_FOR_DATE:
        amount = context.user_data.pop("pending_amount", None)
        target_date = context.user_data.pop("pending_date", today)
        clear_mode(context)
        if amount is not None:
            await _save_transaction(update, context, target_date, amount, text)
        return

    # ── Ждём комментарий для сегодня
    if mode == MODE_WAIT_COMMENT:
        amount = context.user_data.pop("pending_amount", None)
        target_date = context.user_data.pop("pending_date", today)
        clear_mode(context)
        if amount is not None:
            await _save_transaction(update, context, target_date, amount, text)
        return

    # ── Быстрый ввод суммы (сегодня)
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
        if comment:
            await _save_transaction(update, context, today, amount, comment)
        else:
            context.user_data["mode"] = MODE_WAIT_COMMENT
            context.user_data["pending_amount"] = amount
            context.user_data["pending_date"] = today
            day_info = excel_helper.get_day_info(today)
            direction = "доход" if amount > 0 else "расход"
            text_ask = f"Записать {direction} {fmt_money(amount)} на {fmt_date_ru(today)}.\nНапиши комментарий:"
            if day_info["has_records"]:
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


async def _save_transaction(update, context, target_date: date, amount: int, comment: str):
    """Сохраняет транзакцию: в Excel если дата <= сегодня, иначе в pending."""
    today = date.today()
    clear_mode(context)
    direction = "Доход" if amount > 0 else "Расход"

    if target_date > today:
        pending_manager.add(target_date, amount, comment)
        await update.message.reply_text(
            f"🕐 Запланировано на {fmt_date_ru(target_date)}\n"
            f"{direction}: {fmt_money(amount)}"
            + (f"\nКомментарий: {comment}" if comment else "")
        )
    else:
        excel_helper.add_transaction(target_date, amount, comment)
        await update.message.reply_text(
            f"✅ Записано на {fmt_date_ru(target_date)}\n"
            f"{direction}: {fmt_money(amount)}"
            + (f"\nКомментарий: {comment}" if comment else "")
            + "\n\n" + totals_text(target_date.year)
        )


# ── Запуск ────────────────────────────────────────────────────────────────────

def main():
    builder = Application.builder().token(BOT_TOKEN)
    if PROXY_URL:
        builder = builder.proxy(PROXY_URL).get_updates_proxy(PROXY_URL)
    app = builder.build()

    # Flush pending при старте
    async def _flush_job(ctx):
        await flush_pending(app)
    app.job_queue.run_once(_flush_job, when=1)

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("totals", cmd_totals))
    app.add_handler(CommandHandler("download", cmd_download))
    app.add_handler(CommandHandler("upload", cmd_upload))
    app.add_handler(CommandHandler("add", cmd_add))
    app.add_handler(CommandHandler("pending", cmd_pending))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(CallbackQueryHandler(callback_date_pick, pattern="^date_"))
    app.add_handler(CallbackQueryHandler(callback_skip_comment, pattern="^skip_comment$"))
    app.add_handler(CallbackQueryHandler(callback_del_pending, pattern="^del_pending_"))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    logger.info("Бот запущен.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
