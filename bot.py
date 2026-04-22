import os
import io
import re
import json
import logging
import asyncio
from datetime import datetime, date

import anthropic
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from telegram import Update, Document
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# ── КОНФИГ ──────────────────────────────────────────────────────────────────
TELEGRAM_TOKEN    = os.getenv("TELEGRAM_TOKEN")
ALLOWED_CHAT_ID   = int(os.getenv("ALLOWED_CHAT_ID"))   # chat_id Меруерт
SPREADSHEET_ID    = os.getenv("SPREADSHEET_ID")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
GOOGLE_CREDS_JSON = os.getenv("GOOGLE_CREDS_JSON", "/etc/secrets/credentials.json")

SHEET_MASTER  = "Мастер (объединённая)"
SHEET_REESTR  = "Реестр договоров"
SHEET_LOG     = "Лог платежей"

# Индексы столбцов в "Мастер (объединённая)" (1-based)
# Группа ПЛАН: col1=№ПЗ, col2=Департамент ... col10=Плановая сумма
# Группа ДОГОВОР: col13=№договора, col14=Поставщик ... col21=Экономия
# Группа ДС: col22-24
# Группа ОПЛАТА: col25=Оплачено, col26=Остаток, col27=Статус, col28=Дата оплаты, col29=Обеспечение
COL_NUM_DOG  = 13  # M - Номер договора
COL_SUM_DOG  = 18  # R - Сумма без НДС (для расчёта остатка)
COL_OPLACH   = 25  # Y - Оплачено (всего)
COL_STATUS   = 27  # AA - Статус оплаты
COL_DATE_OPL = 28  # AB - Дата оплаты факт

# Индексы в "Реестр договоров"
COL_REG_NUM  = 2   # B - Номер договора
COL_REG_SUM  = 10  # J - Сумма без НДС
COL_REG_PAID = 13  # M - Оплачено
COL_REG_STAT = 20  # T - Статус оплаты

# ── GOOGLE SHEETS ────────────────────────────────────────────────────────────
def get_sheets_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(GOOGLE_CREDS_JSON, scopes=scopes)
    return gspread.authorize(creds)

def get_master_sheet():
    gc = get_sheets_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    return sh.worksheet(SHEET_MASTER)

def get_reestr_sheet():
    gc = get_sheets_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    return sh.worksheet(SHEET_REESTR)

def get_log_sheet():
    gc = get_sheets_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    return sh.worksheet(SHEET_LOG)

def normalize(s: str) -> str:
    """Нормализует строку — заменяет кириллические буквы на латинские для сравнения."""
    return s.strip().upper().replace('А', 'A').replace('В', 'B').replace('С', 'C').replace('Е', 'E').replace('О', 'O').replace('Р', 'P').replace('Х', 'X')

def find_contract_row(ws, num_dog: str) -> int | None:
    """Ищет строку по номеру договора. Возвращает номер строки (1-based) или None."""
    num_dog_norm = normalize(num_dog)
    col_values = ws.col_values(COL_NUM_DOG)
    # Точное совпадение с нормализацией
    for i, val in enumerate(col_values):
        if val and normalize(val) == num_dog_norm:
            return i + 1
    # Нечёткий поиск — ищем по части номера
    short = num_dog_norm.replace('AGP/', '').replace('AGP\\', '')
    for i, val in enumerate(col_values):
        if val and short in normalize(val):
            return i + 1
    return None

def update_payment(ws, row: int, amount: float, pay_date: str):
    """Обновляет оплату в Мастер (объединённая)."""
    current = ws.cell(row, COL_OPLACH).value
    try:
        current_val = float(str(current).replace(' ', '').replace(',', '.')) if current else 0
    except:
        current_val = 0
    new_val = current_val + amount

    summa_str = ws.cell(row, COL_SUM_DOG).value
    try:
        summa = float(str(summa_str).replace(' ', '').replace(',', '.')) if summa_str else 0
    except:
        summa = 0

    if summa > 0:
        status = "Оплачено" if new_val >= summa * 0.999 else ("Частично" if new_val > 0 else "Ожидается")
    else:
        status = "Частично" if new_val > 0 else "Ожидается"

    ws.update_cell(row, COL_OPLACH, new_val)
    ws.update_cell(row, COL_STATUS, status)
    if pay_date:
        ws.update_cell(row, COL_DATE_OPL, pay_date)

def update_payment_reestr(num_dog: str, amount: float, pay_date: str):
    """Параллельно обновляет оплату в листе Реестр договоров."""
    try:
        ws = get_reestr_sheet()
        col_vals = ws.col_values(COL_REG_NUM)
        row = None
        num_norm = normalize(num_dog)
        for i, val in enumerate(col_vals):
            if val and normalize(val) == num_norm:
                row = i + 1
                break
        if not row:
            return
        current = ws.cell(row, COL_REG_PAID).value
        try:
            current_val = float(str(current).replace(' ', '').replace(',', '.')) if current else 0
        except:
            current_val = 0
        new_val = current_val + amount
        summa_str = ws.cell(row, COL_REG_SUM).value
        try:
            summa = float(str(summa_str).replace(' ', '').replace(',', '.')) if summa_str else 0
        except:
            summa = 0
        status = "Оплачено" if (summa > 0 and new_val >= summa * 0.999) else ("Частично" if new_val > 0 else "Ожидается")
        ws.update_cell(row, COL_REG_PAID, new_val)
        ws.update_cell(row, COL_REG_STAT, status)
    except Exception as e:
        logger.warning(f"Не удалось обновить Реестр: {e}")

def log_payment(num_dog: str, amount: float, pay_date: str, source: str, comment: str = ""):
    """Записывает платёж в лог."""
    ws = get_log_sheet()
    now = datetime.now().strftime("%d.%m.%Y %H:%M")
    ws.append_row([now, num_dog, amount, pay_date, source, comment])

# ── CLAUDE OCR ───────────────────────────────────────────────────────────────
def extract_payment_from_image(image_bytes: bytes) -> dict:
    """Использует Claude для распознавания платёжки."""
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    import base64
    b64 = base64.standard_b64encode(image_bytes).decode("utf-8")

    message = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=500,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}
                },
                {
                    "type": "text",
                    "text": """Это платёжное поручение казахстанской компании.
Извлеки следующие данные и верни ТОЛЬКО JSON без пояснений:
{
  "num_dog": "номер договора в формате AGP/GEN/XX-NN/YYYY или похожем (или пустая строка)",
  "summa": число (сумма платежа в тенге, только цифры),
  "data": "дата в формате ДД.ММ.ГГГГ (или пустая строка)"
}
Если данные не найдены — верни пустые значения."""
                }
            ]
        }]
    )

    text = message.content[0].text.strip()
    text = text.replace("```json", "").replace("```", "").strip()
    try:
        return json.loads(text)
    except:
        return {"num_dog": "", "summa": 0, "data": ""}

def extract_payments_from_excel(file_bytes: bytes) -> list[dict]:
    """Читает Excel выгрузку платежей и извлекает данные."""
    df = pd.read_excel(io.BytesIO(file_bytes), dtype=str)

    # Ищем столбцы с номером договора и суммой
    num_dog_col = None
    summa_col   = None
    date_col    = None

    for col in df.columns:
        col_lower = str(col).lower()
        if any(x in col_lower for x in ['договор', 'contract', '№']):
            num_dog_col = col
        if any(x in col_lower for x in ['сумма', 'sum', 'amount', 'оплата']):
            summa_col = col
        if any(x in col_lower for x in ['дата', 'date']):
            date_col = col

    payments = []
    for _, row in df.iterrows():
        num = str(row.get(num_dog_col, '')).strip() if num_dog_col else ''
        # Ищем AGP номер в любом столбце если не нашли
        if not num or num == 'nan':
            for val in row.values:
                if val and re.search(r'AGP|АGP', str(val), re.IGNORECASE):
                    match = re.search(r'[АA]GP/\w+/\w+-\d+/\d{4}', str(val))
                    if match:
                        num = match.group()
                        break

        summa_raw = str(row.get(summa_col, '0')).strip() if summa_col else '0'
        try:
            summa = float(re.sub(r'[^\d.]', '', summa_raw.replace(',', '.')))
        except:
            summa = 0

        date_val = str(row.get(date_col, '')).strip() if date_col else ''

        if num and num != 'nan' and summa > 0:
            payments.append({
                "num_dog": num,
                "summa": summa,
                "data": date_val if date_val != 'nan' else ''
            })

    return payments

# ── TELEGRAM HANDLERS ────────────────────────────────────────────────────────
def is_allowed(update: Update) -> bool:
    return update.effective_chat.id == ALLOWED_CHAT_ID

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update): return
    await update.message.reply_text(
        "👋 Здравствуйте, Меруерт!\n\n"
        "Я ваш помощник по закупкам. Вот что я умею:\n\n"
        "📎 Скиньте Excel-выгрузку с оплатами — я обновлю таблицу\n"
        "📷 Скиньте фото платёжки — я считаю данные и внесу\n"
        "📊 /report — сводка по договорам\n"
        "🔍 /find [номер] — найти договор\n"
        "⚠️ /expiring — договора истекающие через 30 дней\n"
        "🔄 /sync — инструкция как обновить ваш Excel файл"
    )

async def cmd_sync(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update): return
    sid = SPREADSHEET_ID or "SPREADSHEET_ID"
    url_master = f"https://docs.google.com/spreadsheets/d/{sid}/export?format=csv&sheet=Мастер (объединённая)"
    url_reestr  = f"https://docs.google.com/spreadsheets/d/{sid}/export?format=csv&sheet=Реестр договоров"
    text = (
        "🔄 *Как обновить ваш Excel файл*\n\n"
        "Данные автоматически обновляются в Google Sheets после каждого платежа.\n\n"
        "Чтобы подтянуть данные в ваш Excel:\n\n"
        "*Первый раз (настройка один раз):*\n"
        "1. Откройте ваш Excel\n"
        "2. Данные → Получить данные → Из интернета\n"
        "3. Вставьте ссылку для нужного листа (см. ниже)\n"
        "4. Нажмите Загрузить → Преобразовать данные\n"
        "5. В Power Query: Закрыть и загрузить\n\n"
        "*Каждый раз после обновления:*\n"
        "Данные → Обновить все (или Ctrl+Alt+F5)\n\n"
        "*Ссылки для Power Query:*\n\n"
        f"Объединённая таблица:\n`{url_master}`\n\n"
        f"Реестр договоров:\n`{url_reestr}`"
    )
    await update.message.reply_text(text, parse_mode='Markdown')

async def cmd_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update): return
    await update.message.reply_text("⏳ Формирую сводку...")
    try:
        ws = get_master_sheet()
        data = ws.get_all_values()
        if len(data) < 2:
            await update.message.reply_text("Таблица пуста.")
            return

        headers = data[0]
        rows = data[1:]

        total = len(rows)
        paid = sum(1 for r in rows if len(r) > 13 and r[13] == 'Оплачено')
        partial = sum(1 for r in rows if len(r) > 13 and r[13] == 'Частично')
        waiting = sum(1 for r in rows if len(r) > 13 and r[13] == 'Ожидается')

        text = (
            f"📊 *Сводка по договорам*\n\n"
            f"Всего договоров: {total}\n"
            f"✅ Оплачено: {paid}\n"
            f"🟡 Частично: {partial}\n"
            f"🔴 Ожидается: {waiting}\n\n"
            f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
        )
        await update.message.reply_text(text, parse_mode='Markdown')
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка: {e}")

async def cmd_expiring(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update): return
    await update.message.reply_text("⏳ Проверяю сроки...")
    try:
        ws = get_master_sheet()
        rows = ws.get_all_values()
        today = date.today()
        expiring = []

        for row in rows[1:]:
            if len(row) < 9: continue
            date_str = row[8]  # Дата окончания
            if not date_str: continue
            try:
                d = datetime.strptime(date_str, '%d.%m.%Y').date()
                delta = (d - today).days
                if 0 <= delta <= 30:
                    expiring.append((delta, row[1], row[2], date_str))
            except:
                continue

        if not expiring:
            await update.message.reply_text("✅ Договоров истекающих через 30 дней нет.")
            return

        expiring.sort()
        lines = [f"⚠️ *Договора истекают в ближайшие 30 дней:*\n"]
        for delta, num, supplier, dt in expiring[:15]:
            lines.append(f"• {num}\n  {supplier[:40]}\n  Истекает: {dt} (через {delta} дн.)\n")

        await update.message.reply_text('\n'.join(lines), parse_mode='Markdown')
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка: {e}")

async def cmd_find(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update): return
    if not context.args:
        await update.message.reply_text("Укажите номер договора: /find AGP/GEN/TS-01/2025")
        return
    num_dog = ' '.join(context.args)
    try:
        ws = get_master_sheet()
        row_num = find_contract_row(ws, num_dog)
        if not row_num:
            await update.message.reply_text(f"❌ Договор {num_dog} не найден.")
            return
        row = ws.row_values(row_num)
        text = (
            f"🔍 *Договор найден:*\n\n"
            f"№: {row[1] if len(row)>1 else '-'}\n"
            f"Поставщик: {row[2] if len(row)>2 else '-'}\n"
            f"Предмет: {str(row[3])[:80] if len(row)>3 else '-'}\n"
            f"Дата окончания: {row[8] if len(row)>8 else '-'}\n"
            f"Сумма: {row[9] if len(row)>9 else '-'} ₸\n"
            f"Оплачено: {row[12] if len(row)>12 else '-'} ₸\n"
            f"Статус: {row[14] if len(row)>14 else '-'}"
        )
        await update.message.reply_text(text, parse_mode='Markdown')
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка: {e}")

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает Excel выгрузку с оплатами."""
    if not is_allowed(update): return
    doc: Document = update.message.document
    fname = doc.file_name or ''

    if not any(fname.endswith(ext) for ext in ['.xlsx', '.xls', '.xlsb']):
        await update.message.reply_text("Поддерживаются файлы .xlsx и .xls")
        return

    await update.message.reply_text("⏳ Читаю файл с оплатами...")
    try:
        file = await doc.get_file()
        file_bytes = bytes(await file.download_as_bytearray())
        payments = extract_payments_from_excel(file_bytes)

        if not payments:
            await update.message.reply_text(
                "❌ Не нашёл данные об оплатах в файле.\n"
                "Убедитесь что в файле есть столбцы с номером договора и суммой."
            )
            return

        await update.message.reply_text(f"Найдено {len(payments)} платежей. Обновляю таблицу...")

        ws = get_master_sheet()
        updated = []
        not_found = []

        for p in payments:
            row_num = find_contract_row(ws, p['num_dog'])
            if row_num:
                update_payment(ws, row_num, p['summa'], p['data'])
                update_payment_reestr(p['num_dog'], p['summa'], p['data'])
                log_payment(p['num_dog'], p['summa'], p['data'], 'Excel выгрузка')
                updated.append(p['num_dog'])
            else:
                not_found.append(p['num_dog'])

        result = f"✅ Обновлено: {len(updated)} договоров\n"
        if updated:
            result += "\n".join(f"  • {n}" for n in updated[:10])
            if len(updated) > 10:
                result += f"\n  ... и ещё {len(updated)-10}"
        if not_found:
            result += f"\n\n⚠️ Не найдено в таблице: {len(not_found)}\n"
            result += "\n".join(f"  • {n}" for n in not_found[:5])

        await update.message.reply_text(result)
    except Exception as e:
        logger.error(f"Ошибка обработки файла: {e}")
        await update.message.reply_text(f"❌ Ошибка при обработке файла: {e}")

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает фото платёжного поручения через Claude OCR."""
    if not is_allowed(update): return
    await update.message.reply_text("📷 Распознаю платёжку...")
    try:
        photo = update.message.photo[-1]
        file = await photo.get_file()
        image_bytes = bytes(await file.download_as_bytearray())

        result = extract_payment_from_image(image_bytes)
        num_dog = result.get('num_dog', '')
        summa   = result.get('summa', 0)
        data    = result.get('data', '')

        if not num_dog or not summa:
            await update.message.reply_text(
                f"🔍 Распознал:\n"
                f"Номер договора: {num_dog or 'не найден'}\n"
                f"Сумма: {summa or 'не найдена'}\n"
                f"Дата: {data or 'не найдена'}\n\n"
                f"❌ Недостаточно данных для обновления таблицы.\n"
                f"Попробуйте прислать более чёткое фото."
            )
            return

        await update.message.reply_text(
            f"🔍 Распознал:\n"
            f"• Номер договора: {num_dog}\n"
            f"• Сумма: {summa:,.0f} ₸\n"
            f"• Дата: {data}\n\n"
            f"Ищу в таблице..."
        )

        ws = get_master_sheet()
        row_num = find_contract_row(ws, num_dog)

        if not row_num:
            await update.message.reply_text(
                f"❌ Договор {num_dog} не найден в таблице.\n"
                f"Проверьте номер договора."
            )
            return

        row = ws.row_values(row_num)
        supplier = row[2] if len(row) > 2 else ''

        # Подтверждение перед записью
        context.user_data['pending'] = {
            'num_dog': num_dog,
            'summa': summa,
            'data': data,
            'row_num': row_num,
        }
        await update.message.reply_text(
            f"✅ Нашла договор:\n"
            f"• {num_dog}\n"
            f"• {supplier[:50]}\n\n"
            f"Записать оплату {summa:,.0f} ₸ от {data}?\n\n"
            f"Ответьте *да* для подтверждения или *нет* для отмены.",
            parse_mode='Markdown'
        )
    except Exception as e:
        logger.error(f"Ошибка OCR: {e}")
        await update.message.reply_text(f"❌ Ошибка распознавания: {e}")

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает текстовые ответы (подтверждение/отмена)."""
    if not is_allowed(update): return
    text = update.message.text.strip().lower()

    # Подтверждение платежа из OCR
    if 'pending' in context.user_data:
        if text in ['да', 'yes', 'подтвердить', '+']:
            p = context.user_data.pop('pending')
            try:
                ws = get_master_sheet()
                update_payment(ws, p['row_num'], p['summa'], p['data'])
                update_payment_reestr(p['num_dog'], p['summa'], p['data'])
                log_payment(p['num_dog'], p['summa'], p['data'], 'Фото платёжки')
                await update.message.reply_text(
                    f"✅ Готово! Оплата внесена:\n"
                    f"• Договор: {p['num_dog']}\n"
                    f"• Сумма: {p['summa']:,.0f} ₸\n"
                    f"• Дата: {p['data']}"
                )
            except Exception as e:
                await update.message.reply_text(f"❌ Ошибка записи: {e}")
        elif text in ['нет', 'no', 'отмена', '-']:
            context.user_data.pop('pending', None)
            await update.message.reply_text("Отменено.")
        return

    await update.message.reply_text(
        "Я понимаю:\n"
        "📎 Excel файл с оплатами\n"
        "📷 Фото платёжки\n"
        "📊 /report — сводка\n"
        "🔍 /find [номер] — найти договор\n"
        "⚠️ /expiring — истекающие договора\n"
        "🔄 /sync — как обновить ваш Excel"
    )

# ── ЗАПУСК ───────────────────────────────────────────────────────────────────
def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("report", cmd_report))
    app.add_handler(CommandHandler("expiring", cmd_expiring))
    app.add_handler(CommandHandler("find", cmd_find))
    app.add_handler(CommandHandler("sync", cmd_sync))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    logger.info("Бот запущен!")
    app.run_polling()

if __name__ == "__main__":
    main()
