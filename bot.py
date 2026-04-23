import base64
import io
import json
import logging
import os
import re
from calendar import monthrange
from datetime import date, datetime

import anthropic
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from telegram import Document, Update
from telegram.ext import Application, CommandHandler, ContextTypes, MessageHandler, filters

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
ALLOWED_CHAT_ID = int(os.getenv("ALLOWED_CHAT_ID"))
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
GOOGLE_CREDS_JSON = os.getenv("GOOGLE_CREDS_JSON", "/etc/secrets/credentials.json")

SHEET_PLAN = "План"
SHEET_FACT = "Факт"
SHEET_LINK = "Связка"
SHEET_LOG = "Лог платежей"

COL_PLAN_NUM = 1
COL_PLAN_DEPT = 2
COL_PLAN_NAME = 3
COL_PLAN_METHOD = 5
COL_PLAN_MONTH = 6
COL_PLAN_AMOUNT = 11
COL_PLAN_SHEET = 16

COL_FACT_NUM = 2
COL_FACT_SUPP = 3
COL_FACT_SUBJ = 4
COL_FACT_DEPT = 5
COL_FACT_START = 8
COL_FACT_END = 9
COL_FACT_SUM = 10
COL_FACT_NOTE = 12
COL_FACT_PAY_DATE = 13
COL_FACT_PAID = 14
COL_FACT_REM = 15
COL_FACT_STAT = 16
COL_FACT_SEC = 17
COL_FACT_DS = 18
COL_FACT_DS_STATUS = 30

COL_LINK_NUM = 1
COL_LINK_PLAN = 2

MONTHS_RU = {
    "январ": 1,
    "феврал": 2,
    "март": 3,
    "апрел": 4,
    "ма": 5,
    "июн": 6,
    "июл": 7,
    "август": 8,
    "сентябр": 9,
    "октябр": 10,
    "ноябр": 11,
    "декабр": 12,
}


def get_sheets_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(GOOGLE_CREDS_JSON, scopes=scopes)
    return gspread.authorize(creds)



def get_sheet(name: str):
    gc = get_sheets_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    return sh.worksheet(name)



def get_plan_sheet():
    return get_sheet(SHEET_PLAN)



def get_fact_sheet():
    return get_sheet(SHEET_FACT)



def get_link_sheet():
    return get_sheet(SHEET_LINK)



def get_log_sheet():
    return get_sheet(SHEET_LOG)



def normalize(value: str) -> str:
    text = str(value or "").strip().upper()
    text = (
        text.replace("А", "A")
        .replace("В", "B")
        .replace("С", "C")
        .replace("Е", "E")
        .replace("О", "O")
        .replace("Р", "P")
        .replace("Х", "X")
    )
    return re.sub(r"[^A-Z0-9]", "", text)



def parse_amount(value) -> float:
    text = str(value or "").replace(" ", "").replace("\xa0", "").replace(",", ".")
    text = re.sub(r"[^0-9.\-]", "", text)
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0



def format_amount(value) -> str:
    return f"{value:,.0f}".replace(",", " ")



def calc_payment_status(total: float, paid: float) -> str:
    if total > 0 and paid >= total * 0.999:
        return "Оплачено"
    if paid > 0:
        return "Частично"
    return "Ожидается"



def parse_sheet_date(value: str) -> date | None:
    text = str(value or "").strip()
    if not text:
        return None
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    if text.isdigit():
        try:
            base = datetime(1899, 12, 30).date()
            return base.fromordinal(base.toordinal() + int(text))
        except Exception:
            return None
    return None



def parse_plan_month(value: str) -> tuple[date | None, date | None]:
    text = str(value or "").strip().lower()
    if not text:
        return None, None
    year_match = re.search(r"(20\d{2})", text)
    month_num = None
    for key, num in MONTHS_RU.items():
        if key in text:
            month_num = num
            break
    if not month_num or not year_match:
        exact_date = parse_sheet_date(text)
        if exact_date:
            last_day = monthrange(exact_date.year, exact_date.month)[1]
            return exact_date.replace(day=1), exact_date.replace(day=last_day)
        return None, None
    year = int(year_match.group(1))
    start = date(year, month_num, 1)
    end = date(year, month_num, monthrange(year, month_num)[1])
    return start, end



def find_contract_row(ws, num_dog: str, col_idx: int = COL_FACT_NUM) -> int | None:
    num_dog_norm = normalize(num_dog)
    col_values = ws.col_values(col_idx)
    for i, value in enumerate(col_values):
        if value and normalize(value) == num_dog_norm:
            return i + 1

    short = num_dog_norm.replace("AGP", "")
    for i, value in enumerate(col_values):
        if value and short and short in normalize(value):
            return i + 1
    return None



def update_payment_fact(ws, row: int, amount: float, pay_date: str):
    current_paid = parse_amount(ws.cell(row, COL_FACT_PAID).value)
    total = parse_amount(ws.cell(row, COL_FACT_SUM).value)
    new_paid = current_paid + amount
    remaining = max(total - new_paid, 0)
    status = calc_payment_status(total, new_paid)

    ws.update_cell(row, COL_FACT_PAID, new_paid)
    ws.update_cell(row, COL_FACT_REM, remaining)
    ws.update_cell(row, COL_FACT_STAT, status)
    if pay_date:
        ws.update_cell(row, COL_FACT_PAY_DATE, pay_date)



def log_payment(num_dog: str, amount: float, pay_date: str, source: str, comment: str = ""):
    ws = get_log_sheet()
    now = datetime.now().strftime("%d.%m.%Y %H:%M")
    ws.append_row([now, num_dog, amount, pay_date, source, comment])



def get_plan_links() -> dict[str, list[str]]:
    ws = get_link_sheet()
    rows = ws.get_all_values()[1:]
    links: dict[str, list[str]] = {}
    for row in rows:
        if len(row) < COL_LINK_PLAN:
            continue
        contract_num = row[COL_LINK_NUM - 1].strip()
        plan_num = row[COL_LINK_PLAN - 1].strip()
        if not plan_num:
            continue
        links.setdefault(plan_num, [])
        if contract_num:
            links[plan_num].append(contract_num)
    return links



def get_contract_links() -> dict[str, list[str]]:
    ws = get_link_sheet()
    rows = ws.get_all_values()[1:]
    links: dict[str, list[str]] = {}
    for row in rows:
        if len(row) < COL_LINK_PLAN:
            continue
        contract_num = row[COL_LINK_NUM - 1].strip()
        plan_num = row[COL_LINK_PLAN - 1].strip()
        if not contract_num:
            continue
        links.setdefault(normalize(contract_num), [])
        if plan_num:
            links[normalize(contract_num)].append(plan_num)
    return links



def extract_payment_from_image(image_bytes: bytes, media_type: str = "image/jpeg") -> dict:
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    b64 = base64.standard_b64encode(image_bytes).decode("utf-8")

    message = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=500,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {"type": "base64", "media_type": media_type, "data": b64},
                    },
                    {
                        "type": "text",
                        "text": (
                            "Это платежное поручение казахстанской компании.\n"
                            "Извлеки следующие данные и верни только JSON без пояснений:\n"
                            "{\n"
                            '  "num_dog": "номер договора в формате AGP/GEN/XX-NN/YYYY или похожем (или пустая строка)",\n'
                            '  "summa": число,\n'
                            '  "data": "дата в формате ДД.ММ.ГГГГ (или пустая строка)"\n'
                            "}\n"
                            "Если данные не найдены, верни пустые значения."
                        ),
                    },
                ],
            }
        ],
    )

    text = message.content[0].text.strip().replace("```json", "").replace("```", "").strip()
    try:
        data = json.loads(text)
    except Exception:
        return {"num_dog": "", "summa": 0, "data": ""}

    return {
        "num_dog": str(data.get("num_dog", "")).strip(),
        "summa": parse_amount(data.get("summa", 0)),
        "data": str(data.get("data", "")).strip(),
    }



def extract_payments_from_excel(file_bytes: bytes, filename: str = "") -> list[dict]:
    engine = "pyxlsb" if filename.lower().endswith(".xlsb") else None
    workbook = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, dtype=str, engine=engine)
    payments: list[dict] = []

    for df in workbook.values():
        num_dog_col = None
        summa_col = None
        date_col = None

        for col in df.columns:
            col_lower = str(col).lower()
            if any(x in col_lower for x in ["договор", "contract", "№"]):
                num_dog_col = col
            if any(x in col_lower for x in ["сумма", "sum", "amount", "оплата"]):
                summa_col = col
            if any(x in col_lower for x in ["дата", "date"]):
                date_col = col

        for _, row in df.iterrows():
            num = str(row.get(num_dog_col, "")).strip() if num_dog_col else ""
            if not num or num == "nan":
                for value in row.values:
                    if value and re.search(r"AGP|АGP", str(value), re.IGNORECASE):
                        match = re.search(r"[АA]GP/[A-Z0-9-]+/[A-Z0-9-]+/\d{4}", str(value), re.IGNORECASE)
                        if match:
                            num = match.group()
                            break

            amount = parse_amount(row.get(summa_col, "0")) if summa_col else 0
            pay_date = str(row.get(date_col, "")).strip() if date_col else ""

            if num and num != "nan" and amount > 0:
                payments.append({
                    "num_dog": num,
                    "summa": amount,
                    "data": pay_date if pay_date != "nan" else "",
                })

    return payments



def is_allowed(update: Update) -> bool:
    return update.effective_chat.id == ALLOWED_CHAT_ID



async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    await update.message.reply_text(
        "Привет! Я умею:\n\n"
        "Принять Excel с оплатами и обновить лист Факт\n"
        "Принять фото или скан платежки и занести оплату\n"
        "/report — сводка по договорам\n"
        "/find [номер] — найти договор\n"
        "/expiring — договоры, истекающие в ближайшие 30 дней\n"
        "/plan_due — закупки плана, по которым уже пора заключать договор\n"
        "/sync — как обновить Excel из Google Sheets"
    )



async def cmd_sync(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    sid = SPREADSHEET_ID or "SPREADSHEET_ID"
    url_plan = f"https://docs.google.com/spreadsheets/d/{sid}/export?format=csv&sheet={SHEET_PLAN}"
    url_fact = f"https://docs.google.com/spreadsheets/d/{sid}/export?format=csv&sheet={SHEET_FACT}"
    url_link = f"https://docs.google.com/spreadsheets/d/{sid}/export?format=csv&sheet={SHEET_LINK}"
    text = (
        "*Как обновить ваш Excel файл*\n\n"
        "1. Excel -> Данные -> Получить данные -> Из интернета\n"
        "2. Вставьте одну из ссылок ниже\n"
        "3. Загрузите данные через Power Query\n"
        "4. Потом используйте Обновить все\n\n"
        f"План:\n`{url_plan}`\n\n"
        f"Факт:\n`{url_fact}`\n\n"
        f"Связка:\n`{url_link}`"
    )
    await update.message.reply_text(text, parse_mode="Markdown")



async def cmd_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    await update.message.reply_text("Формирую сводку...")
    try:
        ws = get_fact_sheet()
        rows = ws.get_all_values()[1:]
        rows = [row for row in rows if len(row) >= COL_FACT_STAT and row[COL_FACT_NUM - 1]]

        total = len(rows)
        paid = sum(1 for row in rows if row[COL_FACT_STAT - 1] == "Оплачено")
        partial = sum(1 for row in rows if row[COL_FACT_STAT - 1] == "Частично")
        waiting = sum(1 for row in rows if row[COL_FACT_STAT - 1] == "Ожидается")
        paid_amount = sum(parse_amount(row[COL_FACT_PAID - 1]) for row in rows)
        secured = sum(1 for row in rows if row[COL_FACT_SEC - 1].strip())

        text = (
            "*Сводка по договорам*\n\n"
            f"Всего договоров: {total}\n"
            f"Оплачено: {paid}\n"
            f"Частично: {partial}\n"
            f"Ожидается: {waiting}\n"
            f"С обеспечением: {secured}\n"
            f"Оплачено всего: {format_amount(paid_amount)} тг\n\n"
            f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
        )
        await update.message.reply_text(text, parse_mode="Markdown")
    except Exception as exc:
        await update.message.reply_text(f"Ошибка: {exc}")



async def cmd_expiring(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    await update.message.reply_text("Проверяю сроки договоров...")
    try:
        ws = get_fact_sheet()
        today = date.today()
        expiring = []

        for row in ws.get_all_values()[1:]:
            if len(row) < COL_FACT_END or not row[COL_FACT_NUM - 1]:
                continue
            end_date = parse_sheet_date(row[COL_FACT_END - 1])
            if not end_date:
                continue
            delta = (end_date - today).days
            if 0 <= delta <= 30:
                expiring.append((delta, row[COL_FACT_NUM - 1], row[COL_FACT_SUPP - 1], end_date.strftime("%d.%m.%Y")))

        if not expiring:
            await update.message.reply_text("Договоров с окончанием в ближайшие 30 дней нет.")
            return

        expiring.sort()
        lines = ["*Договоры истекают в ближайшие 30 дней:*\n"]
        for delta, num, supplier, dt in expiring[:15]:
            lines.append(f"• {num}\n{supplier[:60]}\nИстекает: {dt} (через {delta} дн.)\n")

        await update.message.reply_text("\n".join(lines), parse_mode="Markdown")
    except Exception as exc:
        await update.message.reply_text(f"Ошибка: {exc}")



async def cmd_find(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    if not context.args:
        await update.message.reply_text("Укажите номер договора: /find AGP/GEN/TS-01/2025")
        return

    num_dog = " ".join(context.args)
    try:
        ws = get_fact_sheet()
        row_num = find_contract_row(ws, num_dog, COL_FACT_NUM)
        if not row_num:
            await update.message.reply_text(f"Договор {num_dog} не найден.")
            return

        row = ws.row_values(row_num)
        contract_links = get_contract_links().get(normalize(num_dog), [])
        link_text = ", ".join(contract_links) if contract_links else "не привязан"
        text = (
            "*Договор найден:*\n\n"
            f"№: {row[COL_FACT_NUM - 1] if len(row) >= COL_FACT_NUM else '-'}\n"
            f"Поставщик: {row[COL_FACT_SUPP - 1] if len(row) >= COL_FACT_SUPP else '-'}\n"
            f"Предмет: {row[COL_FACT_SUBJ - 1] if len(row) >= COL_FACT_SUBJ else '-'}\n"
            f"Дата окончания: {row[COL_FACT_END - 1] if len(row) >= COL_FACT_END else '-'}\n"
            f"Сумма: {row[COL_FACT_SUM - 1] if len(row) >= COL_FACT_SUM else '-'} тг\n"
            f"Оплачено: {row[COL_FACT_PAID - 1] if len(row) >= COL_FACT_PAID else '-'} тг\n"
            f"Остаток: {row[COL_FACT_REM - 1] if len(row) >= COL_FACT_REM else '-'} тг\n"
            f"Обеспечение: {row[COL_FACT_SEC - 1] if len(row) >= COL_FACT_SEC and row[COL_FACT_SEC - 1] else 'не указано'}\n"
            f"Статус: {row[COL_FACT_STAT - 1] if len(row) >= COL_FACT_STAT else '-'}\n"
            f"Позиции плана: {link_text}"
        )
        await update.message.reply_text(text, parse_mode="Markdown")
    except Exception as exc:
        await update.message.reply_text(f"Ошибка: {exc}")



async def cmd_plan_due(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    await update.message.reply_text("Проверяю план закупок...")
    try:
        ws = get_plan_sheet()
        links = get_plan_links()
        today = date.today()
        limit = date(today.year, today.month, 1)
        if today.month == 12:
            next_month = date(today.year + 1, 1, 1)
        else:
            next_month = date(today.year, today.month + 1, 1)
        horizon = date(next_month.year, next_month.month, monthrange(next_month.year, next_month.month)[1])

        overdue = []
        due_now = []
        upcoming = []

        for row in ws.get_all_values()[1:]:
            if len(row) < COL_PLAN_MONTH or not row[COL_PLAN_NUM - 1]:
                continue
            plan_num = row[COL_PLAN_NUM - 1].strip()
            if links.get(plan_num):
                continue

            month_start, month_end = parse_plan_month(row[COL_PLAN_MONTH - 1])
            if not month_start:
                continue

            item = (
                plan_num,
                row[COL_PLAN_NAME - 1],
                row[COL_PLAN_DEPT - 1],
                row[COL_PLAN_MONTH - 1],
                row[COL_PLAN_METHOD - 1],
                row[COL_PLAN_AMOUNT - 1],
            )
            if month_end < today:
                overdue.append(item)
            elif month_start <= horizon:
                if month_start <= limit:
                    due_now.append(item)
                else:
                    upcoming.append(item)

        if not overdue and not due_now and not upcoming:
            await update.message.reply_text("По плану нет закупок без договора, требующих внимания в ближайший период.")
            return

        lines = ["*Закупки плана, по которым пора заключать договор:*\n"]

        if overdue:
            lines.append("Просрочено:")
            for num, name, dept, month_text, method, amount in overdue[:7]:
                lines.append(f"• {num} | {dept} | {month_text}\n{name[:70]}\n{method}, {amount} тг\n")

        if due_now:
            lines.append("Текущий период:")
            for num, name, dept, month_text, method, amount in due_now[:7]:
                lines.append(f"• {num} | {dept} | {month_text}\n{name[:70]}\n{method}, {amount} тг\n")

        if upcoming:
            lines.append("Ближайший месяц:")
            for num, name, dept, month_text, method, amount in upcoming[:7]:
                lines.append(f"• {num} | {dept} | {month_text}\n{name[:70]}\n{method}, {amount} тг\n")

        await update.message.reply_text("\n".join(lines), parse_mode="Markdown")
    except Exception as exc:
        await update.message.reply_text(f"Ошибка: {exc}")



async def process_payment_image(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    image_bytes: bytes,
    media_type: str = "image/jpeg",
):
    await update.message.reply_text("Распознаю платежку...")
    try:
        result = extract_payment_from_image(image_bytes, media_type=media_type)
        num_dog = result.get("num_dog", "")
        summa = parse_amount(result.get("summa", 0))
        pay_date = str(result.get("data", "")).strip()

        if not num_dog or not summa:
            await update.message.reply_text(
                "Недостаточно данных для обновления таблицы.\n"
                f"Номер договора: {num_dog or 'не найден'}\n"
                f"Сумма: {summa or 'не найдена'}\n"
                f"Дата: {pay_date or 'не найдена'}"
            )
            return

        ws = get_fact_sheet()
        row_num = find_contract_row(ws, num_dog, COL_FACT_NUM)
        if not row_num:
            await update.message.reply_text(f"Договор {num_dog} не найден в листе Факт.")
            return

        row = ws.row_values(row_num)
        supplier = row[COL_FACT_SUPP - 1] if len(row) >= COL_FACT_SUPP else ""
        context.user_data["pending"] = {
            "num_dog": num_dog,
            "summa": summa,
            "data": pay_date,
            "fact_row_num": row_num,
        }

        await update.message.reply_text(
            f"Найден договор:\n"
            f"• {num_dog}\n"
            f"• {supplier[:60]}\n\n"
            f"Записать оплату {format_amount(summa)} тг от {pay_date or '-'}?\n"
            f"Ответьте *да* для подтверждения или *нет* для отмены.",
            parse_mode="Markdown",
        )
    except Exception as exc:
        logger.error(f"Ошибка OCR: {exc}")
        await update.message.reply_text(f"Ошибка распознавания: {exc}")



async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return

    doc: Document = update.message.document
    fname = (doc.file_name or "").lower()

    try:
        file = await doc.get_file()
        file_bytes = bytes(await file.download_as_bytearray())

        if fname.endswith((".jpg", ".jpeg", ".png")):
            media_type = "image/png" if fname.endswith(".png") else "image/jpeg"
            await process_payment_image(update, context, file_bytes, media_type=media_type)
            return

        if not fname.endswith((".xlsx", ".xls", ".xlsb")):
            await update.message.reply_text("Поддерживаются файлы .xlsx, .xls, .xlsb, .jpg, .jpeg, .png")
            return

        await update.message.reply_text("Читаю файл с оплатами...")
        payments = extract_payments_from_excel(file_bytes, fname)
        if not payments:
            await update.message.reply_text(
                "Не нашел данные об оплатах в файле.\n"
                "Проверьте, что там есть номер договора и сумма."
            )
            return

        await update.message.reply_text(f"Найдено {len(payments)} платежей. Обновляю лист Факт...")

        fact_ws = get_fact_sheet()
        updated = []
        not_found = []

        for payment in payments:
            row_num = find_contract_row(fact_ws, payment["num_dog"], COL_FACT_NUM)
            if not row_num:
                not_found.append(payment["num_dog"])
                continue

            update_payment_fact(fact_ws, row_num, payment["summa"], payment["data"])
            log_payment(payment["num_dog"], payment["summa"], payment["data"], "Excel выгрузка")
            updated.append(payment["num_dog"])

        result = f"Обновлено договоров: {len(updated)}"
        if updated:
            result += "\n" + "\n".join(f"• {num}" for num in updated[:10])
            if len(updated) > 10:
                result += f"\n... и еще {len(updated) - 10}"
        if not_found:
            result += f"\n\nНе найдено в Факте: {len(not_found)}"
            result += "\n" + "\n".join(f"• {num}" for num in not_found[:5])

        await update.message.reply_text(result)
    except Exception as exc:
        logger.error(f"Ошибка обработки файла: {exc}")
        await update.message.reply_text(f"Ошибка при обработке файла: {exc}")



async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    photo = update.message.photo[-1]
    file = await photo.get_file()
    image_bytes = bytes(await file.download_as_bytearray())
    await process_payment_image(update, context, image_bytes)



async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return

    text = update.message.text.strip().lower()
    if "pending" in context.user_data:
        if text in ["да", "yes", "подтвердить", "+"]:
            pending = context.user_data.pop("pending")
            try:
                ws = get_fact_sheet()
                update_payment_fact(ws, pending["fact_row_num"], pending["summa"], pending["data"])
                log_payment(pending["num_dog"], pending["summa"], pending["data"], "Фото платежки")
                await update.message.reply_text(
                    "Готово!\n"
                    f"Договор: {pending['num_dog']}\n"
                    f"Сумма: {format_amount(pending['summa'])} тг\n"
                    f"Дата: {pending['data'] or '-'}"
                )
            except Exception as exc:
                await update.message.reply_text(f"Ошибка записи: {exc}")
        elif text in ["нет", "no", "отмена", "-"]:
            context.user_data.pop("pending", None)
            await update.message.reply_text("Отменено.")
        return

    await update.message.reply_text(
        "Я понимаю:\n"
        "Excel файл с оплатами\n"
        "Фото или скан платежки\n"
        "/report — сводка по договорам\n"
        "/find [номер] — найти договор\n"
        "/expiring — истекающие договоры\n"
        "/plan_due — закупки плана без договора\n"
        "/sync — как обновить Excel"
    )



def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("report", cmd_report))
    app.add_handler(CommandHandler("expiring", cmd_expiring))
    app.add_handler(CommandHandler("find", cmd_find))
    app.add_handler(CommandHandler("plan_due", cmd_plan_due))
    app.add_handler(CommandHandler("sync", cmd_sync))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    logger.info("Бот запущен")
    app.run_polling()


if __name__ == "__main__":
    main()
