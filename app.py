import os
import sqlite3
import json
import logging
from io import BytesIO
from datetime import datetime, timedelta
from fastapi import FastAPI, Request, Response
import httpx
from openpyxl import Workbook
from weasyprint import HTML

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("medjol")

app = FastAPI()

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")
OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "meta-llama/llama-3.3-70b-instruct")
TELEGRAM_WEBHOOK_SECRET = os.getenv("TELEGRAM_WEBHOOK_SECRET")

DB_NAME = "harvest.db"


def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS daily_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            workers_count INTEGER,
            wage_per_worker REAL,
            expenses REAL,
            notes TEXT
        )
    ''')
    conn.commit()
    conn.close()


init_db()


# ---------------------------------------------------------------------------
# Telegram helpers
# ---------------------------------------------------------------------------

async def send_telegram_message(chat_id: int, text: str):
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    async with httpx.AsyncClient() as client:
        try:
            res = await client.post(url, json={"chat_id": chat_id, "text": text})
            if res.status_code != 200:
                logger.error(f"Telegram sendMessage failed {res.status_code}: {res.text}")
        except Exception as e:
            logger.error(f"Telegram sendMessage exception: {e}")


async def send_telegram_document(chat_id: int, filename: str, content: bytes, caption: str = ""):
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
    files = {"document": (filename, content)}
    data = {"chat_id": chat_id, "caption": caption}
    async with httpx.AsyncClient() as client:
        try:
            res = await client.post(url, data=data, files=files)
            if res.status_code != 200:
                logger.error(f"Telegram sendDocument failed {res.status_code}: {res.text}")
        except Exception as e:
            logger.error(f"Telegram sendDocument exception: {e}")


# ---------------------------------------------------------------------------
# AI Router — تحديد نية الرسالة: تسجيل / تعديل / تقرير / غير مفهوم
# ---------------------------------------------------------------------------

def build_prompt(text: str, today_date: str, yesterday_date: str, day_before_date: str) -> str:
    return f"""أنت محاسب خبير ودقيق جداً لمزرعة نخيل في الأردن (منطقة الأغوار)، وموجّه ذكي يحدد نية المستخدم بدقة من رسالته.

## الخطوة الأولى — حدد نية الرسالة (intent):
- "record": تسجيل يومية جديدة كاملة (عدد عمال + أجرة، لأول مرة).
- "update": تعديل على سجل مسجّل مسبقاً — إضافة مصروف جديد لسجل موجود، تغيير رقم، حذف سجل، أو إضافة ملاحظة. أي رسالة تبدأ بـ "أضف"، "عدّل"، "غيّر"، "احذف"، "صحح" وتشير لسجل سابق تعتبر update.
- "report": طلب استرجاع بيانات مسجلة (تقرير، كشف، "شو سجلت"...).
- "other": أي شيء آخر لا علاقة له بالتسجيل أو التعديل أو التقارير.

## إذا كانت النية "record": (تسجيل جديد بالكامل)
- إذا أُعطي أجر بالساعة: الأجرة اليومية = الساعات × سعر الساعة.
- عمال بأجور مختلفة: اجمع العدد الكلي، واحسب wage_per_worker كمتوسط مرجّح، واذكر التفصيل بالـ notes.
- مصاريف متعددة: اجمعها برقم واحد بـ expenses، فصّل كل بند بالـ notes.
- الأرقام بالحروف العربية حوّلها لأرقام.
- confidence: "full" إذا وضح عدد العمال والأجرة، "partial" إذا نقص عنصر أساسي، "none" إذا لا علاقة.

## إذا كانت النية "update": (تعديل سجل موجود مسبقاً)
حدد الحقول التالية:
- update_target_date: التاريخ المقصود بصيغة YYYY-MM-DD. إذا لم يُذكر تاريخ صراحة وكان الكلام يشير ضمنياً لآخر سجل تم الحديث عنه (مثال: "أضف كذا" بدون تاريخ بعد تسجيل يومية)، اجعلها "last".
- update_action: واحدة من:
  - "add_expense": إضافة مبلغ ثابت للمصاريف الحالية (مثال: "أضف أجرة نقل 25 دينار").
  - "add_expense_per_worker": إضافة مبلغ لكل عامل يُضرب بعدد العمال ويُضاف للمصاريف (مثال: "أضف لكل عامل دينار بدل طعام" ← المبلغ 1 يُضرب بعدد عمال السجل).
  - "set_expense": استبدال المصاريف بقيمة جديدة كلياً (مثال: "خلي المصاريف 50 دينار").
  - "set_wage": تغيير أجرة العامل لقيمة جديدة.
  - "set_workers": تغيير عدد العمال لقيمة جديدة.
  - "delete_record": حذف السجل بالكامل (مثال: "احذف يومية امس", "الغي التسجيل الأخير").
  - "append_note": إضافة ملاحظة نصية فقط بدون تغيير أرقام.
- update_value: القيمة الرقمية المرتبطة بالعملية (رقم واحد). لعملية delete_record أو append_note بدون رقم، اجعلها 0.
- update_note: نص قصير يوضح ماذا حدث (سيُضاف لملاحظات السجل)، أو نص فارغ لو غير مناسب.

مهم جداً: لا تحاول حساب المجموع النهائي بنفسك بهذه الحالة — فقط استخرج نوع العملية وقيمتها، والكود سيطبقها على السجل الفعلي.

## إذا كانت النية "report": حدد نطاق التقرير
- report_scope: "single_day" أو "range" أو "all".
- report_date_from / report_date_to: YYYY-MM-DD أو null لو "all".
- report_format: "pdf" أو "excel" لو ذُكرا صراحة، وإلا "text".

## قواعد التاريخ (اليوم هو {today_date}):
- "امس"/"أمس" → {yesterday_date}
- "اول امس"/"أول أمس"/"قبل امس" → {day_before_date}
- تاريخ محدد → حوّله YYYY-MM-DD
- بدون تاريخ + record → {today_date}
- بدون تاريخ + report → scope = "all"
- بدون تاريخ + update → "last"

## صيغة الإخراج: JSON فقط بدون أي نص إضافي:
{{
  "intent": "record" | "update" | "report" | "other",
  "workers_count": عدد صحيح أو 0,
  "wage_per_worker": رقم أو 0,
  "expenses": رقم أو 0,
  "date": "YYYY-MM-DD",
  "notes": "نص",
  "confidence": "full" | "partial" | "none",
  "missing_info": "نص أو فارغ",
  "update_target_date": "YYYY-MM-DD" أو "last" أو null,
  "update_action": "add_expense" | "add_expense_per_worker" | "set_expense" | "set_wage" | "set_workers" | "delete_record" | "append_note" | null,
  "update_value": رقم أو 0,
  "update_note": "نص أو فارغ",
  "report_scope": "single_day" | "range" | "all" | null,
  "report_date_from": "YYYY-MM-DD" أو null,
  "report_date_to": "YYYY-MM-DD" أو null,
  "report_format": "text" | "pdf" | "excel" | null
}}

## أمثلة:

النص: "اشتغل معي 23 عامل لمدة 6 ساعات اجرة الساعة الواحدة دينار ونصف"
الإخراج: {{"intent": "record", "workers_count": 23, "wage_per_worker": 9.0, "expenses": 0, "date": "{today_date}", "notes": "6 ساعات × 1.5 دينار/ساعة = 9 دينار للعامل", "confidence": "full", "missing_info": "", "update_target_date": null, "update_action": null, "update_value": 0, "update_note": "", "report_scope": null, "report_date_from": null, "report_date_to": null, "report_format": null}}

النص: "أضف أجرة نقل عمال = 25"
الإخراج: {{"intent": "update", "workers_count": 0, "wage_per_worker": 0, "expenses": 0, "date": "", "notes": "", "confidence": "none", "missing_info": "", "update_target_date": "last", "update_action": "add_expense", "update_value": 25, "update_note": "أجرة نقل عمال 25 دينار", "report_scope": null, "report_date_from": null, "report_date_to": null, "report_format": null}}

النص: "أضف لكل عامل دينار واحد بدل طعام"
الإخراج: {{"intent": "update", "workers_count": 0, "wage_per_worker": 0, "expenses": 0, "date": "", "notes": "", "confidence": "none", "missing_info": "", "update_target_date": "last", "update_action": "add_expense_per_worker", "update_value": 1, "update_note": "بدل طعام دينار لكل عامل", "report_scope": null, "report_date_from": null, "report_date_to": null, "report_format": null}}

النص: "احذف يومية امس"
الإخراج: {{"intent": "update", "workers_count": 0, "wage_per_worker": 0, "expenses": 0, "date": "", "notes": "", "confidence": "none", "missing_info": "", "update_target_date": "{yesterday_date}", "update_action": "delete_record", "update_value": 0, "update_note": "", "report_scope": null, "report_date_from": null, "report_date_to": null, "report_format": null}}

النص: "غيّر أجرة العامل ليوم امس تصير 11 دينار"
الإخراج: {{"intent": "update", "workers_count": 0, "wage_per_worker": 0, "expenses": 0, "date": "", "notes": "", "confidence": "none", "missing_info": "", "update_target_date": "{yesterday_date}", "update_action": "set_wage", "update_value": 11, "update_note": "تصحيح أجرة العامل إلى 11 دينار", "report_scope": null, "report_date_from": null, "report_date_to": null, "report_format": null}}

النص: "اعطني تقرير يوم {yesterday_date}"
الإخراج: {{"intent": "report", "workers_count": 0, "wage_per_worker": 0, "expenses": 0, "date": "", "notes": "", "confidence": "none", "missing_info": "", "update_target_date": null, "update_action": null, "update_value": 0, "update_note": "", "report_scope": "single_day", "report_date_from": "{yesterday_date}", "report_date_to": "{yesterday_date}", "report_format": "text"}}

النص: "شو الجو اليوم؟"
الإخراج: {{"intent": "other", "workers_count": 0, "wage_per_worker": 0, "expenses": 0, "date": "", "notes": "", "confidence": "none", "missing_info": "", "update_target_date": null, "update_action": null, "update_value": 0, "update_note": "", "report_scope": null, "report_date_from": null, "report_date_to": null, "report_format": null}}

النص الحالي المطلوب تحليله: "{text}"
"""


async def analyze_message(text: str) -> dict:
    if not OPENROUTER_API_KEY:
        logger.error("OPENROUTER_API_KEY is missing from environment variables!")
        return {"intent": "other"}

    today = datetime.now()
    today_str = today.strftime("%Y-%m-%d")
    yesterday_str = (today - timedelta(days=1)).strftime("%Y-%m-%d")
    day_before_str = (today - timedelta(days=2)).strftime("%Y-%m-%d")

    prompt = build_prompt(text, today_str, yesterday_str, day_before_str)

    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://medjol.onrender.com",
        "X-Title": "Medjol Farm Bot"
    }

    payload = {
        "model": OPENROUTER_MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.0,
        "response_format": {"type": "json_object"}
    }

    async with httpx.AsyncClient() as client:
        try:
            res = await client.post(
                "https://openrouter.ai/api/v1/chat/completions",
                headers=headers,
                json=payload,
                timeout=25.0
            )
            if res.status_code == 200:
                result = res.json()
                content = result['choices'][0]['message']['content']
                logger.info(f"AI raw response: {content}")
                return json.loads(content)
            else:
                logger.error(f"OpenRouter API error {res.status_code}: {res.text}")
        except json.JSONDecodeError as e:
            logger.error(f"AI returned invalid JSON: {e}")
        except Exception as e:
            logger.error(f"AI request exception: {e}")
    return {"intent": "other"}


# ---------------------------------------------------------------------------
# Database helpers
# ---------------------------------------------------------------------------

def query_records(date_from: str = None, date_to: str = None):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    if date_from and date_to:
        cursor.execute(
            "SELECT date, workers_count, wage_per_worker, expenses, notes "
            "FROM daily_records WHERE date BETWEEN ? AND ? ORDER BY date ASC",
            (date_from, date_to)
        )
    else:
        cursor.execute(
            "SELECT date, workers_count, wage_per_worker, expenses, notes "
            "FROM daily_records ORDER BY date ASC"
        )
    rows = cursor.fetchall()
    conn.close()
    return rows


def format_text_report(rows, title: str) -> str:
    if not rows:
        return f"{title}\n\nلا توجد أي سجلات لهذه الفترة."

    lines = [title, ""]
    grand_total = 0.0
    for date, count, wage, exp, notes in rows:
        total = (count * wage) + exp
        grand_total += total
        lines.append(f"📅 {date}")
        lines.append(f"   عدد العمال: {count}")
        lines.append(f"   أجرة العامل: {wage:.2f} د.أ")
        lines.append(f"   المصاريف: {exp:.2f} د.أ")
        lines.append(f"   المجموع: {total:.2f} د.أ")
        if notes:
            lines.append(f"   ملاحظات: {notes}")
        lines.append("")

    lines.append(f"💰 المجموع الكلي لكل الفترة: {grand_total:.2f} دينار أردني")
    return "\n".join(lines)


def find_target_record(target_date: str):
    """يرجع (id, date, workers_count, wage_per_worker, expenses, notes) أو None"""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    if target_date == "last" or not target_date:
        cursor.execute(
            "SELECT id, date, workers_count, wage_per_worker, expenses, notes "
            "FROM daily_records ORDER BY id DESC LIMIT 1"
        )
    else:
        cursor.execute(
            "SELECT id, date, workers_count, wage_per_worker, expenses, notes "
            "FROM daily_records WHERE date = ? ORDER BY id DESC LIMIT 1",
            (target_date,)
        )
    row = cursor.fetchone()
    conn.close()
    return row


def apply_update(record, action: str, value: float, note: str):
    """يعدّل السجل بقاعدة البيانات ويرجع رسالة تأكيد نصية."""
    rec_id, date, workers_count, wage, expenses, notes = record

    if action == "delete_record":
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM daily_records WHERE id = ?", (rec_id,))
        conn.commit()
        conn.close()
        return f"🗑️ تم حذف سجل يوم {date} بالكامل."

    new_workers = workers_count
    new_wage = wage
    new_expenses = expenses
    new_notes = notes

    if action == "add_expense":
        new_expenses = expenses + value
        change_desc = f"إضافة {value:.2f} د.أ للمصاريف"
    elif action == "add_expense_per_worker":
        added = value * workers_count
        new_expenses = expenses + added
        change_desc = f"إضافة {value:.2f} د.أ × {workers_count} عامل = {added:.2f} د.أ للمصاريف"
    elif action == "set_expense":
        new_expenses = value
        change_desc = f"تعديل المصاريف لتصبح {value:.2f} د.أ"
    elif action == "set_wage":
        new_wage = value
        change_desc = f"تعديل أجرة العامل لتصبح {value:.2f} د.أ"
    elif action == "set_workers":
        new_workers = int(value)
        change_desc = f"تعديل عدد العمال ليصبح {int(value)}"
    elif action == "append_note":
        change_desc = "إضافة ملاحظة"
    else:
        change_desc = "تعديل غير معروف"

    extra_note = note or change_desc
    new_notes = f"{notes}; {extra_note}" if notes else extra_note

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute(
        "UPDATE daily_records SET workers_count=?, wage_per_worker=?, expenses=?, notes=? WHERE id=?",
        (new_workers, new_wage, new_expenses, new_notes, rec_id)
    )
    conn.commit()
    conn.close()

    total = (new_workers * new_wage) + new_expenses
    return (
        f"✏️ تم تعديل سجل يوم {date}\n"
        f"- {change_desc}\n\n"
        f"البيانات بعد التعديل:\n"
        f"- عدد العمال: {new_workers}\n"
        f"- أجرة العامل: {new_wage:.2f} د.أ\n"
        f"- المصاريف: {new_expenses:.2f} د.أ\n"
        f"- المجموع الكلي: {total:.2f} د.أ"
    )


# ---------------------------------------------------------------------------
# File reports (PDF / Excel)
# ---------------------------------------------------------------------------

def generate_pdf_report(date_from: str = None, date_to: str = None):
    rows = query_records(date_from, date_to)

    table_rows = ""
    total_cost_all = 0
    for row in rows:
        date, count, wage, exp, notes = row
        total = (count * wage) + exp
        total_cost_all += total
        table_rows += f"""
        <tr>
            <td>{date}</td>
            <td>{count}</td>
            <td>{wage:.2f} د.أ</td>
            <td>{exp:.2f} د.أ</td>
            <td><b>{total:.2f} د.أ</b></td>
            <td>{notes or '-'}</td>
        </tr>
        """

    html_content = f"""
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
        <meta charset="UTF-8">
        <style>
            body {{ font-family: 'DejaVu Sans', sans-serif; padding: 20px; }}
            h1 {{ text-align: center; color: #2c3e50; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: center; }}
            th {{ background-color: #27ae60; color: white; }}
            .total {{ margin-top: 20px; font-size: 18px; font-weight: bold; text-align: left; }}
        </style>
    </head>
    <body>
        <h1>تقرير حسابات حصاد المجدول</h1>
        <table>
            <thead>
                <tr>
                    <th>التاريخ</th>
                    <th>عدد العمال</th>
                    <th>أجرة العامل</th>
                    <th>المصاريف</th>
                    <th>المجموع</th>
                    <th>ملاحظات</th>
                </tr>
            </thead>
            <tbody>
                {table_rows}
            </tbody>
        </table>
        <div class="total">المجموع الكلي: {total_cost_all:.2f} دينار أردني</div>
    </body>
    </html>
    """
    pdf_bytes = HTML(string=html_content).write_pdf()
    return pdf_bytes


def generate_excel_report(date_from: str = None, date_to: str = None):
    rows = query_records(date_from, date_to)

    wb = Workbook()
    ws = wb.active
    ws.title = "اليوميات"

    ws.append(["التاريخ", "عدد العمال", "أجرة العامل", "المصاريف", "المجموع", "ملاحظات"])
    for row in rows:
        date, count, wage, exp, notes = row
        total = (count * wage) + exp
        ws.append([date, count, wage, exp, total, notes or ""])

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


# ---------------------------------------------------------------------------
# Webhook
# ---------------------------------------------------------------------------

@app.post("/tg-webhook")
async def telegram_webhook(request: Request):
    secret_header = request.headers.get("X-Telegram-Bot-Api-Secret-Token")
    if TELEGRAM_WEBHOOK_SECRET and secret_header != TELEGRAM_WEBHOOK_SECRET:
        return Response(status_code=403)

    data = await request.json()
    if "message" in data:
        chat_id = data["message"]["chat"]["id"]
        text = data["message"].get("text", "")

        if text.startswith("/start"):
            await send_telegram_message(
                chat_id,
                "أهلاً بك في بوت إدارة حسابات المزرعة!\n\n"
                "📝 لتسجيل يومية: «اشتغل 23 عامل 6 ساعات بدينار ونص»\n"
                "✏️ للتعديل: «أضف أجرة نقل 25 دينار» أو «عدّل أجرة امس تصير 11»\n"
                "🗑️ للحذف: «احذف يومية امس»\n"
                "📊 للتقرير: «تقرير امس» أو «كل السجلات pdf»"
            )
            return {"status": "ok"}

        result = await analyze_message(text)
        intent = result.get("intent", "other")

        # -------------------- تسجيل بيانات جديدة --------------------
        if intent == "record":
            confidence = result.get("confidence", "none")

            if confidence == "partial":
                missing = result.get("missing_info", "بعض البيانات غير واضحة")
                w_count = result.get("workers_count", 0) or 0
                wage = result.get("wage_per_worker", 0) or 0
                await send_telegram_message(
                    chat_id,
                    f"البيانات ناقصة: {missing}\n"
                    f"(ما فهمته: عدد العمال = {w_count}, الأجرة = {wage})\n"
                    f"يرجى إعادة الإرسال مع التفاصيل الكاملة."
                )
            elif confidence == "full":
                w_count = result.get("workers_count", 0) or 0
                wage = float(result.get("wage_per_worker", 0.0) or 0.0)
                exp = float(result.get("expenses", 0.0) or 0.0)
                notes = result.get("notes", "")
                record_date = result.get("date") or datetime.now().strftime("%Y-%m-%d")

                conn = sqlite3.connect(DB_NAME)
                cursor = conn.cursor()
                cursor.execute(
                    "INSERT INTO daily_records (date, workers_count, wage_per_worker, expenses, notes) VALUES (?, ?, ?, ?, ?)",
                    (record_date, w_count, wage, exp, notes)
                )
                conn.commit()
                conn.close()

                total = (w_count * wage) + exp
                response_msg = (
                    f"✅ تم تسجيل اليومية بنجاح! ({record_date})\n"
                    f"- عدد العمال: {w_count}\n"
                    f"- أجرة العامل اليومية: {wage:.2f} د.أ\n"
                    f"- المصاريف: {exp:.2f} د.أ\n"
                    f"- المجموع الكلي: {total:.2f} د.أ"
                )
                if notes:
                    response_msg += f"\n\nملاحظات: {notes}"
                await send_telegram_message(chat_id, response_msg)
            else:
                await send_telegram_message(
                    chat_id,
                    "لم أفهم أي بيانات يومية بهذه الرسالة.\n"
                    "اكتب مثلاً: 10 عمال اجرة 12 دينار ومصاريف 5 دينار."
                )

        # -------------------- تعديل سجل موجود --------------------
        elif intent == "update":
            target_date = result.get("update_target_date") or "last"
            action = result.get("update_action")
            value = float(result.get("update_value", 0) or 0)
            note = result.get("update_note", "")

            record = find_target_record(target_date)

            if not record:
                await send_telegram_message(
                    chat_id,
                    "لم أجد سجلاً مطابقاً لأعدّل عليه.\n"
                    "تأكد من التاريخ، أو سجّل يومية جديدة أولاً."
                )
            elif not action:
                await send_telegram_message(
                    chat_id,
                    "فهمت أنك تريد تعديل سجل، لكن لم أفهم نوع التعديل بالضبط.\n"
                    "جرب مثلاً: «أضف 10 دينار مصاريف» أو «غيّر عدد العمال ليصير 20»."
                )
            else:
                confirmation = apply_update(record, action, value, note)
                await send_telegram_message(chat_id, confirmation)

        # -------------------- طلب تقرير --------------------
        elif intent == "report":
            scope = result.get("report_scope", "all")
            date_from = result.get("report_date_from")
            date_to = result.get("report_date_to")
            fmt = result.get("report_format") or "text"

            if scope == "single_day" and date_from:
                title = f"📋 تقرير يوم {date_from}"
            elif scope == "range" and date_from and date_to:
                title = f"📋 تقرير من {date_from} إلى {date_to}"
            else:
                title = "📋 تقرير كامل لكل السجلات"
                date_from = date_to = None

            if fmt == "pdf":
                pdf_data = generate_pdf_report(date_from, date_to)
                await send_telegram_document(chat_id, "report.pdf", pdf_data, title)
            elif fmt == "excel":
                excel_data = generate_excel_report(date_from, date_to)
                await send_telegram_document(chat_id, "report.xlsx", excel_data, title)
            else:
                rows = query_records(date_from, date_to)
                report_text = format_text_report(rows, title)
                await send_telegram_message(chat_id, report_text)

        # -------------------- غير مفهوم --------------------
        else:
            await send_telegram_message(
                chat_id,
                "يمكنني مساعدتك بتسجيل يوميات المزرعة، تعديلها، أو استخراج تقارير عنها.\n"
                "جرب مثلاً: «10 عمال اجرة 12 دينار» أو «أضف 5 دينار مصاريف» أو «تقرير امس»."
            )

    return {"status": "ok"}
