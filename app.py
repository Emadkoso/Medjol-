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
# AI parsing (OpenRouter) — برومبت موسّع بأسلوب استدلال دقيق
# ---------------------------------------------------------------------------

def build_prompt(text: str, today_date: str, yesterday_date: str, day_before_date: str) -> str:
    return f"""أنت محاسب خبير ودقيق جداً، متخصص بحسابات مزارع النخيل في الأردن (منطقة الأغوار). أسلوبك في التفكير منهجي: تقرأ النص كاملاً، تحدد كل رقم ومعناه، تتحقق من الحساب خطوة بخطوة قبل ما ترجع النتيجة، ولا تخمّن أبداً معلومة غير مذكورة.

## منهجية التفكير المطلوبة (طبّقها داخلياً قبل الإخراج، دون كتابتها):
1. اقرأ النص وحدد: كم مجموعة عمال؟ كل مجموعة كم عدد وبأي أجرة؟
2. حدد كل بند مصاريف مذكور صراحة (بنزين، أكل، نقل، أدوات...).
3. حدد التاريخ المقصود.
4. احسب الأرقام النهائية بدقة رياضية، وتحقق من صحة الجمع والضرب قبل الإخراج.
5. قيّم مدى اكتمال المعلومات قبل ما تقرر الـ confidence.

## قواعد الحساب:
- إذا أُعطي أجر بالساعة: الأجرة اليومية = عدد الساعات × سعر الساعة.
- إذا وُجدت أكثر من مجموعة عمال بأجور مختلفة: اجمع كل العمال بـ workers_count، واحسب wage_per_worker كمتوسط مرجّح = (مجموع [عدد×أجرة] لكل مجموعة) ÷ (مجموع العمال). اذكر التفصيل الكامل بالـ notes.
- إذا وُجدت مصاريف متعددة: اجمعها كلها برقم واحد بـ expenses، واذكر تفصيل كل بند بالـ notes.
- الأرقام المكتوبة بالحروف العربية أو العامية (خمسطعش=15، عشرين=20، تلاتين=30...) حوّلها لأرقام.
- تجاهل أي كلام غير مرتبط بالحسابات (تحيات، أسئلة عامة، دعاء...).

## قواعد التاريخ (اليوم هو {today_date}):
- بدون ذكر تاريخ → استخدم {today_date}
- "امس" أو "أمس" → {yesterday_date}
- "اول امس" أو "أول أمس" أو "قبل امس" → {day_before_date}
- تاريخ محدد مذكور نصاً → حوّله لصيغة YYYY-MM-DD

## تقييم confidence (كن صارماً وحذراً هنا، هذا أهم جزء):
- "full": ذُكر عدد العمال وأجرتهم (أو ما يكفي لحسابها) بوضوح تام.
- "partial": ذُكرت بعض الأرقام لكن ينقص عنصر أساسي (مثلاً عدد عمال بدون أي إشارة للأجرة، أو أجرة بدون عدد عمال). في هذه الحالة أرجع الأرقام المتوفرة فقط ولا تخترع الناقص، واملأ missing_info بدقة.
- "none": النص لا يحتوي أي بيانات يومية مالية إطلاقاً (سؤال، تحية، طلب تقرير...).
عند أي شك حقيقي بين full وpartial، اختر partial — الدقة أهم من اكتمال الشكل.

## صيغة الإخراج: JSON فقط، بدون أي نص أو شرح أو Markdown قبله أو بعده:
{{
  "workers_count": عدد صحيح,
  "wage_per_worker": رقم,
  "expenses": رقم,
  "date": "YYYY-MM-DD",
  "notes": "شرح موجز لأي حساب غير مباشر (متوسطات، بنود مصاريف)",
  "confidence": "full" | "partial" | "none",
  "missing_info": "وصف قصير لما هو ناقص، أو نص فارغ إذا لا شيء ناقص"
}}

## أمثلة توضيحية:

النص: "اشتغل معي 23 عامل لمدة 6 ساعات اجرة الساعة الواحدة دينار ونصف"
التفكير: 23 عامل، 6 ساعات × 1.5 د = 9 د للعامل، لا مصاريف مذكورة، لا تاريخ مذكور → اليوم.
الإخراج: {{"workers_count": 23, "wage_per_worker": 9.0, "expenses": 0, "date": "{today_date}", "notes": "6 ساعات × 1.5 دينار/ساعة = 9 دينار للعامل", "confidence": "full", "missing_info": ""}}

النص: "امس اشتغلوا 10 عمال باجرة 12 دينار و5 عمال باجرة 15 دينار، وصرفنا 20 دينار بنزين و10 دينار أكل"
التفكير: مجموع العمال = 15. المتوسط المرجّح = (10×12 + 5×15)/15 = (120+75)/15 = 13. المصاريف = 20+10 = 30. التاريخ = أمس.
الإخراج: {{"workers_count": 15, "wage_per_worker": 13.0, "expenses": 30, "date": "{yesterday_date}", "notes": "10 عمال بـ12 دينار + 5 عمال بـ15 دينار = متوسط 13 دينار للعامل. المصاريف: 20 بنزين + 10 أكل", "confidence": "full", "missing_info": ""}}

النص: "اشتغل اليوم خمسطعش عامل"
التفكير: عدد العمال معروف (15) لكن الأجرة غير مذكورة إطلاقاً → partial.
الإخراج: {{"workers_count": 15, "wage_per_worker": 0, "expenses": 0, "date": "{today_date}", "notes": "", "confidence": "partial", "missing_info": "لم تُذكر أجرة العامل"}}

النص: "شو الجو اليوم؟"
التفكير: لا توجد أي بيانات مالية أو عمالية.
الإخراج: {{"workers_count": 0, "wage_per_worker": 0, "expenses": 0, "date": "{today_date}", "notes": "", "confidence": "none", "missing_info": "لا توجد بيانات يومية بالنص"}}

النص الحالي المطلوب تحليله: "{text}"
"""


async def parse_with_ai(text: str) -> dict:
    if not OPENROUTER_API_KEY:
        logger.error("OPENROUTER_API_KEY is missing from environment variables!")
        return {}

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
    return {}


# ---------------------------------------------------------------------------
# Reports
# ---------------------------------------------------------------------------

def generate_pdf_report():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT date, workers_count, wage_per_worker, expenses, notes FROM daily_records ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()

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


def generate_excel_report():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT date, workers_count, wage_per_worker, expenses, notes FROM daily_records ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()

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
                "أهلاً بك في بوت إدارة حسابات المزرعة!\n"
                "يمكنك كتابة اليوميات بأي طريقة، مثال:\n"
                "- اشتغل 23 عامل 6 ساعات الساعة بدينار ونص\n"
                "- امس 10 عمال باجرة 12 دينار و5 عمال باجرة 15 دينار، ومصاريف 30 دينار\n"
                "أو اطلب 'تقرير PDF' أو 'تقرير اكسل'."
            )

        elif "pdf" in text.lower() or "تقرير" in text and "اكسل" not in text.lower():
            pdf_data = generate_pdf_report()
            await send_telegram_document(chat_id, "report.pdf", pdf_data, "إليك تقرير الحسابات بصيغة PDF")

        elif "excel" in text.lower() or "اكسل" in text:
            excel_data = generate_excel_report()
            await send_telegram_document(chat_id, "report.xlsx", excel_data, "إليك تقرير الحسابات بصيغة Excel")

        else:
            parsed = await parse_with_ai(text)
            confidence = parsed.get("confidence", "none")

            if confidence == "none":
                await send_telegram_message(
                    chat_id,
                    "لم أفهم أي بيانات يومية بهذه الرسالة.\n"
                    "اكتب مثلاً: 10 عمال اجرة 12 دينار ومصاريف 5 دينار."
                )

            elif confidence == "partial":
                missing = parsed.get("missing_info", "بعض البيانات غير واضحة")
                w_count = parsed.get("workers_count", 0) or 0
                wage = parsed.get("wage_per_worker", 0) or 0
                await send_telegram_message(
                    chat_id,
                    f"البيانات ناقصة: {missing}\n"
                    f"(ما فهمته حتى الآن: عدد العمال = {w_count}, الأجرة = {wage})\n"
                    f"يرجى إعادة الإرسال مع التفاصيل الكاملة."
                )

            else:  # full
                w_count = parsed.get("workers_count", 0) or 0
                wage = float(parsed.get("wage_per_worker", 0.0) or 0.0)
                exp = float(parsed.get("expenses", 0.0) or 0.0)
                notes = parsed.get("notes", "")
                record_date = parsed.get("date") or datetime.now().strftime("%Y-%m-%d")

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
                    f"تم تسجيل اليومية بنجاح! ({record_date})\n"
                    f"- عدد العمال: {w_count}\n"
                    f"- أجرة العامل اليومية: {wage:.2f} د.أ\n"
                    f"- المصاريف: {exp:.2f} د.أ\n"
                    f"- المجموع الكلي: {total:.2f} د.أ"
                )
                if notes:
                    response_msg += f"\n\nملاحظات: {notes}"
                await send_telegram_message(chat_id, response_msg)

    return {"status": "ok"}
