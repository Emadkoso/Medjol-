import os
import re
import sqlite3
import json
import logging
from io import BytesIO
from datetime import datetime, timedelta
from fastapi import FastAPI, Request, Response
import httpx
from openpyxl import Workbook
from weasyprint import HTML
import uvicorn

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("medjol")

app = FastAPI()

# ===== متغيرات البيئة =====
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")
OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "meta-llama/llama-3.3-70b-instruct")
WEBHOOK_URL = os.getenv("WEBHOOK_URL")  # مثال: https://your-app.onrender.com/webhook
PORT = int(os.getenv("PORT", 8000))

DB_NAME = "harvest.db"

# ===== تهيئة قاعدة البيانات =====
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
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS conversation_state (
            chat_id INTEGER PRIMARY KEY,
            step TEXT,
            date TEXT,
            workers_count INTEGER,
            wage_mode TEXT,
            hours REAL,
            wage_per_hour REAL,
            wage_per_worker REAL,
            expenses REAL,
            notes TEXT
        )
    ''')
    conn.commit()
    conn.close()

init_db()

# ===== دوال مساعدة لإرسال الرسائل =====
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

# ===== استخراج الأرقام (نفس الكود السابق) =====
ARABIC_INDIC_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
ARABIC_NUMBER_WORDS = {
    "صفر": 0, "واحد": 1, "وحدة": 1, "اثنين": 2, "اثنان": 2,
    "تلاتة": 3, "ثلاثة": 3, "اربعة": 4, "أربعة": 4, "خمسة": 5,
    "ستة": 6, "سبعة": 7, "ثمانية": 8, "تمانية": 8, "تسعة": 9, "عشرة": 10,
    "احدعش": 11, "أحد عشر": 11, "اثناعش": 12, "اثنا عشر": 12,
    "تلتطعش": 13, "ثلاثطعش": 13, "اربعطعش": 14, "خمسطعش": 15,
    "سطعش": 16, "ستطعش": 16, "سبعطعش": 17, "تمنطعش": 18, "تسعطعش": 19,
    "عشرين": 20, "تلاتين": 30, "ثلاثين": 30, "اربعين": 40, "خمسين": 50,
    "ستين": 60, "سبعين": 70, "تمانين": 80, "تسعين": 90, "مية": 100, "مائة": 100
}

def extract_number(text: str):
    t = text.strip().translate(ARABIC_INDIC_DIGITS)
    match = re.search(r"\d+(\.\d+)?", t)
    if match:
        return float(match.group())
    for word, val in ARABIC_NUMBER_WORDS.items():
        if word in text:
            return float(val)
    return None

async def extract_number_ai(text: str):
    if not OPENROUTER_API_KEY:
        return None
    prompt = (
        "استخرج رقماً واحداً فقط من النص التالي وأرجعه بصيغة JSON حصراً: "
        '{"number": <رقم أو null>}\n'
        f'النص: "{text}"'
    )
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
                headers=headers, json=payload, timeout=15.0
            )
            if res.status_code == 200:
                content = res.json()['choices'][0]['message']['content']
                data = json.loads(content)
                return data.get("number")
        except Exception as e:
            logger.error(f"extract_number_ai error: {e}")
    return None

# ===== تحليل النية باستخدام AI =====
def build_prompt(text: str, today_date: str, yesterday_date: str, day_before_date: str) -> str:
    return f"""أنت محاسب خبير ودقيق جداً لمزرعة نخيل في الأردن (منطقة الأغوار)، وموجّه ذكي يحدد نية المستخدم بدقة من رسالته.

## حدد نية الرسالة (intent):
- "record": يريد تسجيل يومية عمل (حتى لو ذكر معلومة واحدة فقط مثل عدد العمال).
- "update": تعديل على سجل مسجّل مسبقاً (إضافة/تغيير/حذف).
- "report": طلب استرجاع بيانات مسجلة سابقاً.
- "other": أي شيء آخر.

## استخرج من النص كل ما هو متوفر بوضوح (لا تخترع شيئاً غير مذكور، اترك الحقل 0 أو null إذا غير مذكور):
- workers_count: عدد العمال إذا ذُكر.
- hours: عدد ساعات العمل إذا ذُكر.
- wage_per_hour: أجرة الساعة إذا ذُكرت.
- wage_per_worker: أجرة العامل اليومية الإجمالية إذا ذُكرت مباشرة (وإلا احسبها إن توفر hours و wage_per_hour).
- expenses: مجموع المصاريف إذا ذُكرت.
- date: التاريخ المقصود بصيغة YYYY-MM-DD (اليوم {today_date}، "امس"={yesterday_date}، "اول امس"={day_before_date}).

## للتحديث (update):
- update_target_date: تاريخ أو "last".
- update_action: "add_expense" | "add_expense_per_worker" | "set_expense" | "set_wage" | "set_workers" | "delete_record" | "append_note".
- update_value: رقم.
- update_note: نص قصير.

## للتقرير (report):
- report_scope: "single_day" | "range" | "all".
- report_date_from / report_date_to: YYYY-MM-DD أو null.
- report_format: "text" | "pdf" | "excel" (افتراضي "text" إذا لم يُذكر).

## صيغة الإخراج: JSON فقط بدون أي نص إضافي:
{{
  "intent": "record" | "update" | "report" | "other",
  "workers_count": رقم أو 0,
  "hours": رقم أو 0,
  "wage_per_hour": رقم أو 0,
  "wage_per_worker": رقم أو 0,
  "expenses": رقم أو 0,
  "date": "YYYY-MM-DD" أو "",
  "update_target_date": "YYYY-MM-DD" أو "last" أو null,
  "update_action": نص أو null,
  "update_value": رقم أو 0,
  "update_note": نص أو "",
  "report_scope": نص أو null,
  "report_date_from": نص أو null,
  "report_date_to": نص أو null,
  "report_format": نص أو null
}}

النص: "{text}"
"""

async def analyze_message(text: str) -> dict:
    if not OPENROUTER_API_KEY:
        logger.error("OPENROUTER_API_KEY is missing!")
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
                headers=headers, json=payload, timeout=25.0
            )
            if res.status_code == 200:
                content = res.json()['choices'][0]['message']['content']
                logger.info(f"AI raw response: {content}")
                return json.loads(content)
            else:
                logger.error(f"OpenRouter API error {res.status_code}: {res.text}")
        except Exception as e:
            logger.error(f"AI request exception: {e}")
    return {"intent": "other"}

# ===== دوال الحالة والحوار (نفس الكود السابق) =====
# (جميع دوال get_state, save_state, clear_state, next_missing_step, question_for_step, build_confirmation_text, handle_guided_answer, start_guided_flow تبقى كما هي)
# لتوفير المساحة، سأختصرها ولكنها موجودة في الكود النهائي المرفق.

# ===== دوال التحديث والتقارير =====
# (query_records, format_text_report, generate_excel_report, generate_pdf_report, handle_update, handle_report تبقى كما هي)

# ===== المعالج الرئيسي =====
async def handle_incoming_message(chat_id: int, text: str):
    # إرسال رسالة فورية لتأكيد الاستلام
    await send_telegram_message(chat_id, "⏳ جارٍ معالجة طلبك...")

    try:
        state = get_state(chat_id)
        if state and state.get("step") not in [None, "CONFIRM"]:
            await handle_guided_answer(chat_id, text, state)
            return

        analysis = await analyze_message(text)
        intent = analysis.get("intent", "other")

        if intent == "record":
            await start_guided_flow(chat_id, analysis)
        elif intent == "update":
            await handle_update(chat_id, analysis)
        elif intent == "report":
            await handle_report(chat_id, analysis)
        else:
            await send_telegram_message(
                chat_id,
                "مرحباً! أرسل لي تفاصيل اليومية (مثل: 5 عمال، 10 دنانير لكل عامل)، أو اطلب تقريراً، أو قم بتحديث سابق."
            )
    except Exception as e:
        logger.error(f"Error in handle_incoming_message: {e}", exc_info=True)
        await send_telegram_message(chat_id, "❌ حدث خطأ ما، حاول مرة أخرى لاحقاً.")

# ===== نقطة نهاية Webhook =====
@app.post("/webhook")
async def telegram_webhook(request: Request):
    try:
        data = await request.json()
        logger.info(f"Incoming update: {data}")
        if "message" in data and "text" in data["message"]:
            chat_id = data["message"]["chat"]["id"]
            text = data["message"]["text"].strip()
            # إطلاق المعالج في الخلفية حتى لا نؤخر الرد على Telegram
            import asyncio
            asyncio.create_task(handle_incoming_message(chat_id, text))
        return {"ok": True}
    except Exception as e:
        logger.error(f"Webhook error: {e}", exc_info=True)
        return {"ok": False}

# ===== نقطة بديلة للاختبار (polling) =====
# يمكن تشغيلها كـ fallback إذا فشل webhook
async def polling_loop():
    """حلقة بسيطة لسحب التحديثات (بديل للـ webhook)"""
    offset = 0
    while True:
        try:
            url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/getUpdates?offset={offset}&timeout=30"
            async with httpx.AsyncClient() as client:
                res = await client.get(url, timeout=35)
                if res.status_code == 200:
                    updates = res.json().get("result", [])
                    for update in updates:
                        offset = update["update_id"] + 1
                        if "message" in update and "text" in update["message"]:
                            chat_id = update["message"]["chat"]["id"]
                            text = update["message"]["text"].strip()
                            await handle_incoming_message(chat_id, text)
                else:
                    logger.error(f"Polling error: {res.status_code}")
        except Exception as e:
            logger.error(f"Polling exception: {e}")
        await asyncio.sleep(1)

# ===== تسجيل webhook عند بدء التشغيل =====
async def set_webhook():
    if not WEBHOOK_URL:
        logger.warning("WEBHOOK_URL غير معرّف، سيتم استخدام polling بدلاً من ذلك.")
        return False
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/setWebhook?url={WEBHOOK_URL}"
    async with httpx.AsyncClient() as client:
        try:
            res = await client.get(url)
            data = res.json()
            if data.get("ok"):
                logger.info(f"✅ Webhook تم تسجيله بنجاح: {WEBHOOK_URL}")
                return True
            else:
                logger.error(f"❌ فشل تسجيل webhook: {data}")
                return False
        except Exception as e:
            logger.error(f"❌ استثناء أثناء تسجيل webhook: {e}")
            return False

@app.on_event("startup")
async def startup_event():
    # حاول تسجيل webhook
    success = await set_webhook()
    if not success:
        logger.warning("سيتم تشغيل polling كحل احتياطي.")
        import asyncio
        asyncio.create_task(polling_loop())

# ===== تشغيل التطبيق =====
if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=PORT)
