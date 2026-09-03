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
# Number extraction — أرقام، أرقام عربية، وكلمات عامية شائعة
# ---------------------------------------------------------------------------

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
    """احتياطي ذكي: إذا فشل الاستخراج المباشر، نسأل الموديل يستخرج رقم واحد فقط."""
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


# ---------------------------------------------------------------------------
# AI Router — تحديد نية الرسالة عند عدم وجود محادثة نشطة
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Conversation state (الحوار التفاعلي خطوة بخطوة)
# ---------------------------------------------------------------------------

def get_state(chat_id: int):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM conversation_state WHERE chat_id = ?", (chat_id,))
    row = cursor.fetchone()
    conn.close()
    if not row:
        return None
    cols = ["chat_id", "step", "date", "workers_count", "wage_mode",
            "hours", "wage_per_hour", "wage_per_worker", "expenses", "notes"]
    return dict(zip(cols, row))


def save_state(chat_id: int, **fields):
    existing = get_state(chat_id)
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    if existing:
        merged = {**existing, **fields}
        cursor.execute('''
            UPDATE conversation_state SET step=?, date=?, workers_count=?, wage_mode=?,
                hours=?, wage_per_hour=?, wage_per_worker=?, expenses=?, notes=?
            WHERE chat_id=?
        ''', (merged["step"], merged["date"], merged["workers_count"], merged["wage_mode"],
              merged["hours"], merged["wage_per_hour"], merged["wage_per_worker"],
              merged["expenses"], merged["notes"], chat_id))
    else:
        defaults = {
            "step": None, "date": datetime.now().strftime("%Y-%m-%d"),
            "workers_count": None, "wage_mode": None, "hours": None,
            "wage_per_hour": None, "wage_per_worker": None, "expenses": None, "notes": ""
        }
        merged = {**defaults, **fields}
        cursor.execute('''
            INSERT INTO conversation_state
            (chat_id, step, date, workers_count, wage_mode, hours, wage_per_hour, wage_per_worker, expenses, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (chat_id, merged["step"], merged["date"], merged["workers_count"], merged["wage_mode"],
              merged["hours"], merged["wage_per_hour"], merged["wage_per_worker"],
              merged["expenses"], merged["notes"]))
    conn.commit()
    conn.close()


def clear_state(chat_id: int):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM conversation_state WHERE chat_id = ?", (chat_id,))
    conn.commit()
    conn.close()


def next_missing_step(state: dict) -> str:
    """يحدد أول سؤال ناقص بذكاء — يتخطى ما هو معروف مسبقاً."""
    if not state.get("workers_count"):
        return "ASK_WORKERS"
    if not state.get("wage_per_worker"):
        if not state.get("wage_mode"):
            return "ASK_WAGE_MODE"
        if state["wage_mode"] == "hourly":
            if not state.get("hours"):
                return "ASK_HOURS"
            if not state.get("wage_per_hour"):
                return "ASK_WAGE_PER_HOUR"
        else:
            return "ASK_WAGE_DAILY"
    if state.get("expenses") is None:
        return "ASK_EXPENSES"
    return "CONFIRM"


def question_for_step(step: str) -> str:
    return {
        "ASK_WORKERS": "🧑‍🌾 كم عدد العمال الذين اشتغلوا معك؟",
        "ASK_WAGE_MODE": "💵 كيف تحدد الأجرة؟\n1️⃣ أجرة يومية مباشرة للعامل\n2️⃣ بالساعة (عدد الساعات × سعر الساعة)\n\nاكتب 1 أو 2.",
        "ASK_HOURS": "⏱️ كم عدد ساعات العمل؟",
        "ASK_WAGE_PER_HOUR": "💰 كم أجرة الساعة الواحدة (بالدينار)؟",
        "ASK_WAGE_DAILY": "💰 كم أجرة العامل الواحد لهذا اليوم (بالدينار)؟",
        "ASK_EXPENSES": "🧾 هل هناك مصاريف إضافية (بنزين، أكل، نقل...)؟\nاكتب المجموع، أو اكتب 'لا' إذا لا يوجد.",
    }.get(step, "")


def build_confirmation_text(state: dict) -> str:
    wage = state.get("wage_per_worker") or 0
    workers = state.get("workers_count") or 0
    expenses = state.get("expenses") or 0
    total = (workers * wage) + expenses
    return (
        f"📋 راجع البيانات قبل الحفظ ({state.get('date')}):\n"
        f"- عدد العمال: {workers}\n"
        f"- أجرة العامل: {wage:.2f} د.أ\n"
        f"- المصاريف: {expenses:.2f} د.أ\n"
        f"- المجموع الكلي: {total:.2f} د.أ\n\n"
        f"✅ اكتب 'نعم' للحفظ، أو ❌ 'الغاء' للإلغاء."
    )


NO_EXPENSE_WORDS = ["لا", "ﻻ", "لا يوجد", "ماكو", "ولا شي", "بدون", "no", "0"]
YES_WORDS = ["نعم", "ايوه", "ايه", "تمام", "صح", "اوك", "ok", "yes", "احفظ", "أكيد"]
CANCEL_WORDS = ["الغاء", "إلغاء", "كنسل", "cancel", "لغي", "الغي"]


async def handle_guided_answer(chat_id: int, text: str, state: dict):
    step = state["step"]
    stripped = text.strip()

    if step == "ASK_WORKERS":
        num = extract_number(stripped) or await extract_number_ai(stripped)
        if num is None:
            await send_telegram_message(chat_id, "لم أفهم الرقم 🙏 كم عدد العمال؟ (اكتب رقماً فقط)")
            return
        state["workers_count"] = int(num)

    elif step == "ASK_WAGE_MODE":
        if "2" in stripped or "ساعة" in stripped:
            state["wage_mode"] = "hourly"
        elif "1" in stripped or "يومي" in stripped or "مباشر" in stripped:
            state["wage_mode"] = "daily"
        else:
            await send_telegram_message(chat_id, "من فضلك اكتب 1 (أجرة يومية مباشرة) أو 2 (بالساعة).")
            return

    elif step == "ASK_HOURS":
        num = extract_number(stripped) or await extract_number_ai(stripped)
        if num is None:
            await send_telegram_message(chat_id, "لم أفهم الرقم 🙏 كم عدد ساعات العمل؟")
            return
        state["hours"] = num

    elif step == "ASK_WAGE_PER_HOUR":
        num = extract_number(stripped) or await extract_number_ai(stripped)
        if num is None:
            await send_telegram_message(chat_id, "لم أفهم الرقم 🙏 كم أجرة الساعة الواحدة؟")
            return
        state["wage_per_hour"] = num
        state["wage_per_worker"] = (state.get("hours") or 0) * num
        state["notes"] = (state.get("notes") or "") + f" {state['hours']}س×{num}د/س={state['wage_per_worker']:.2f}د للعامل"

    elif step == "ASK_WAGE_DAILY":
        num = extract_number(stripped) or await extract_number_ai(stripped)
        if num is None:
            await send_telegram_message(chat_id, "لم أفهم الرقم 🙏 كم أجرة العامل اليومية؟")
            return
        state["wage_per_worker"] = num

    elif step == "ASK_EXPENSES":
        if any(w in stripped for w in NO_EXPENSE_WORDS) and extract_number(stripped) is None:
            state["expenses"] = 0.0
        else:
            num = extract_number(stripped) or await extract_number_ai(stripped)
            if num is None:
                await send_telegram_message(chat_id, "لم أفهم 🙏 اكتب مبلغ المصاريف كرقم، أو 'لا' إذا لا يوجد.")
                return
            state["expenses"] = num

    elif step == "CONFIRM":
        if any(w in stripped for w in YES_WORDS):
            workers = state.get("workers_count") or 0
            wage = state.get("wage_per_worker") or 0
            expenses = state.get("expenses") or 0
            notes = (state.get("notes") or "").strip()
            record_date = state.get("date") or datetime.now().strftime("%Y-%m-%d")

            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO daily_records (date, workers_count, wage_per_worker, expenses, notes) VALUES (?, ?, ?, ?, ?)",
                (record_date, workers, wage, expenses, notes)
            )
            conn.commit()
            conn.close()
            clear_state(chat_id)

            total = (workers * wage) + expenses
            await send_telegram_message(
                chat_id,
                f"✅ تم حفظ اليومية بنجاح! ({record_date})\n"
                f"- عدد العمال: {workers}\n"
                f"- أجرة العامل: {wage:.2f} د.أ\n"
                f"- المصاريف: {expenses:.2f} د.أ\n"
                f"- المجموع الكلي: {total:.2f} د.أ"
            )
        elif any(w in stripped for w in CANCEL_WORDS):
            clear_state(chat_id)
            await send_telegram_message(chat_id, "❌ تم إلغاء العملية، لم يُحفظ أي شيء.")
        else:
            await send_telegram_message(chat_id, "من فضلك اكتب 'نعم' للحفظ أو 'الغاء' للإلغاء.")
        return

    # تحديد السؤال التالي بذكاء وحفظ الحالة
    step_after = next_missing_step(state)
    save_state(
        chat_id, step=step_after, date=state.get("date"),
        workers_count=state.get("workers_count"), wage_mode=state.get("wage_mode"),
        hours=state.get("hours"), wage_per_hour=state.get("wage_per_hour"),
        wage_per_worker=state.get("wage_per_worker"), expenses=state.get("expenses"),
        notes=state.get("notes")
    )

    if step_after == "CONFIRM":
        await send_telegram_message(chat_id, build_confirmation_text(state))
    else:
        await send_telegram_message(chat_id, question_for_step(step_after))


async def start_guided_flow(chat_id: int, prefill: dict):
    """يبدأ حواراً جديداً، ويتخطى الأسئلة التي عرف إجابتها من الرسالة الأصلية."""
    workers = int(prefill.get("workers_count") or 0) or None
    hours = prefill.get("hours") or None
    wage_per_hour = prefill.get("wage_per_hour") or None
    wage_per_worker = prefill.get("wage_per_worker") or None
    expenses = prefill.get("expenses")
    date_val = prefill.get("date") or datetime.now().strftime("%Y-%m-%d")

    wage_mode = None
    if wage_per_worker:
        wage_mode = "daily"  # موجود مسبقاً، ما رح نحتاج نسأل عن الطريقة
    elif hours and wage_per_hour:
        wage_per_worker = hours * wage_per_hour
        wage_mode = "hourly"

    state = {
        "date": date_val, "workers_count": workers, "wage_mode": wage_mode,
        "hours": hours, "wage_per_hour": wage_per_hour,
        "wage_per_worker": wage_per_worker,
        "expenses": expenses if expenses not in (None, 0) else (0.0 if expenses == 0 else None),
        "notes": ""
    }

    step = next_missing_step(state)
    save_state(chat_id, step=step, **state)

    if step == "CONFIRM":
        await send_telegram_message(chat_id, build_confirmation_text(state))
    else:
        intro = "تمام، خلينا نكمل التفاصيل 👇\n\n" if workers or wage_per_worker else ""
        await send_telegram_message(chat_id, intro + question_for_step(step))


# ---------------------------------------------------------------------------
# Update / Report helpers
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
    """توليد تقرير نصي من الصفوف."""
    if not rows:
        return f"{title}\n\nلا توجد أي سجلات لهذه الفترة."
    lines = [title, ""]
    grand_total = 0.0
    for date, count, wage, exp, notes in rows:
        tota
