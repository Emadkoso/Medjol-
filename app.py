# app.py - نسخة نهائية تعتمد على Polling فقط (بدون Webhook)
# انسخ هذا الكود بالكامل وألصقه في ملف app.py على Render

import os
import re
import sqlite3
import json
import logging
import asyncio
from datetime import datetime, timedelta
from fastapi import FastAPI
import httpx
from openpyxl import Workbook
from weasyprint import HTML
import uvicorn

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("medjol")

app = FastAPI()

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")
OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "meta-llama/llama-3.3-70b-instruct")
PORT = int(os.getenv("PORT", 8000))
DB_NAME = "harvest.db"

# ========== قاعدة البيانات ==========
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

# ========== إرسال رسائل تلغرام ==========
async def send_telegram_message(chat_id: int, text: str):
    if not TELEGRAM_BOT_TOKEN:
        return
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    async with httpx.AsyncClient() as client:
        try:
            await client.post(url, json={"chat_id": chat_id, "text": text}, timeout=10)
        except Exception as e:
            logger.error(f"send error: {e}")

async def send_telegram_document(chat_id: int, filename: str, content: bytes, caption: str = ""):
    if not TELEGRAM_BOT_TOKEN:
        return
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
    files = {"document": (filename, content)}
    data = {"chat_id": chat_id, "caption": caption}
    async with httpx.AsyncClient() as client:
        try:
            await client.post(url, data=data, files=files, timeout=20)
        except Exception as e:
            logger.error(f"doc error: {e}")

# ========== استخراج الأرقام ==========
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
        "استخرج رقماً واحداً فقط من النص التالي وأرجعه بصيغة JSON: "
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
                headers=headers, json=payload, timeout=15
            )
            if res.status_code == 200:
                data = res.json()
                content = data['choices'][0]['message']['content']
                return json.loads(content).get("number")
        except:
            pass
    return None

# ========== تحليل النية ==========
def build_prompt(text: str, today: str, yesterday: str, day_before: str):
    return f"""أنت محاسب خبير لمزرعة نخيل في الأردن. حدد نية المستخدم.

الأوامر:
- "record": تسجيل يومية.
- "update": تعديل سجل.
- "report": تقرير.
- "other": أي شيء آخر.

استخرج البيانات المذكورة فقط:
- workers_count, hours, wage_per_hour, wage_per_worker, expenses, date (YYYY-MM-DD)
- للتحديث: update_target_date ("last" أو تاريخ), update_action, update_value, update_note
- للتقرير: report_scope ("single_day"/"range"/"all"), report_date_from, report_date_to, report_format ("text"/"pdf"/"excel")

أخرج JSON فقط بدون نص إضافي.

التاريخ: اليوم {today}، امس {yesterday}، اول امس {day_before}.
النص: "{text}"
"""

async def analyze_message(text: str) -> dict:
    if not OPENROUTER_API_KEY:
        return {"intent": "other"}
    today = datetime.now().strftime("%Y-%m-%d")
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    day_before = (datetime.now() - timedelta(days=2)).strftime("%Y-%m-%d")
    prompt = build_prompt(text, today, yesterday, day_before)
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
                headers=headers, json=payload, timeout=25
            )
            if res.status_code == 200:
                content = res.json()['choices'][0]['message']['content']
                return json.loads(content)
        except:
            pass
    return {"intent": "other"}

# ========== دوال الحالة والحوار ==========
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
        "ASK_WORKERS": "🧑‍🌾 كم عدد العمال؟",
        "ASK_WAGE_MODE": "💵 كيف تحدد الأجرة؟\n1️⃣ يومية\n2️⃣ بالساعة\nاكتب 1 أو 2.",
        "ASK_HOURS": "⏱️ كم ساعة عمل؟",
        "ASK_WAGE_PER_HOUR": "💰 كم أجرة الساعة؟",
        "ASK_WAGE_DAILY": "💰 كم أجرة العامل اليومية؟",
        "ASK_EXPENSES": "🧾 المصاريف الإضافية؟\nاكتب الرقم أو 'لا'.",
    }.get(step, "")

def build_confirmation_text(state: dict) -> str:
    wage = state.get("wage_per_worker") or 0
    workers = state.get("workers_count") or 0
    expenses = state.get("expenses") or 0
    total = (workers * wage) + expenses
    return (
        f"📋 راجع البيانات ({state.get('date')}):\n"
        f"- عمال: {workers}\n- أجرة/عامل: {wage:.2f} د.أ\n- مصاريف: {expenses:.2f} د.أ\n"
        f"- المجموع: {total:.2f} د.أ\n\n✅ اكتب 'نعم' للحفظ أو 'الغاء' للإلغاء."
    )

NO_EXPENSE_WORDS = ["لا", "ﻻ", "لا يوجد", "ماكو", "بدون", "no", "0"]
YES_WORDS = ["نعم", "ايوه", "ايه", "تمام", "صح", "اوك", "ok", "yes", "احفظ", "أكيد"]
CANCEL_WORDS = ["الغاء", "إلغاء", "كنسل", "cancel", "لغي", "الغي"]

async def handle_guided_answer(chat_id: int, text: str, state: dict):
    step = state["step"]
    stripped = text.strip()
    if step == "ASK_WORKERS":
        num = extract_number(stripped) or await extract_number_ai(stripped)
        if num is None:
            await send_telegram_message(chat_id, "لم أفهم الرقم، حاول مرة أخرى.")
            return
        state["workers_count"] = int(num)
    elif step == "ASK_WAGE_MODE":
        if "2" in stripped or "ساعة" in stripped:
            state["wage_mode"] = "hourly"
        elif "1" in stripped or "يومي" in stripped or "مباشر" in stripped:
            state["wage_mode"] = "daily"
        else:
            await send_telegram_message(chat_id, "اكتب 1 أو 2.")
            return
    elif step == "ASK_HOURS":
        num = extract_number(stripped) or await extract_number_ai(stripped)
        if num is None:
            await send_telegram_message(chat_id, "لم أفهم، كم ساعة؟")
            return
        state["hours"] = num
    elif step == "ASK_WAGE_PER_HOUR":
        num = extract_number(stripped) or await extract_number_ai(stripped)
        if num is None:
            await send_telegram_message(chat_id, "لم أفهم، كم أجرة الساعة؟")
            return
        state["wage_per_hour"] = num
        state["wage_per_worker"] = (state.get("hours") or 0) * num
        state["notes"] = (state.get("notes") or "") + f" {state['hours']}س×{num}د"
    elif step == "ASK_WAGE_DAILY":
        num = extract_number(stripped) or await extract_number_ai(stripped)
        if num is None:
            await send_telegram_message(chat_id, "لم أفهم، كم أجرة العامل؟")
            return
        state["wage_per_worker"] = num
    elif step == "ASK_EXPENSES":
        if any(w in stripped for w in NO_EXPENSE_WORDS) and extract_number(stripped) is None:
            state["expenses"] = 0.0
        else:
            num = extract_number(stripped) or await extract_number_ai(stripped)
            if num is None:
                await send_telegram_message(chat_id, "لم أفهم، اكتب رقماً أو 'لا'.")
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
            await send_telegram_message(chat_id,
                f"✅ تم حفظ اليومية ({record_date})\nعمال: {workers}\nأجرة: {wage:.2f}\nمصاريف: {expenses:.2f}\nالمجموع: {total:.2f}")
        elif any(w in stripped for w in CANCEL_WORDS):
            clear_state(chat_id)
            await send_telegram_message(chat_id, "❌ تم الإلغاء.")
        else:
            await send_telegram_message(chat_id, "اكتب 'نعم' أو 'الغاء'.")
        return
    step_after = next_missing_step(state)
    save_state(chat_id, step=step_after, **state)
    if step_after == "CONFIRM":
        await send_telegram_message(chat_id, build_confirmation_text(state))
    else:
        await send_telegram_message(chat_id, question_for_step(step_after))

async def start_guided_flow(chat_id: int, prefill: dict):
    workers = int(prefill.get("workers_count") or 0) or None
    hours = prefill.get("hours") or None
    wage_per_hour = prefill.get("wage_per_hour") or None
    wage_per_worker = prefill.get("wage_per_worker") or None
    expenses = prefill.get("expenses")
    date_val = prefill.get("date") or datetime.now().strftime("%Y-%m-%d")
    wage_mode = None
    if wage_per_worker:
        wage_mode = "daily"
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
        intro = "تمام، خلينا نكمل 👇\n\n" if workers or wage_per_worker else ""
        await send_telegram_message(chat_id, intro + question_for_step(step))

# ========== دوال التحديث والتقارير ==========
def query_records(date_from=None, date_to=None):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    if date_from and date_to:
        cursor.execute(
            "SELECT date, workers_count, wage_per_worker, expenses, notes FROM daily_records WHERE date BETWEEN ? AND ? ORDER BY date ASC",
            (date_from, date_to)
        )
    else:
        cursor.execute("SELECT date, workers_count, wage_per_worker, expenses, notes FROM daily_records ORDER BY date ASC")
    rows = cursor.fetchall()
    conn.close()
    return rows

def format_text_report(rows, title):
    if not rows:
        return f"{title}\n\nلا توجد سجلات."
    lines = [title, ""]
    grand = 0.0
    for date, count, wage, exp, notes in rows:
        total = (count * wage) + exp
        grand += total
        lines.append(f"📅 {date} | عمال: {count} | أجرة: {wage:.2f} | مصاريف: {exp:.2f} | المجموع: {total:.2f}")
        if notes:
            lines.append(f"   ملاحظات: {notes}")
    lines.append("")
    lines.append(f"🏆 الإجمالي الكلي: {grand:.2f} د.أ")
    return "\n".join(lines)

def generate_excel_report(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "تقرير"
    headers = ["التاريخ", "عدد العمال", "أجرة العامل", "المصاريف", "المجموع", "ملاحظات"]
    ws.append(headers)
    for date, count, wage, exp, notes in rows:
        total = (count * wage) + exp
        ws.append([date, count, wage, exp, total, notes])
    wb.save("temp.xlsx")
    with open("temp.xlsx", "rb") as f:
        data = f.read()
    os.remove("temp.xlsx")
    return data

def generate_pdf_report(rows, title):
    html = f"""
    <html><head><meta charset="UTF-8"><style>
    body {{ font-family: Arial; direction: rtl; }}
    h1 {{ color: #2c3e50; }}
    table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
    th, td {{ border: 1px solid #ccc; padding: 8px; text-align: center; }}
    th {{ background-color: #f2f2f2; }}
    </style></head><body>
    <h1>{title}</h1>
    <table><tr><th>التاريخ</th><th>العمال</th><th>أجرة العامل</th><th>المصاريف</th><th>المجموع</th><th>ملاحظات</th></tr>
    """
    grand = 0.0
    for date, count, wage, exp, notes in rows:
        total = (count * wage) + exp
        grand += total
        html += f"<tr><td>{date}</td><td>{count}</td><td>{wage:.2f}</td><td>{exp:.2f}</td><td>{total:.2f}</td><td>{notes or ''}</td></tr>"
    html += f"</table><p style='font-weight:bold;'>الإجمالي الكلي: {grand:.2f} د.أ</p></body></html>"
    return HTML(string=html).write_pdf()

async def handle_update(chat_id: int, update_data: dict):
    target = update_data.get("update_target_date")
    action = update_data.get("update_action")
    value = update_data.get("update_value")
    note = update_data.get("update_note", "")
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    if target == "last":
        cursor.execute("SELECT id, date, workers_count, wage_per_worker, expenses, notes FROM daily_records ORDER BY id DESC LIMIT 1")
        row = cursor.fetchone()
        if not row:
            await send_telegram_message(chat_id, "لا يوجد سجلات.")
            conn.close()
            return
        record_id, old_date, old_workers, old_wage, old_exp, old_notes = row
        current = {"date": old_date, "workers": old_workers, "wage": old_wage, "expenses": old_exp, "notes": old_notes}
    else:
        cursor.execute("SELECT id, date, workers_count, wage_per_worker, expenses, notes FROM daily_records WHERE date = ?", (target,))
        row = cursor.fetchone()
        if not row:
            await send_telegram_message(chat_id, f"لا يوجد سجل بتاريخ {target}.")
            conn.close()
            return
        record_id, old_date, old_workers, old_wage, old_exp, old_notes = row
        current = {"date": old_date, "workers": old_workers, "wage": old_wage, "expenses": old_exp, "notes": old_notes}

    new_exp = current["expenses"]
    new_wage = current["wage"]
    new_workers = current["workers"]
    new_notes = current["notes"] or ""
    if action == "add_expense":
        new_exp += value
        new_notes += f" إضافة مصاريف {value:.2f}."
    elif action == "add_expense_per_worker":
        new_exp += value * current["workers"]
        new_notes += f" إضافة مصاريف لكل عامل {value:.2f}."
    elif action == "set_expense":
        new_exp = value
        new_notes += f" تعديل المصاريف إلى {value:.2f}."
    elif action == "set_wage":
        new_wage = value
        new_notes += f" تعديل الأجرة إلى {value:.2f}."
    elif action == "set_workers":
        new_workers = int(value)
        new_notes += f" تعديل العمال إلى {int(value)}."
    elif action == "delete_record":
        cursor.execute("DELETE FROM daily_records WHERE id = ?", (record_id,))
        conn.commit()
        conn.close()
        await send_telegram_message(chat_id, f"🗑️ تم حذف سجل {current['date']}.")
        return
    elif action == "append_note":
        new_notes += f" {note}"
    else:
        await send_telegram_message(chat_id, "عملية غير معروفة.")
        conn.close()
        return
    cursor.execute(
        "UPDATE daily_records SET workers_count=?, wage_per_worker=?, expenses=?, notes=? WHERE id=?",
        (new_workers, new_wage, new_exp, new_notes.strip(), record_id)
    )
    conn.commit()
    conn.close()
    await send_telegram_message(chat_id, f"✅ تم تحديث سجل {current['date']}.")

async def handle_report(chat_id: int, report_data: dict):
    scope = report_data.get("report_scope", "all")
    fmt = report_data.get("report_format", "text")
    date_from = report_data.get("report_date_from")
    date_to = report_data.get("report_date_to")
    if scope == "single_day":
        date_from = date_to = date_from or date_to
    elif scope == "all":
        date_from = date_to = None
    rows = query_records(date_from, date_to)
    if not rows:
        await send_telegram_message(chat_id, "لا توجد بيانات.")
        return
    if fmt == "text":
        title = f"تقرير المزرعة ({date_from or 'الكل'} - {date_to or 'الكل'})"
        await send_telegram_message(chat_id, format_text_report(rows, title))
    elif fmt == "excel":
        data = generate_excel_report(rows)
        await send_telegram_document(chat_id, "تقرير.xlsx", data, "تقرير Excel")
    elif fmt == "pdf":
        title = f"تقرير المزرعة ({date_from or 'الكل'} - {date_to or 'الكل'})"
        pdf = generate_pdf_report(rows, title)
        await send_telegram_document(chat_id, "تقرير.pdf", pdf, title)
    else:
        await send_telegram_message(chat_id, "صيغة غير مدعومة.")

# ========== المعالج الرئيسي ==========
async def handle_incoming_message(chat_id: int, text: str):
    await send_telegram_message(chat_id, "⏳ جارٍ المعالجة...")
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
            await send_telegram_message(chat_id,
                "مرحباً! أرسل تفاصيل اليومية (مثل: 5 عمال، 10 دنانير)، أو اطلب تقريراً، أو تحديثاً.")
    except Exception as e:
        logger.error(f"Error: {e}", exc_info=True)
        await send_telegram_message(chat_id, "❌ حدث خطأ، حاول لاحقاً.")

# ========== حلقة Polling ==========
async def polling_loop():
    offset = 0
    logger.info("🔄 بدء polling...")
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

# ========== FastAPI Startup ==========
@app.on_event("startup")
async def startup_event():
    if not TELEGRAM_BOT_TOKEN:
        logger.error("❌ TELEGRAM_BOT_TOKEN غير موجود")
        return
    asyncio.create_task(polling_loop())
    logger.info("✅ البوت يعمل عبر polling")

@app.get("/")
async def root():
    return {"status": "running", "mode": "polling"}

# ========== التشغيل ==========
if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=PORT)
