# app.py – بوت مشرف العمال (نسخة ذكية ومبسطة)

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
DB_NAME = "farm.db"

# ========== قاعدة البيانات ==========
def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    # جدول اليوميات
    c.execute('''
        CREATE TABLE IF NOT EXISTS daily_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT UNIQUE,
            workers_count INTEGER,
            workers_names TEXT,
            total_hours REAL,
            expenses REAL,
            notes TEXT,
            created_at TEXT
        )
    ''')
    # جدول الحالة (للحوار)
    c.execute('''
        CREATE TABLE IF NOT EXISTS session (
            chat_id INTEGER PRIMARY KEY,
            step TEXT,
            data TEXT  -- JSON
        )
    ''')
    conn.commit()
    conn.close()
init_db()

# ========== إرسال رسائل ==========
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

# ========== استخراج الأرقام والنصوص ==========
def extract_number(text: str):
    match = re.search(r"\d+(\.\d+)?", text)
    if match:
        return float(match.group())
    return None

def extract_names(text: str):
    # استخراج الأسماء (كلمات بعد كلمة "اسم" أو "أسماؤهم" أو "هم")
    # بسيط: نأخذ كل الكلمات التي ليست أرقاماً
    words = re.findall(r"[^\d\s]+", text)
    # نفلتر الكلمات التي قد تكون أسماء (طولها > 2)
    names = [w for w in words if len(w) > 2 and not w.isdigit()]
    return names if names else None

# ========== الذكاء الاصطناعي لاستخراج البيانات ==========
async def extract_data_with_ai(text: str):
    if not OPENROUTER_API_KEY:
        return None
    prompt = f"""
أنت مساعد ذكي لمشرف عمال في مزرعة. استخرج من النص التالي:
- عدد العمال (عدد)
- أسماء العمال (قائمة أسماء مفصولة بفواصل)
- عدد ساعات العمل (رقم)
- المصاريف (رقم)
- أي ملاحظات إضافية

أخرج JSON بهذا الشكل:
{{
  "workers_count": عدد أو null,
  "workers_names": "قائمة الأسماء" أو null,
  "total_hours": رقم أو null,
  "expenses": رقم أو null,
  "notes": "نص" أو null
}}

النص: "{text}"
"""
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://medjol.onrender.com",
        "X-Title": "Medjol Bot"
    }
    payload = {
        "model": OPENROUTER_MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.2,
        "response_format": {"type": "json_object"}
    }
    async with httpx.AsyncClient() as client:
        try:
            res = await client.post(
                "https://openrouter.ai/api/v1/chat/completions",
                headers=headers, json=payload, timeout=20
            )
            if res.status_code == 200:
                content = res.json()['choices'][0]['message']['content']
                return json.loads(content)
        except Exception as e:
            logger.error(f"AI error: {e}")
    return None

# ========== دوال الجلسة (الحوار) ==========
def get_session(chat_id: int):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT step, data FROM session WHERE chat_id = ?", (chat_id,))
    row = c.fetchone()
    conn.close()
    if row:
        return {"step": row[0], "data": json.loads(row[1])}
    return None

def save_session(chat_id: int, step: str, data: dict):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("REPLACE INTO session (chat_id, step, data) VALUES (?, ?, ?)",
              (chat_id, step, json.dumps(data)))
    conn.commit()
    conn.close()

def clear_session(chat_id: int):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("DELETE FROM session WHERE chat_id = ?", (chat_id,))
    conn.commit()
    conn.close()

# ========== دوال اليوميات ==========
def save_daily_log(date, workers_count, workers_names, total_hours, expenses, notes):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''
        INSERT OR REPLACE INTO daily_logs
        (date, workers_count, workers_names, total_hours, expenses, notes, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', (date, workers_count, workers_names, total_hours, expenses, notes, datetime.now().isoformat()))
    conn.commit()
    conn.close()

def get_daily_log(date: str):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT * FROM daily_logs WHERE date = ?", (date,))
    row = c.fetchone()
    conn.close()
    if row:
        return {
            "id": row[0],
            "date": row[1],
            "workers_count": row[2],
            "workers_names": row[3],
            "total_hours": row[4],
            "expenses": row[5],
            "notes": row[6]
        }
    return None

def get_all_logs(date_from=None, date_to=None):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    if date_from and date_to:
        c.execute("SELECT * FROM daily_logs WHERE date BETWEEN ? AND ? ORDER BY date", (date_from, date_to))
    else:
        c.execute("SELECT * FROM daily_logs ORDER BY date")
    rows = c.fetchall()
    conn.close()
    return rows

# ========== توليد التقارير ==========
def generate_text_report(rows):
    if not rows:
        return "لا توجد بيانات."
    lines = ["📋 تقرير اليوميات", "="*30]
    grand_total = 0.0
    for row in rows:
        date, workers_count, workers_names, hours, expenses, notes = row[1], row[2], row[3], row[4], row[5], row[6]
        total_wages = (workers_count or 0) * (hours or 0) * 5  # افترض 5 دنانير للساعة (يمكن تعديله)
        total = total_wages + (expenses or 0)
        grand_total += total
        lines.append(f"\n📅 {date}")
        lines.append(f"   عدد العمال: {workers_count or 0}")
        if workers_names:
            lines.append(f"   الأسماء: {workers_names}")
        lines.append(f"   ساعات العمل: {hours or 0}")
        lines.append(f"   المصاريف: {expenses or 0:.2f} د.أ")
        lines.append(f"   الإجمالي: {total:.2f} د.أ")
        if notes:
            lines.append(f"   ملاحظات: {notes}")
    lines.append("\n" + "="*30)
    lines.append(f"🏆 الإجمالي الكلي: {grand_total:.2f} د.أ")
    return "\n".join(lines)

def generate_excel_report(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "اليوميات"
    headers = ["التاريخ", "عدد العمال", "أسماء العمال", "ساعات العمل", "المصاريف", "الملاحظات", "الإجمالي"]
    ws.append(headers)
    for row in rows:
        date, workers_count, names, hours, expenses, notes = row[1], row[2], row[3], row[4], row[5], row[6]
        total = (workers_count or 0) * (hours or 0) * 5 + (expenses or 0)
        ws.append([date, workers_count, names, hours, expenses, notes, total])
    wb.save("temp.xlsx")
    with open("temp.xlsx", "rb") as f:
        data = f.read()
    os.remove("temp.xlsx")
    return data

def generate_pdf_report(rows):
    html = """
    <html><head><meta charset="UTF-8"><style>
    body { font-family: Arial; direction: rtl; }
    h1 { color: #2c3e50; }
    table { width: 100%; border-collapse: collapse; margin-top: 20px; }
    th, td { border: 1px solid #ccc; padding: 8px; text-align: center; }
    th { background-color: #f2f2f2; }
    </style></head><body>
    <h1>تقرير اليوميات</h1>
    <table><tr><th>التاريخ</th><th>العمال</th><th>الأسماء</th><th>الساعات</th><th>المصاريف</th><th>الملاحظات</th><th>الإجمالي</th></tr>
    """
    grand = 0.0
    for row in rows:
        date, workers_count, names, hours, expenses, notes = row[1], row[2], row[3], row[4], row[5], row[6]
        total = (workers_count or 0) * (hours or 0) * 5 + (expenses or 0)
        grand += total
        html += f"<tr><td>{date}</td><td>{workers_count or 0}</td><td>{names or ''}</td><td>{hours or 0}</td><td>{expenses or 0:.2f}</td><td>{notes or ''}</td><td>{total:.2f}</td></tr>"
    html += f"</table><p style='font-weight:bold;'>الإجمالي الكلي: {grand:.2f} د.أ</p></body></html>"
    return HTML(string=html).write_pdf()

# ========== المعالج الرئيسي (ذكي) ==========
async def handle_incoming_message(chat_id: int, text: str):
    # أولاً: نرسل تأكيد استلام
    await send_telegram_message(chat_id, "⏳ جارٍ معالجة طلبك...")

    # نتحقق من وجود جلسة نشطة
    session = get_session(chat_id)
    if session:
        step = session["step"]
        data = session["data"]

        if step == "AWAITING_CONFIRMATION":
            if "نعم" in text or "تأكيد" in text or "حفظ" in text:
                # حفظ اليومية
                save_daily_log(
                    date=data["date"],
                    workers_count=data.get("workers_count"),
                    workers_names=data.get("workers_names"),
                    total_hours=data.get("total_hours"),
                    expenses=data.get("expenses"),
                    notes=data.get("notes")
                )
                clear_session(chat_id)
                await send_telegram_message(chat_id, "✅ تم حفظ اليومية بنجاح!")
                # نعرض ملخص
                log = get_daily_log(data["date"])
                if log:
                    msg = f"📋 ملخص اليوم ({log['date']}):\n"
                    msg += f"عدد العمال: {log['workers_count'] or 0}\n"
                    if log['workers_names']:
                        msg += f"الأسماء: {log['workers_names']}\n"
                    msg += f"ساعات العمل: {log['total_hours'] or 0}\n"
                    msg += f"المصاريف: {log['expenses'] or 0:.2f} د.أ\n"
                    total = (log['workers_count'] or 0) * (log['total_hours'] or 0) * 5 + (log['expenses'] or 0)
                    msg += f"الإجمالي: {total:.2f} د.أ"
                    await send_telegram_message(chat_id, msg)
            elif "الغاء" in text or "إلغاء" in text:
                clear_session(chat_id)
                await send_telegram_message(chat_id, "❌ تم إلغاء العملية.")
            else:
                await send_telegram_message(chat_id, "أكتب 'نعم' للتأكيد والحفظ، أو 'الغاء' للإلغاء.")
            return

        elif step == "AWAITING_NAMES":
            names = extract_names(text)
            if not names:
                await send_telegram_message(chat_id, "لم أفهم الأسماء، حاول كتابتها مفصولة بفواصل (مثل: أحمد، محمد، خالد).")
                return
            data["workers_names"] = ", ".join(names)
            data["workers_count"] = len(names)  # تحديث العدد تلقائياً
            # ننتقل للسؤال عن الساعات
            data["step"] = "ASK_HOURS"
            save_session(chat_id, "ASK_HOURS", data)
            await send_telegram_message(chat_id, f"✅ تم تسجيل {len(names)} عاملاً. كم عدد ساعات العمل اليوم؟")
            return

        elif step == "ASK_HOURS":
            num = extract_number(text)
            if num is None:
                await send_telegram_message(chat_id, "لم أفهم، كم ساعة؟ (اكتب رقماً)")
                return
            data["total_hours"] = num
            data["step"] = "ASK_EXPENSES"
            save_session(chat_id, "ASK_EXPENSES", data)
            await send_telegram_message(chat_id, f"✅ ساعات العمل: {num} ساعة. كم المصاريف الإضافية (نقل، أكل، مواد)؟ (اكتب رقماً أو 'لا')")
            return

        elif step == "ASK_EXPENSES":
            if "لا" in text or "ما في" in text or "0" in text:
                data["expenses"] = 0.0
            else:
                num = extract_number(text)
                if num is None:
                    await send_telegram_message(chat_id, "لم أفهم، اكتب رقم المصاريف أو 'لا'.")
                    return
                data["expenses"] = num
            # سؤال عن ملاحظات
            data["step"] = "ASK_NOTES"
            save_session(chat_id, "ASK_NOTES", data)
            await send_telegram_message(chat_id, "📝 هل تريد إضافة ملاحظات؟ (اكتبها، أو اكتب 'لا')")
            return

        elif step == "ASK_NOTES":
            if "لا" not in text:
                data["notes"] = text
            else:
                data["notes"] = ""
            # عرض ملخص للتأكيد
            date = data["date"]
            workers_count = data.get("workers_count", 0)
            workers_names = data.get("workers_names", "")
            hours = data.get("total_hours", 0)
            expenses = data.get("expenses", 0)
            total = workers_count * hours * 5 + expenses
            msg = f"📋 راجع بيانات اليوم ({date}):\n"
            msg += f"عدد العمال: {workers_count}\n"
            if workers_names:
                msg += f"الأسماء: {workers_names}\n"
            msg += f"ساعات العمل: {hours}\n"
            msg += f"المصاريف: {expenses:.2f} د.أ\n"
            msg += f"الإجمالي: {total:.2f} د.أ\n"
            if data.get("notes"):
                msg += f"ملاحظات: {data['notes']}\n"
            msg += "\n✅ هل تريد حفظ هذه البيانات؟ (نعم / الغاء)"
            data["step"] = "AWAITING_CONFIRMATION"
            save_session(chat_id, "AWAITING_CONFIRMATION", data)
            await send_telegram_message(chat_id, msg)
            return

    # لا جلسة نشطة: نحلل الرسالة بالذكاء الاصطناعي
    ai_data = await extract_data_with_ai(text)
    if ai_data and (ai_data.get("workers_count") or ai_data.get("workers_names")):
        # استخرجنا بيانات مفيدة
        today = datetime.now().strftime("%Y-%m-%d")
        data = {
            "date": today,
            "workers_count": ai_data.get("workers_count"),
            "workers_names": ai_data.get("workers_names"),
            "total_hours": ai_data.get("total_hours"),
            "expenses": ai_data.get("expenses"),
            "notes": ai_data.get("notes") or ""
        }
        # نتحقق من المفقودات
        if not data["workers_names"] and data["workers_count"]:
            # نطلب الأسماء
            save_session(chat_id, "AWAITING_NAMES", data)
            await send_telegram_message(chat_id, f"عدد العمال: {data['workers_count']}. الرجاء كتابة أسمائهم مفصولة بفواصل.")
            return
        elif not data["workers_count"] and not data["workers_names"]:
            # ما فيه بيانات كافية
            await send_telegram_message(chat_id, "لم أفهم بيانات اليوم، هل يمكنك إعادة صياغتها؟ (مثال: 5 عمال: أحمد، محمد، ساعة 8، مصاريف 10 دنانير)")
            return
        else:
            # اكتمل كل شيء تقريباً
            if not data["total_hours"]:
                save_session(chat_id, "ASK_HOURS", data)
                await send_telegram_message(chat_id, f"✅ تم التعرف على {data['workers_count'] or len(data['workers_names'].split(','))} عاملاً. كم عدد ساعات العمل اليوم؟")
                return
            if data["expenses"] is None:
                save_session(chat_id, "ASK_EXPENSES", data)
                await send_telegram_message(chat_id, f"✅ ساعات العمل: {data['total_hours']} ساعة. كم المصاريف الإضافية؟")
                return
            # كل البيانات موجودة، نعرض للتأكيد
            data["step"] = "AWAITING_CONFIRMATION"
            workers_count = data["workers_count"] or len(data["workers_names"].split(","))
            total = workers_count * (data["total_hours"] or 0) * 5 + (data["expenses"] or 0)
            msg = f"📋 بيانات اليوم ({data['date']}):\n"
            msg += f"عدد العمال: {workers_count}\n"
            if data["workers_names"]:
                msg += f"الأسماء: {data['workers_names']}\n"
            msg += f"ساعات العمل: {data['total_hours'] or 0}\n"
            msg += f"المصاريف: {data['expenses'] or 0:.2f} د.أ\n"
            msg += f"الإجمالي: {total:.2f} د.أ\n"
            if data["notes"]:
                msg += f"ملاحظات: {data['notes']}\n"
            msg += "\n✅ هل تريد حفظ هذه البيانات؟ (نعم / الغاء)"
            save_session(chat_id, "AWAITING_CONFIRMATION", data)
            await send_telegram_message(chat_id, msg)
            return
    else:
        # لم نفهم، نقدم مساعدة
        await send_telegram_message(chat_id, "مرحباً! أنا مساعدك الذكي لتسجيل اليوميات.\nأرسل لي بيانات اليوم مثل:\n'5 عمال: أحمد، محمد، خالد، ساعة 8، مصاريف 10 دنانير'\nأو اكتب 'تقرير' لعرض تقرير، أو 'تقرير PDF'، أو 'تقرير Excel'.")
        # نتعامل مع أوامر التقرير
        if "تقرير" in text:
            rows = get_all_logs()
            if not rows:
                await send_telegram_message(chat_id, "لا توجد سجلات بعد.")
                return
            if "pdf" in text.lower():
                pdf = generate_pdf_report(rows)
                await send_telegram_document(chat_id, "تقرير.pdf", pdf, "تقرير اليوميات")
            elif "excel" in text.lower() or "اكسل" in text:
                excel = generate_excel_report(rows)
                await send_telegram_document(chat_id, "تقرير.xlsx", excel, "تقرير اليوميات")
            else:
                report = generate_text_report(rows)
                await send_telegram_message(chat_id, report)
        return

# ========== حلقة Polling ==========
async def delete_webhook():
    if not TELEGRAM_BOT_TOKEN:
        return
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/deleteWebhook"
    async with httpx.AsyncClient() as client:
        try:
            await client.get(url, timeout=10)
        except:
            pass

async def polling_loop():
    await delete_webhook()
    offset = 0
    logger.info("🔄 بدء polling (البوت الذكي للعمال)...")
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

@app.on_event("startup")
async def startup():
    if not TELEGRAM_BOT_TOKEN:
        logger.error("❌ TELEGRAM_BOT_TOKEN مفقود")
        return
    asyncio.create_task(polling_loop())
    logger.info("✅ البوت الذكي للعمال يعمل عبر polling")

@app.get("/")
async def root():
    return {"status": "running", "mode": "polling"}

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=PORT)
