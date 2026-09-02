import os
import sqlite3
from io import BytesIO
from fastapi import FastAPI, Request, Response
import httpx
from openpyxl import Workbook
from weasyprint import HTML

app = FastAPI()

# المتغيرات البيئية
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
TELEGRAM_WEBHOOK_SECRET = os.getenv("TELEGRAM_WEBHOOK_SECRET")

# إعداد قاعدة البيانات
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

async def send_telegram_message(chat_id: int, text: str):
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    async with httpx.AsyncClient() as client:
        await client.post(url, json={"chat_id": chat_id, "text": text})

async def send_telegram_document(chat_id: int, filename: str, content: bytes, caption: str = ""):
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
    files = {"document": (filename, content)}
    data = {"chat_id": chat_id, "caption": caption}
    async with httpx.AsyncClient() as client:
        await client.post(url, data=data, files=files)

async def parse_with_groq(text: str) -> dict:
    if not GROQ_API_KEY:
        return {}
    
    prompt = f"""
أنت مساعد مالي لمزرعة نخل مجدول. قم باستخراج البيانات التالية من النص وإرجاعها بتنسيق JSON فقط بدون أي نص آخر:
- workers_count (عدد العمال كعدد صحيح)
- wage_per_worker (أجرة العامل بالدينار كعدد)
- expenses (المصاريف الإضافية كعدد)
- notes (أي تفاصيل أو ملاحظات)

النص: "{text}"
    """
    
    headers = {
        "Authorization": f"Bearer {GROQ_API_KEY}",
        "Content-Type": "application/json"
    }
    
    payload = {
        "model": "llama-3.3-70b-versatile",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.1,
        "response_format": {"type": "json_object"}
    }
    
    async with httpx.AsyncClient() as client:
        try:
            res = await client.post("https://api.groq.com/openai/v1/chat/completions", headers=headers, json=payload, timeout=20.0)
            if res.status_code == 200:
                import json
                result = res.json()
                content = result['choices'][0]['message']['content']
                return json.loads(content)
        except Exception:
            pass
    return {}

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

@app.post("/tg-webhook")
async def telegram_webhook(request: Request"):
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
                "أهلاً بك في بوت إدارة حسابات المزرعة!\nيمكنك كتابة اليوميات بشكل طبيعي (مثال: اليوم اشتغل 5 عمال اليومية 12 دينار ومصاريف 10) أو اطلب 'تقرير PDF' أو 'تقرير اكسل'."
            )
        elif "pdf" in text.lower() or "تقرير" in text:
            pdf_data = generate_pdf_report()
            await send_telegram_document(chat_id, "report.pdf", pdf_data, "إليك تقرير الحسابات بصيغة PDF")
        elif "excel" in text.lower() or "اكسل" in text:
            excel_data = generate_excel_report()
            await send_telegram_document(chat_id, "report.xlsx", excel_data, "إليك تقرير الحسابات بصيغة Excel")
        else:
            parsed = await parse_with_groq(text)
            if parsed and (parsed.get("workers_count") or parsed.get("expenses")):
                from datetime import datetime
                today = datetime.now().strftime("%Y-%m-%d")
                w_count = parsed.get("workers_count", 0) or 0
                wage = parsed.get("wage_per_worker", 0.0) or 0.0
                exp = parsed.get("expenses", 0.0) or 0.0
                notes = parsed.get("notes", "")
                
                conn = sqlite3.connect(DB_NAME)
                cursor = conn.cursor()
                cursor.execute(
                    "INSERT INTO daily_records (date, workers_count, wage_per_worker, expenses, notes) VALUES (?, ?, ?, ?, ?)",
                    (today, w_count, wage, exp, notes)
                )
                conn.commit()
                conn.close()
                
                total = (w_count * wage) + exp
                response_msg = f"تم تسجيل اليومية بنجاح!\n- عدد العمال: {w_count}\n- اليومية: {wage} د.أ\n- المصاريف: {exp} د.أ\n- المجموع: {total} د.أ"
                await send_telegram_message(chat_id, response_msg)
            else:
                await send_telegram_message(chat_id, "لم أتمكن من فهم البيانات. يرجى توضيح عدد العمال والأجرة أو طلب التقرير.")
                
    return {"status": "ok"}

