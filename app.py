# app.py
# ============================================================
# MEDJOL FARM MANAGER V2
# بوت ذكي لإدارة العمال واليوميات والمصاريف والتقارير
# ============================================================

import os
import re
import json
import html
import sqlite3
import logging
import asyncio
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import httpx
import uvicorn
from fastapi import FastAPI
from openpyxl import Workbook
from weasyprint import HTML


# ============================================================
# CONFIG
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s"
)

logger = logging.getLogger("medjol")

app = FastAPI(title="Medjol Farm Manager V2")


TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")

OPENROUTER_MODEL = os.getenv(
    "OPENROUTER_MODEL",
    "meta-llama/llama-3.3-70b-instruct"
)

PORT = int(os.getenv("PORT", "8000"))

DB_NAME = os.getenv("DB_NAME", "farm.db")

# الأردن
TIMEZONE = os.getenv("TIMEZONE", "Asia/Amman")

# أجر افتراضي فقط إذا لم يكن للعامل أجر خاص
DEFAULT_HOURLY_RATE = float(
    os.getenv("DEFAULT_HOURLY_RATE", "5")
)

MAX_WORKERS = 100
MAX_HOURS_PER_DAY = 24
MAX_EXPENSE = 1_000_000


# ============================================================
# TIME
# ============================================================

def now_local():
    return datetime.now(ZoneInfo(TIMEZONE))


def today_str():
    return now_local().strftime("%Y-%m-%d")


# ============================================================
# DATABASE
# ============================================================

def db():
    conn = sqlite3.connect(DB_NAME, timeout=30)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():

    conn = db()
    c = conn.cursor()

    # --------------------------------------------------------
    # العمال
    # --------------------------------------------------------

    c.execute("""
        CREATE TABLE IF NOT EXISTS workers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE COLLATE NOCASE,
            phone TEXT,
            hourly_rate REAL NOT NULL DEFAULT 5,
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL
        )
    """)

    # --------------------------------------------------------
    # اليوميات
    # --------------------------------------------------------

    c.execute("""
        CREATE TABLE IF NOT EXISTS daily_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL UNIQUE,
            notes TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
    """)

    # --------------------------------------------------------
    # ساعات العمال
    # --------------------------------------------------------

    c.execute("""
        CREATE TABLE IF NOT EXISTS work_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            daily_log_id INTEGER NOT NULL,
            worker_id INTEGER NOT NULL,
            hours REAL NOT NULL,
            task TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            UNIQUE(daily_log_id, worker_id),
            FOREIGN KEY(daily_log_id)
                REFERENCES daily_logs(id)
                ON DELETE CASCADE,
            FOREIGN KEY(worker_id)
                REFERENCES workers(id)
        )
    """)

    # --------------------------------------------------------
    # المصاريف
    # --------------------------------------------------------

    c.execute("""
        CREATE TABLE IF NOT EXISTS expenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            daily_log_id INTEGER NOT NULL,
            category TEXT DEFAULT 'عام',
            amount REAL NOT NULL,
            notes TEXT DEFAULT '',
            FOREIGN KEY(daily_log_id)
                REFERENCES daily_logs(id)
                ON DELETE CASCADE
        )
    """)

    # --------------------------------------------------------
    # جلسات الحوار
    # --------------------------------------------------------

    c.execute("""
        CREATE TABLE IF NOT EXISTS sessions (
            chat_id INTEGER PRIMARY KEY,
            step TEXT NOT NULL,
            data TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
    """)

    # --------------------------------------------------------
    # الإعدادات
    # --------------------------------------------------------

    c.execute("""
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
    """)

    c.execute("""
        INSERT OR IGNORE INTO settings(key, value)
        VALUES ('default_hourly_rate', ?)
    """, (str(DEFAULT_HOURLY_RATE),))

    conn.commit()
    conn.close()


init_db()


# ============================================================
# SETTINGS
# ============================================================

def get_setting(key, default=None):

    conn = db()
    row = conn.execute(
        "SELECT value FROM settings WHERE key = ?",
        (key,)
    ).fetchone()

    conn.close()

    if not row:
        return default

    return row["value"]


def set_setting(key, value):

    conn = db()

    conn.execute("""
        INSERT INTO settings(key, value)
        VALUES (?, ?)
        ON CONFLICT(key)
        DO UPDATE SET value = excluded.value
    """, (key, str(value)))

    conn.commit()
    conn.close()


def default_hourly_rate():

    try:
        return float(
            get_setting(
                "default_hourly_rate",
                DEFAULT_HOURLY_RATE
            )
        )
    except Exception:
        return DEFAULT_HOURLY_RATE


# ============================================================
# TELEGRAM
# ============================================================

async def telegram_request(method, **kwargs):

    if not TELEGRAM_BOT_TOKEN:
        logger.error("TELEGRAM_BOT_TOKEN مفقود")
        return None

    url = (
        f"https://api.telegram.org/"
        f"bot{TELEGRAM_BOT_TOKEN}/{method}"
    )

    try:

        async with httpx.AsyncClient() as client:

            response = await client.post(
                url,
                **kwargs,
                timeout=30
            )

            if response.status_code != 200:
                logger.error(
                    "Telegram %s error: %s",
                    method,
                    response.text
                )
                return None

            return response.json()

    except Exception as e:

        logger.exception(
            "Telegram request failed: %s",
            e
        )

        return None


async def send_message(chat_id, text):

    return await telegram_request(
        "sendMessage",
        json={
            "chat_id": chat_id,
            "text": text
        }
    )


async def send_document(
    chat_id,
    filename,
    content,
    caption=""
):

    if not TELEGRAM_BOT_TOKEN:
        return None

    url = (
        f"https://api.telegram.org/"
        f"bot{TELEGRAM_BOT_TOKEN}/sendDocument"
    )

    try:

        async with httpx.AsyncClient() as client:

            response = await client.post(
                url,
                data={
                    "chat_id": chat_id,
                    "caption": caption
                },
                files={
                    "document": (
                        filename,
                        content,
                        "application/octet-stream"
                    )
                },
                timeout=60
            )

            if response.status_code != 200:

                logger.error(
                    "Document error: %s",
                    response.text
                )

                return None

            return response.json()

    except Exception as e:

        logger.exception(
            "Document error: %s",
            e
        )

        return None


# ============================================================
# TEXT NORMALIZATION
# ============================================================

def normalize_text(text):

    if not text:
        return ""

    text = text.strip()

    # تحويل الأرقام العربية
    translation = str.maketrans(
        "٠١٢٣٤٥٦٧٨٩",
        "0123456789"
    )

    text = text.translate(translation)

    text = text.replace(
        "٫",
        "."
    )

    text = text.replace(
        "،",
        ","
    )

    return text


def normalize_name(name):

    name = normalize_text(name)

    name = re.sub(
        r"\s+",
        " ",
        name
    )

    name = name.strip(
        " ,.;:،؛-–—"
    )

    return name


def number_from_text(text):

    text = normalize_text(text)

    match = re.search(
        r"(?<!\d)(\d+(?:[.,]\d+)?)(?!\d)",
        text
    )

    if not match:
        return None

    try:
        return float(
            match.group(1).replace(",", ".")
        )
    except Exception:
        return None


def is_no(text):

    text = normalize_text(text).lower()

    return any(
        phrase in text
        for phrase in [
            "لا",
            "ما في",
            "مفيش",
            "بدون",
            "ما عندي",
            "لا يوجد",
            "0"
        ]
    )


def is_yes(text):

    text = normalize_text(text).lower()

    return any(
        phrase in text
        for phrase in [
            "نعم",
            "ايوه",
            "أيوه",
            "اه",
            "آه",
            "موافق",
            "تأكيد",
            "أكد",
            "احفظ",
            "حفظ",
            "صحيح"
        ]
    )


# ============================================================
# SESSION
# ============================================================

def get_session(chat_id):

    conn = db()

    row = conn.execute(
        """
        SELECT step, data
        FROM sessions
        WHERE chat_id = ?
        """,
        (chat_id,)
    ).fetchone()

    conn.close()

    if not row:
        return None

    try:

        return {
            "step": row["step"],
            "data": json.loads(row["data"])
        }

    except Exception:

        clear_session(chat_id)
        return None


def save_session(
    chat_id,
    step,
    data
):

    conn = db()

    conn.execute("""
        INSERT INTO sessions(
            chat_id,
            step,
            data,
            updated_at
        )
        VALUES (?, ?, ?, ?)
        ON CONFLICT(chat_id)
        DO UPDATE SET
            step = excluded.step,
            data = excluded.data,
            updated_at = excluded.updated_at
    """, (
        chat_id,
        step,
        json.dumps(
            data,
            ensure_ascii=False
        ),
        now_local().isoformat()
    ))

    conn.commit()
    conn.close()


def clear_session(chat_id):

    conn = db()

    conn.execute(
        "DELETE FROM sessions WHERE chat_id = ?",
        (chat_id,)
    )

    conn.commit()
    conn.close()


# ============================================================
# VALIDATION
# ============================================================

def validate_worker_name(name):

    name = normalize_name(name)

    if not name:
        return False

    if len(name) < 2:
        return False

    if len(name) > 100:
        return False

    # منع أسماء تبدو كجمل كاملة
    if len(name.split()) > 6:
        return False

    return True


def clean_names(names):

    if not isinstance(names, list):
        return []

    result = []

    for name in names:

        if not isinstance(name, str):
            continue

        name = normalize_name(name)

        if not validate_worker_name(name):
            continue

        if name.lower() not in [
            x.lower() for x in result
        ]:
            result.append(name)

    return result


def validate_extracted_data(data):

    if not isinstance(data, dict):
        return False, "صيغة بيانات غير صحيحة."

    workers_count = data.get("workers_count")
    names = clean_names(
        data.get("workers_names", [])
    )

    hours = data.get("total_hours")
    expenses = data.get("expenses")

    # --------------------------------------------------------
    # workers_count
    # --------------------------------------------------------

    if workers_count is not None:

        try:
            workers_count = int(
                float(workers_count)
            )
        except Exception:

            return False, "عدد العمال غير صحيح."

        if not 1 <= workers_count <= MAX_WORKERS:

            return False, (
                "عدد العمال خارج النطاق المقبول."
            )

    # --------------------------------------------------------
    # hours
    # --------------------------------------------------------

    if hours is not None:

        try:
            hours = float(hours)
        except Exception:

            return False, "عدد الساعات غير صحيح."

        if not 0 <= hours <= MAX_HOURS_PER_DAY:

            return False, (
                "عدد الساعات يجب أن يكون بين "
                "0 و24 ساعة."
            )

    # --------------------------------------------------------
    # expenses
    # --------------------------------------------------------

    if expenses is not None:

        try:
            expenses = float(expenses)
        except Exception:

            return False, "المصاريف غير صحيحة."

        if not 0 <= expenses <= MAX_EXPENSE:

            return False, "قيمة المصاريف غير صحيحة."

    # --------------------------------------------------------
    # consistency
    # --------------------------------------------------------

    if names and workers_count:

        if len(names) != workers_count:

            return False, (
                f"ذكرت {workers_count} عمال، "
                f"لكن وجدت {len(names)} أسماء. "
                "لن أخمّن العدد الصحيح."
            )

    return True, None


# ============================================================
# AI EXTRACTION
# ============================================================

AI_SYSTEM_PROMPT = """
أنت محرك استخراج بيانات فقط لنظام إدارة مزرعة.

ممنوع التخمين.
ممنوع اختراع أي اسم.
ممنوع استنتاج عدد العمال من شيء غير واضح.
ممنوع اعتبار كلمات مثل ساعة، دينار، مصاريف، عامل أسماء أشخاص.

استخرج فقط المعلومات الموجودة صراحة في النص.

إذا لم توجد معلومة، أعد null.

workers_names يجب أن تكون ARRAY من أسماء العمال فقط.

إذا قال المستخدم:
"أحمد ومحمد وخالد"
فالأسماء:
["أحمد","محمد","خالد"]

إذا قال:
"3 عمال"
ولا توجد أسماء:
workers_count = 3
workers_names = []

إذا قال:
"أحمد ومحمد"
ولا ذكر العدد:
workers_count = 2
workers_names = ["أحمد","محمد"]

hours هو عدد ساعات العمل.

expenses هو مجموع المصاريف المذكورة إذا كان واضحاً.

notes تحتوي فقط على المعلومات الإضافية المهمة.

لا تضع تفسيراً خارج JSON.

JSON المطلوب:

{
  "workers_count": number|null,
  "workers_names": ["name"] ,
  "total_hours": number|null,
  "expenses": number|null,
  "notes": string|null
}
"""


async def extract_data_with_ai(text):

    if not OPENROUTER_API_KEY:
        return None

    text = normalize_text(text)

    prompt = (
        AI_SYSTEM_PROMPT
        + "\n\nالنص:\n"
        + text
    )

    headers = {
        "Authorization":
            f"Bearer {OPENROUTER_API_KEY}",

        "Content-Type":
            "application/json",

        "HTTP-Referer":
            "https://medjol.onrender.com",

        "X-Title":
            "Medjol Farm Manager"
    }

    payload = {
        "model": OPENROUTER_MODEL,

        "messages": [
            {
                "role": "system",
                "content": AI_SYSTEM_PROMPT
            },
            {
                "role": "user",
                "content": text
            }
        ],

        "temperature": 0,

        "response_format": {
            "type": "json_object"
        }
    }

    try:

        async with httpx.AsyncClient() as client:

            response = await client.post(
                "https://openrouter.ai/api/v1/chat/completions",
                headers=headers,
                json=payload,
                timeout=30
            )

        if response.status_code != 200:

            logger.error(
                "OpenRouter error %s: %s",
                response.status_code,
                response.text
            )

            return None

        body = response.json()

        content = (
            body
            .get("choices", [{}])[0]
            .get("message", {})
            .get("content")
        )

        if not content:
            return None

        data = json.loads(content)

        # توحيد الشكل
        if not isinstance(
            data.get("workers_names"),
            list
        ):
            data["workers_names"] = []

        data["workers_names"] = clean_names(
            data["workers_names"]
        )

        valid, error = validate_extracted_data(
            data
        )

        if not valid:

            logger.warning(
                "AI data rejected: %s",
                error
            )

            return {
                "_validation_error": error
            }

        return data

    except Exception as e:

        logger.exception(
            "AI extraction failed: %s",
            e
        )

        return None


# ============================================================
# WORKERS
# ============================================================

def get_worker_by_name(name):

    name = normalize_name(name)

    conn = db()

    row = conn.execute(
        """
        SELECT *
        FROM workers
        WHERE name = ? COLLATE NOCASE
        """,
        (name,)
    ).fetchone()

    conn.close()

    return row


def create_worker(
    name,
    hourly_rate=None
):

    name = normalize_name(name)

    if not validate_worker_name(name):
        return None

    if hourly_rate is None:
        hourly_rate = default_hourly_rate()

    try:
        hourly_rate = float(hourly_rate)

        if hourly_rate < 0:
            return None

    except Exception:
        return None

    conn = db()

    try:

        cursor = conn.execute(
            """
            INSERT INTO workers(
                name,
                hourly_rate,
                active,
                created_at
            )
            VALUES (?, ?, 1, ?)
            """,
            (
                name,
                hourly_rate,
                now_local().isoformat()
            )
        )

        conn.commit()

        worker_id = cursor.lastrowid

    except sqlite3.IntegrityError:

        row = conn.execute(
            """
            SELECT id
            FROM workers
            WHERE name = ? COLLATE NOCASE
            """,
            (name,)
        ).fetchone()

        worker_id = (
            row["id"]
            if row
            else None
        )

    finally:
        conn.close()

    return worker_id


def get_active_workers():

    conn = db()

    rows = conn.execute("""
        SELECT *
        FROM workers
        WHERE active = 1
        ORDER BY name COLLATE NOCASE
    """).fetchall()

    conn.close()

    return rows


def deactivate_worker(name):

    conn = db()

    cursor = conn.execute(
        """
        UPDATE workers
        SET active = 0
        WHERE name = ? COLLATE NOCASE
        """,
        (normalize_name(name),)
    )

    conn.commit()

    changed = cursor.rowcount > 0

    conn.close()

    return changed


# ============================================================
# DAILY LOGS
# ============================================================

def get_daily_log(date):

    conn = db()

    daily = conn.execute(
        """
        SELECT *
        FROM daily_logs
        WHERE date = ?
        """,
        (date,)
    ).fetchone()

    if not daily:

        conn.close()
        return None

    workers = conn.execute("""
        SELECT
            w.id,
            w.name,
            w.hourly_rate,
            wl.hours,
            wl.task,
            wl.notes
        FROM work_logs wl
        JOIN workers w
            ON w.id = wl.worker_id
        WHERE wl.daily_log_id = ?
        ORDER BY w.name COLLATE NOCASE
    """, (daily["id"],)).fetchall()

    expenses = conn.execute("""
        SELECT *
        FROM expenses
        WHERE daily_log_id = ?
        ORDER BY id
    """, (daily["id"],)).fetchall()

    conn.close()

    return {
        "id": daily["id"],
        "date": daily["date"],
        "notes": daily["notes"] or "",
        "workers": workers,
        "expenses": expenses
    }


def daily_exists(date):

    conn = db()

    row = conn.execute(
        """
        SELECT id
        FROM daily_logs
        WHERE date = ?
        """,
        (date,)
    ).fetchone()

    conn.close()

    return row is not None


def save_daily_data(data):

    """
    الحفظ داخل transaction واحدة.
    إذا فشل أي جزء -> rollback.
    """

    date = data["date"]

    workers = data["workers"]

    expenses = data.get("expenses", [])

    notes = data.get("notes", "")

    conn = db()

    try:

        conn.execute("BEGIN")

        existing = conn.execute(
            """
            SELECT id
            FROM daily_logs
            WHERE date = ?
            """,
            (date,)
        ).fetchone()

        if existing:

            raise ValueError(
                f"هناك يومية موجودة بالفعل بتاريخ {date}"
            )

        cursor = conn.execute(
            """
            INSERT INTO daily_logs(
                date,
                notes,
                created_at,
                updated_at
            )
            VALUES (?, ?, ?, ?)
            """,
            (
                date,
                notes,
                now_local().isoformat(),
                now_local().isoformat()
            )
        )

        daily_id = cursor.lastrowid

        for worker in workers:

            name = normalize_name(
                worker["name"]
            )

            hours = float(
                worker["hours"]
            )

            rate = worker.get(
                "hourly_rate"
            )

            if rate is None:
                rate = default_hourly_rate()

            worker_row = conn.execute(
                """
                SELECT id
                FROM workers
                WHERE name = ? COLLATE NOCASE
                """,
                (name,)
            ).fetchone()

            if worker_row:

                worker_id = worker_row["id"]

                # لا نغير أجر العامل الموجود
                # من خلال اليومية.

            else:

                cursor = conn.execute(
                    """
                    INSERT INTO workers(
                        name,
                        hourly_rate,
                        active,
                        created_at
                    )
                    VALUES (?, ?, 1, ?)
                    """,
                    (
                        name,
                        float(rate),
                        now_local().isoformat()
                    )
                )

                worker_id = cursor.lastrowid

            conn.execute(
                """
                INSERT INTO work_logs(
                    daily_log_id,
                    worker_id,
                    hours,
                    task,
                    notes
                )
                VALUES (?, ?, ?, ?, ?)
                """,
                (
                    daily_id,
                    worker_id,
                    hours,
                    worker.get("task", ""),
                    worker.get("notes", "")
                )
            )

        for expense in expenses:

            amount = float(
                expense["amount"]
            )

            if amount < 0:
                raise ValueError(
                    "لا يمكن أن تكون المصروفات سالبة."
                )

            conn.execute(
                """
                INSERT INTO expenses(
                    daily_log_id,
                    category,
                    amount,
                    notes
                )
                VALUES (?, ?, ?, ?)
                """,
                (
                    daily_id,
                    expense.get(
                        "category",
                        "عام"
                    ),
                    amount,
                    expense.get(
                        "notes",
                        ""
                    )
                )
            )

        conn.commit()

        return daily_id

    except Exception:

        conn.rollback()
        raise

    finally:

        conn.close()


# ============================================================
# CALCULATIONS
# ============================================================

def calculate_daily_total(log):

    wages = 0.0

    for worker in log["workers"]:

        wages += (
            float(worker["hours"])
            * float(worker["hourly_rate"])
        )

    expenses = sum(
        float(x["amount"])
        for x in log["expenses"]
    )

    return wages, expenses, wages + expenses


# ============================================================
# DISPLAY
# ============================================================

def format_money(value):

    return f"{float(value):.2f} د.أ"


def format_daily_summary(log):

    wages, expenses, total = (
        calculate_daily_total(log)
    )

    lines = []

    lines.append(
        f"📋 اليومية: {log['date']}"
    )

    lines.append("")
    lines.append(
        f"👷 عدد العمال: {len(log['workers'])}"
    )

    if log["workers"]:

        lines.append("")

        for index, worker in enumerate(
            log["workers"],
            1
        ):

            wage = (
                float(worker["hours"])
                * float(worker["hourly_rate"])
            )

            lines.append(
                f"{index}. {worker['name']} — "
                f"{worker['hours']} ساعة — "
                f"{format_money(wage)}"
            )

    lines.append("")
    lines.append(
        f"💵 أجور العمال: {format_money(wages)}"
    )

    lines.append(
        f"💰 المصاريف: {format_money(expenses)}"
    )

    lines.append(
        f"🏆 الإجمالي: {format_money(total)}"
    )

    if log["expenses"]:

        lines.append("")
        lines.append("🧾 المصاريف:")

        for expense in log["expenses"]:

            lines.append(
                f"- {expense['category']}: "
                f"{format_money(expense['amount'])}"
                + (
                    f" ({expense['notes']})"
                    if expense["notes"]
                    else ""
                )
            )

    if log["notes"]:

        lines.append("")
        lines.append(
            f"📝 ملاحظات: {log['notes']}"
        )

    return "\n".join(lines)


# ============================================================
# REPORT DATA
# ============================================================

def get_logs_between(
    date_from,
    date_to
):

    conn = db()

    daily_rows = conn.execute(
        """
        SELECT *
        FROM daily_logs
        WHERE date BETWEEN ? AND ?
        ORDER BY date
        """,
        (
            date_from,
            date_to
        )
    ).fetchall()

    result = []

    for daily in daily_rows:

        workers = conn.execute("""
            SELECT
                w.name,
                w.hourly_rate,
                wl.hours,
                wl.task
            FROM work_logs wl
            JOIN workers w
                ON w.id = wl.worker_id
            WHERE wl.daily_log_id = ?
            ORDER BY w.name COLLATE NOCASE
        """, (daily["id"],)).fetchall()

        expenses = conn.execute("""
            SELECT
                category,
                amount,
                notes
            FROM expenses
            WHERE daily_log_id = ?
        """, (daily["id"],)).fetchall()

        result.append({
            "date": daily["date"],
            "notes": daily["notes"] or "",
            "workers": workers,
            "expenses": expenses
        })

    conn.close()

    return result


# ============================================================
# TEXT REPORT
# ============================================================

def generate_text_report(
    logs,
    title="📊 التقرير"
):

    if not logs:
        return "لا توجد بيانات في الفترة المطلوبة."

    lines = [
        title,
        "=" * 35
    ]

    grand_wages = 0.0
    grand_expenses = 0.0

    for log in logs:

        wages = sum(
            float(w["hours"])
            * float(w["hourly_rate"])
            for w in log["workers"]
        )

        expenses = sum(
            float(e["amount"])
            for e in log["expenses"]
        )

        total = wages + expenses

        grand_wages += wages
        grand_expenses += expenses

        lines.append("")
        lines.append(
            f"📅 {log['date']}"
        )

        lines.append(
            f"👷 العمال: {len(log['workers'])}"
        )

        lines.append(
            f"⏱️ مجموع الساعات: "
            f"{sum(float(w['hours']) for w in log['workers']):.2f}"
        )

        lines.append(
            f"💵 الأجور: {format_money(wages)}"
        )

        lines.append(
            f"💰 المصاريف: "
            f"{format_money(expenses)}"
        )

        lines.append(
            f"🏆 الإجمالي: "
            f"{format_money(total)}"
        )

    lines.append("")
    lines.append("=" * 35)

    lines.append(
        f"💵 إجمالي الأجور: "
        f"{format_money(grand_wages)}"
    )

    lines.append(
        f"💰 إجمالي المصاريف: "
        f"{format_money(grand_expenses)}"
    )

    lines.append(
        f"🏆 الإجمالي النهائي: "
        f"{format_money(grand_wages + grand_expenses)}"
    )

    return "\n".join(lines)


# ============================================================
# EXCEL
# ============================================================

def generate_excel_report(logs):

    wb = Workbook()

    ws = wb.active
    ws.title = "التقرير"

    ws.append([
        "التاريخ",
        "العامل",
        "الساعات",
        "أجر الساعة",
        "أجر العامل",
        "نوع المهمة",
        "ملاحظات"
    ])

    for log in logs:

        for worker in log["workers"]:

            wage = (
                float(worker["hours"])
                * float(worker["hourly_rate"])
            )

            ws.append([
                log["date"],
                worker["name"],
                float(worker["hours"]),
                float(worker["hourly_rate"]),
                wage,
                worker["task"],
                log["notes"]
            ])

    expense_ws = wb.create_sheet(
        "المصاريف"
    )

    expense_ws.append([
        "التاريخ",
        "التصنيف",
        "المبلغ",
        "ملاحظات"
    ])

    for log in logs:

        for expense in log["expenses"]:

            expense_ws.append([
                log["date"],
                expense["category"],
                float(expense["amount"]),
                expense["notes"]
            ])

    summary = wb.create_sheet(
        "ملخص"
    )

    summary.append([
        "التاريخ",
        "عدد العمال",
        "الساعات",
        "الأجور",
        "المصاريف",
        "الإجمالي"
    ])

    for log in logs:

        wages = sum(
            float(w["hours"])
            * float(w["hourly_rate"])
            for w in log["workers"]
        )

        expenses = sum(
            float(e["amount"])
            for e in log["expenses"]
        )

        summary.append([
            log["date"],
            len(log["workers"]),
            sum(
                float(w["hours"])
                for w in log["workers"]
            ),
            wages,
            expenses,
            wages + expenses
        ])

    filename = "report.xlsx"

    wb.save(filename)

    with open(
        filename,
        "rb"
    ) as file:

        content = file.read()

    os.remove(filename)

    return content


# ============================================================
# PDF
# ============================================================

def generate_pdf_report(logs):

    rows_html = ""

    grand_wages = 0
    grand_expenses = 0

    for log in logs:

        wages = sum(
            float(w["hours"])
            * float(w["hourly_rate"])
            for w in log["workers"]
        )

        expenses = sum(
            float(e["amount"])
            for e in log["expenses"]
        )

        grand_wages += wages
        grand_expenses += expenses

        for worker in log["workers"]:

            wage = (
                float(worker["hours"])
                * float(worker["hourly_rate"])
            )

            rows_html += f"""
            <tr>
                <td>{html.escape(log["date"])}</td>
                <td>{html.escape(worker["name"])}</td>
                <td>{worker["hours"]}</td>
                <td>{worker["hourly_rate"]:.2f}</td>
                <td>{wage:.2f}</td>
            </tr>
            """

    total = (
        grand_wages
        + grand_expenses
    )

    document = f"""
    <!DOCTYPE html>
    <html lang="ar" dir="rtl">
    <head>
        <meta charset="UTF-8">

        <style>

        body {{
            font-family: Arial, sans-serif;
            direction: rtl;
            padding: 30px;
        }}

        h1 {{
            text-align: center;
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 25px;
        }}

        th, td {{
            border: 1px solid #999;
            padding: 8px;
            text-align: center;
        }}

        th {{
            background: #eeeeee;
        }}

        .summary {{
            margin-top: 25px;
            font-size: 18px;
            line-height: 2;
        }}

        </style>
    </head>

    <body>

        <h1>تقرير إدارة المزرعة</h1>

        <table>

            <tr>
                <th>التاريخ</th>
                <th>العامل</th>
                <th>الساعات</th>
                <th>أجر الساعة</th>
                <th>الأجر</th>
            </tr>

            {rows_html}

        </table>

        <div class="summary">

            <strong>
                إجمالي أجور العمال:
                {grand_wages:.2f} د.أ
            </strong>

            <br>

            <strong>
                إجمالي المصاريف:
                {grand_expenses:.2f} د.أ
            </strong>

            <br>

            <strong>
                الإجمالي النهائي:
                {total:.2f} د.أ
            </strong>

        </div>

    </body>
    </html>
    """

    return HTML(
        string=document
    ).write_pdf()


# ============================================================
# COMMANDS
# ============================================================

async def handle_command(
    chat_id,
    text
):

    normalized = normalize_text(
        text
    ).lower()

    # --------------------------------------------------------
    # CANCEL
    # --------------------------------------------------------

    if (
        normalized in [
            "/cancel",
            "الغاء",
            "إلغاء",
            "الغاء العملية"
        ]
    ):

        clear_session(chat_id)

        await send_message(
            chat_id,
            "❌ تم إلغاء العملية."
        )

        return True

    # --------------------------------------------------------
    # HELP
    # --------------------------------------------------------

    if normalized in [
        "/start",
        "/help",
        "مساعدة"
    ]:

        clear_session(chat_id)

        await send_message(
            chat_id,
            """
🤖 أهلاً بك في Medjol Farm Manager

أرسل بيانات العمل بأي طريقة طبيعية.

مثال:

"اليوم اشتغل أحمد ومحمد وخالد 8 ساعات وصرفت 10 دنانير بنزين"

أو:

"عندي 3 عمال: أحمد، محمد، خالد. كل واحد 7 ساعات. مصاريف 15"

الأوامر:

📋 تقرير
📅 تقرير اليوم
📆 تقرير الأسبوع
🗓️ تقرير الشهر
📄 تقرير PDF
📗 تقرير Excel

👷 العمال
➕ أضف العامل أحمد
❌ احذف العامل أحمد

⚙️ الإعدادات

❌ إلغاء
"""
        )

        return True

    # --------------------------------------------------------
    # WORKERS LIST
    # --------------------------------------------------------

    if normalized in [
        "العمال",
        "قائمة العمال",
        "/workers"
    ]:

        workers = get_active_workers()

        if not workers:

            await send_message(
                chat_id,
                "لا يوجد عمال مسجلون."
            )

            return True

        lines = [
            "👷 العمال المسجلون:",
            ""
        ]

        for i, worker in enumerate(
            workers,
            1
        ):

            lines.append(
                f"{i}. {worker['name']} "
                f"— {worker['hourly_rate']:.2f} د.أ/ساعة"
            )

        await send_message(
            chat_id,
            "\n".join(lines)
        )

        return True

    # --------------------------------------------------------
    # ADD WORKER
    # --------------------------------------------------------

    if (
        normalized.startswith(
            "اضف العامل "
        )
        or normalized.startswith(
            "أضف العامل "
        )
        or normalized.startswith(
            "/addworker "
        )
    ):

        if normalized.startswith(
            "/addworker "
        ):

            name = text.split(
                " ",
                1
            )[1]

        else:

            name = re.sub(
                r"^(اضف|أضف)\s+العامل\s+",
                "",
                text,
                flags=re.IGNORECASE
            )

        name = normalize_name(name)

        if not validate_worker_name(name):

            await send_message(
                chat_id,
                "⚠️ اسم العامل غير واضح."
            )

            return True

        if get_worker_by_name(name):

            await send_message(
                chat_id,
                f"⚠️ العامل {name} موجود بالفعل."
            )

            return True

        create_worker(name)

        await send_message(
            chat_id,
            f"✅ تم إضافة العامل: {name}"
        )

        return True

    # --------------------------------------------------------
    # DELETE WORKER
    # --------------------------------------------------------

    if (
        normalized.startswith(
            "احذف العامل "
        )
        or normalized.startswith(
            "حذف العامل "
        )
    ):

        name = re.sub(
            r"^(احذف|حذف)\s+العامل\s+",
            "",
            text,
            flags=re.IGNORECASE
        )

        name = normalize_name(name)

        if deactivate_worker(name):

            await send_message(
                chat_id,
                f"✅ تم تعطيل العامل: {name}"
            )

        else:

            await send_message(
                chat_id,
                f"⚠️ لم أجد العامل: {name}"
            )

        return True

    # --------------------------------------------------------
    # SETTINGS
    # --------------------------------------------------------

    if normalized in [
        "الاعدادات",
        "الإعدادات",
        "/settings"
    ]:

        await send_message(
            chat_id,
            f"""
⚙️ الإعدادات

أجر الساعة الافتراضي:
{default_hourly_rate():.2f} د.أ

لتغييره اكتب:

"أجر الساعة 6"

أو:

"سعر الساعة 6"
"""
        )

        return True

    # --------------------------------------------------------
    # RATE
    # --------------------------------------------------------

    rate_match = re.search(
        r"(?:أجر|اجر|سعر)\s*(?:الساعة|الساعه)\s*[:=]?\s*(\d+(?:[.,]\d+)?)",
        normalize_text(text)
    )

    if rate_match:

        try:

            rate = float(
                rate_match.group(1)
                .replace(",", ".")
            )

            if rate < 0:
                raise ValueError

            set_setting(
                "default_hourly_rate",
                rate
            )

            await send_message(
                chat_id,
                f"✅ تم تغيير أجر الساعة الافتراضي إلى "
                f"{rate:.2f} د.أ."
            )

        except Exception:

            await send_message(
                chat_id,
                "⚠️ قيمة الأجر غير صحيحة."
            )

        return True

    return False


# ============================================================
# REPORT COMMANDS
# ============================================================

async def handle_report_command(
    chat_id,
    text
):

    t = normalize_text(
        text
    ).lower()

    is_report = (
        "تقرير" in t
        or t in [
            "/report",
            "report"
        ]
    )

    if not is_report:
        return False

    # --------------------------------------------------------
    # DATE RANGE
    # --------------------------------------------------------

    today = now_local().date()

    if "اليوم" in t:

        date_from = today
        date_to = today

    elif "اسبوع" in t or "أسبوع" in t:

        date_from = (
            today
            - timedelta(
                days=today.weekday()
            )
        )

        date_to = today

    elif (
        "شهر" in t
        or "شهري" in t
    ):

        date_from = today.replace(
            day=1
        )

        date_to = today

    else:

        # افتراضي: كل السجلات
        date_from = datetime(
            2000,
            1,
            1
        ).date()

        date_to = today

    logs = get_logs_between(
        date_from.isoformat(),
        date_to.isoformat()
    )

    # --------------------------------------------------------
    # PDF
    # --------------------------------------------------------

    if "pdf" in t:

        if not logs:

            await send_message(
                chat_id,
                "لا توجد بيانات لإنشاء التقرير."
            )

            return True

        pdf = generate_pdf_report(
            logs
        )

        await send_document(
            chat_id,
            "تقرير_المزرعة.pdf",
            pdf,
            "📄 تقرير المزرعة"
        )

        return True

    # --------------------------------------------------------
    # EXCEL
    # --------------------------------------------------------

    if (
        "excel" in t
        or "اكسل" in t
        or "إكسل" in t
    ):

        if not logs:

            await send_message(
                chat_id,
                "لا توجد بيانات لإنشاء التقرير."
            )

            return True

        excel = generate_excel_report(
            logs
        )

        await send_document(
            chat_id,
            "تقرير_المزرعة.xlsx",
            excel,
            "📗 تقرير المزرعة"
        )

        return True

    # --------------------------------------------------------
    # TEXT
    # --------------------------------------------------------

    if (
        date_from == date_to
    ):

        title = (
            f"📊 تقرير يوم "
            f"{date_from.isoformat()}"
        )

    else:

        title = (
            f"📊 التقرير من "
            f"{date_from.isoformat()} "
            f"إلى "
            f"{date_to.isoformat()}"
        )

    report = generate_text_report(
        logs,
        title
    )

    await send_message(
        chat_id,
        report
    )

    return True


# ============================================================
# BUILD CONFIRMATION DATA
# ============================================================

def build_confirmation_data(
    ai_data
):

    names = clean_names(
        ai_data.get(
            "workers_names",
            []
        )
    )

    count = ai_data.get(
        "workers_count"
    )

    hours = ai_data.get(
        "total_hours"
    )

    expenses = ai_data.get(
        "expenses"
    )

    # --------------------------------------------------------
    # If names exist but count doesn't
    # --------------------------------------------------------

    if names and count is None:

        count = len(names)

    # --------------------------------------------------------
    # If count exists but names don't
    # --------------------------------------------------------

    workers = []

    if names:

        for name in names:

            existing = get_worker_by_name(
                name
            )

            rate = (
                float(existing["hourly_rate"])
                if existing
                else default_hourly_rate()
            )

            workers.append({
                "name": name,
                "hours": hours,
                "hourly_rate": rate,
                "task": "",
                "notes": ""
            })

    return {
        "date": today_str(),
        "workers_count": count,
        "workers": workers,
        "pending_hours": hours,
        "pending_expenses": expenses,
        "expenses": (
            []
            if expenses is None
            else [{
                "category": "عام",
                "amount": float(expenses),
                "notes": ""
            }]
        ),
        "notes": (
            ai_data.get("notes")
            or ""
        )
    }


# ============================================================
# CONFIRMATION MESSAGE
# ============================================================

def build_confirmation_message(data):

    lines = []

    lines.append(
        f"📋 مراجعة اليومية "
        f"({data['date']})"
    )

    lines.append("")

    workers = data.get(
        "workers",
        []
    )

    if workers:

        lines.append(
            f"👷 عدد العمال: {len(workers)}"
        )

        for i, worker in enumerate(
            workers,
            1
        ):

            hours = (
                worker.get("hours")
                or 0
            )

            rate = (
                worker.get(
                    "hourly_rate"
                )
                or default_hourly_rate()
            )

            wage = (
                float(hours)
                * float(rate)
            )

            lines.append(
                f"{i}. {worker['name']} — "
                f"{hours} ساعة × "
                f"{rate:.2f} = "
                f"{wage:.2f} د.أ"
            )

    else:

        count = data.get(
            "workers_count"
        )

        lines.append(
            f"👷 عدد العمال: "
            f"{count or 'غير محدد'}"
        )

    expenses = sum(
        float(e["amount"])
        for e in data.get(
            "expenses",
            []
        )
    )

    wages = sum(
        float(w.get("hours") or 0)
        * float(
            w.get(
                "hourly_rate"
            )
            or default_hourly_rate()
        )
        for w in workers
    )

    lines.append("")
    lines.append(
        f"💵 الأجور: "
        f"{wages:.2f} د.أ"
    )

    lines.append(
        f"💰 المصاريف: "
        f"{expenses:.2f} د.أ"
    )

    lines.append(
        f"🏆 الإجمالي: "
        f"{wages + expenses:.2f} د.أ"
    )

    if data.get("notes"):

        lines.append("")
        lines.append(
            f"📝 {data['notes']}"
        )

    lines.append("")
    lines.append(
        "هل البيانات صحيحة؟"
    )

    lines.append("")
    lines.append(
        "✅ نعم / حفظ"
    )

    lines.append(
        "✏️ تعديل"
    )

    lines.append(
        "❌ إلغاء"
    )

    return "\n".join(lines)


# ============================================================
# EDIT COMMAND
# ============================================================

async def edit_pending_data(
    chat_id,
    text,
    data
):

    t = normalize_text(
        text
    )

    # --------------------------------------------------------
    # HOURS
    # --------------------------------------------------------

    match = re.search(
        r"(?:الساعات|ساعات|ساعة|الساعه)\s*(?:=|:)?\s*(\d+(?:[.,]\d+)?)",
        t,
        flags=re.IGNORECASE
    )

    if match:

        hours = float(
            match.group(1).replace(
                ",",
                "."
            )
        )

        if not 0 <= hours <= 24:

            await send_message(
                chat_id,
                "⚠️ الساعات يجب أن تكون بين 0 و24."
            )

            return True

        for worker in data.get(
            "workers",
            []
        ):

            worker["hours"] = hours

        data["pending_hours"] = hours

        save_session(
            chat_id,
            "AWAITING_CONFIRMATION",
            data
        )

        await send_message(
            chat_id,
            build_confirmation_message(
                data
            )
        )

        return True

    # --------------------------------------------------------
    # EXPENSE
    # --------------------------------------------------------

    match = re.search(
        r"(?:المصاريف|مصاريف|صرف|دفعت)\s*(?:=|:)?\s*(\d+(?:[.,]\d+)?)",
        t,
        flags=re.IGNORECASE
    )

    if match:

        amount = float(
            match.group(1).replace(
                ",",
                "."
            )
        )

        if amount < 0:

            await send_message(
                chat_id,
                "⚠️ المصروف لا يمكن أن يكون سالباً."
            )

            return True

        data["expenses"] = [{
            "category": "عام",
            "amount": amount,
            "notes": ""
        }]

        save_session(
            chat_id,
            "AWAITING_CONFIRMATION",
            data
        )

        await send_message(
            chat_id,
            build_confirmation_message(
                data
            )
        )

        return True

    # --------------------------------------------------------
    # ADD WORKER
    # --------------------------------------------------------

    match = re.search(
        r"(?:أضف|اضف)\s+(?:العامل\s+)?(.+)",
        t,
        flags=re.IGNORECASE
    )

    if match:

        name = normalize_name(
            match.group(1)
        )

        if validate_worker_name(name):

            existing = get_worker_by_name(
                name
            )

            if not existing:

                create_worker(name)

            rate = (
                float(existing["hourly_rate"])
                if existing
                else default_hourly_rate()
            )

            data.setdefault(
                "workers",
                []
            ).append({
                "name": name,
                "hours": data.get(
                    "pending_hours"
                ),
                "hourly_rate": rate,
                "task": "",
                "notes": ""
            })

            data["workers_count"] = len(
                data["workers"]
            )

            save_session(
                chat_id,
                "AWAITING_CONFIRMATION",
                data
            )

            await send_message(
                chat_id,
                build_confirmation_message(
                    data
                )
            )

            return True

    return False


# ============================================================
# SESSION HANDLER
# ============================================================

async def handle_active_session(
    chat_id,
    text,
    session
):

    step = session["step"]

    data = session["data"]

    # --------------------------------------------------------
    # CONFIRMATION
    # --------------------------------------------------------

    if step == "AWAITING_CONFIRMATION":

        if is_yes(text):

            # يجب توفر العمال
            if not data.get("workers"):

                await send_message(
                    chat_id,
                    "⚠️ لا يمكن الحفظ قبل تحديد أسماء العمال."
                )

                return True

            # يجب توفر ساعات
            missing_hours = [
                w["name"]
                for w in data["workers"]
                if w.get("hours") is None
            ]

            if missing_hours:

                await send_message(
                    chat_id,
                    "⚠️ لم يتم تحديد ساعات العمل لـ:\n"
                    + "\n".join(
                        missing_hours
                    )
                )

                return True

            # منع duplicate
            if daily_exists(
                data["date"]
            ):

                await send_message(
                    chat_id,
                    f"⚠️ توجد يومية محفوظة بالفعل "
                    f"بتاريخ {data['date']}.\n"
                    "لن أستبدلها تلقائياً."
                )

                clear_session(chat_id)

                return True

            try:

                save_daily_data(
                    data
                )

            except Exception as e:

                logger.exception(
                    "Save failed: %s",
                    e
                )

                await send_message(
                    chat_id,
                    "❌ حدث خطأ أثناء الحفظ. "
                    "لم يتم حفظ أي جزء من البيانات."
                )

                return True

            clear_session(chat_id)

            log = get_daily_log(
                data["date"]
            )

            await send_message(
                chat_id,
                "✅ تم حفظ اليومية بنجاح.\n\n"
                + format_daily_summary(
                    log
                )
            )

            return True

        if (
            "الغاء" in normalize_text(text)
            or "إلغاء" in text
        ):

            clear_session(chat_id)

            await send_message(
                chat_id,
                "❌ تم إلغاء العملية."
            )

            return True

        if (
            "تعديل" in normalize_text(text)
            or "عدل" in normalize_text(text)
        ):

            save_session(
                chat_id,
                "EDITING",
                data
            )

            await send_message(
                chat_id,
                """
✏️ ماذا تريد تعديل؟

مثال:

"الساعات 7"

أو:

"المصاريف 15"

أو:

"أضف العامل أحمد"

اكتب التعديل مباشرة.
"""
            )

            return True

        await send_message(
            chat_id,
            "اكتب: نعم للحفظ، تعديل للتعديل، أو إلغاء."
        )

        return True

    # --------------------------------------------------------
    # EDITING
    # --------------------------------------------------------

    if step == "EDITING":

        if await edit_pending_data(
            chat_id,
            text,
            data
        ):

            return True

        await send_message(
            chat_id,
            "لم أفهم التعديل. مثال: الساعات 7 أو المصاريف 15."
        )

        return True

    # --------------------------------------------------------
    # WAITING FOR COUNT
    # --------------------------------------------------------

    if step == "ASK_WORKERS_COUNT":

        count = number_from_text(
            text
        )

        if count is None:

            await send_message(
                chat_id,
                "اكتب عدد العمال كرقم."
            )

            return True

        count = int(count)

        if not 1 <= count <= MAX_WORKERS:

            await send_message(
                chat_id,
                f"عدد العمال يجب أن يكون بين 1 و{MAX_WORKERS}."
            )

            return True

        data["workers_count"] = count

        save_session(
            chat_id,
            "ASK_NAMES",
            data
        )

        await send_message(
            chat_id,
            f"👷 تم تسجيل {count} عمال.\n"
            "اكتب أسماءهم مفصولة بفواصل."
        )

        return True

    # --------------------------------------------------------
    # WAITING FOR NAMES
    # --------------------------------------------------------

    if step == "ASK_NAMES":

        names = clean_names(
            re.split(
                r"[,،\n]+",
                text
            )
        )

        expected = data.get(
            "workers_count"
        )

        if len(names) != expected:

            await send_message(
                chat_id,
                f"⚠️ قلت إن عدد العمال {expected}، "
                f"لكن وجدت {len(names)} أسماء.\n"
                "لن أخمّن.\n\n"
                "اكتب الأسماء مفصولة بفواصل."
            )

            return True

        data["workers"] = []

        for name in names:

            existing = get_worker_by_name(
                name
            )

            rate = (
                float(existing["hourly_rate"])
                if existing
                else default_hourly_rate()
            )

            data["workers"].append({
                "name": name,
                "hours": None,
                "hourly_rate": rate,
                "task": "",
                "notes": ""
            })

        save_session(
            chat_id,
            "ASK_HOURS",
            data
        )

        await send_message(
            chat_id,
            "⏱️ كم ساعة عمل لكل عامل؟\n"
            "إذا كانت الساعات مختلفة، اكتبها لكل عامل.\n\n"
            "مثال:\n"
            "أحمد 8، محمد 7، خالد 8"
        )

        return True

    # --------------------------------------------------------
    # HOURS
    # --------------------------------------------------------

    if step == "ASK_HOURS":

        # محاولة استخراج اسم + ساعات
        pairs = re.findall(
            r"([^\d,،\n]+?)\s+(\d+(?:[.,]\d+)?)",
            normalize_text(text)
        )

        if pairs:

            assigned = {}

            for raw_name, raw_hours in pairs:

                name = normalize_name(
                    raw_name
                )

                hours = float(
                    raw_hours.replace(
                        ",",
                        "."
                    )
                )

                if not 0 <= hours <= 24:

                    await send_message(
                        chat_id,
                        "⚠️ الساعات يجب أن تكون بين 0 و24."
                    )

                    return True

                assigned[
                    name.lower()
                ] = hours

            missing = []

            for worker in data["workers"]:

                key = worker["name"].lower()

                if key in assigned:

                    worker["hours"] = assigned[key]

                else:

                    missing.append(
                        worker["name"]
                    )

            if missing:

                await send_message(
                    chat_id,
                    "لم أحدد ساعات:\n"
                    + "\n".join(missing)
                )

                return True

        else:

            hours = number_from_text(
                text
            )

            if hours is None:

                await send_message(
                    chat_id,
                    "اكتب عدد الساعات."
                )

                return True

            if not 0 <= hours <= 24:

                await send_message(
                    chat_id,
                    "الساعات يجب أن تكون بين 0 و24."
                )

                return True

            for worker in data["workers"]:

                worker["hours"] = hours

            data["pending_hours"] = hours

        save_session(
            chat_id,
            "ASK_EXPENSES",
            data
        )

        await send_message(
            chat_id,
            "💰 كم المصاريف الإضافية؟\n"
            "اكتب المبلغ أو اكتب «لا»."
        )

        return True

    # --------------------------------------------------------
    # EXPENSES
    # --------------------------------------------------------

    if step == "ASK_EXPENSES":

        if is_no(text):

            data["expenses"] = []

        else:

            amount = number_from_text(
                text
            )

            if amount is None:

                await send_message(
                    chat_id,
                    "اكتب مبلغ المصاريف أو «لا»."
                )

                return True

            if amount < 0:

                await send_message(
                    chat_id,
                    "المصاريف لا يمكن أن تكون سالبة."
                )

                return True

            data["expenses"] = [{
                "category": "عام",
                "amount": amount,
                "notes": ""
            }]

        save_session(
            chat_id,
            "ASK_NOTES",
            data
        )

        await send_message(
            chat_id,
            "📝 هل توجد ملاحظات؟\n"
            "اكتبها أو «لا»."
        )

        return True

    # --------------------------------------------------------
    # NOTES
    # --------------------------------------------------------

    if step == "ASK_NOTES":

        if is_no(text):

            data["notes"] = ""

        else:

            data["notes"] = text.strip()

        save_session(
            chat_id,
            "AWAITING_CONFIRMATION",
            data
        )

        await send_message(
            chat_id,
            build_confirmation_message(
                data
            )
        )

        return True

    return False


# ============================================================
# MAIN MESSAGE HANDLER
# ============================================================

async def handle_incoming_message(
    chat_id,
    text
):

    text = text.strip()

    if not text:
        return

    # --------------------------------------------------------
    # COMMANDS FIRST
    # --------------------------------------------------------

    if await handle_command(
        chat_id,
        text
    ):

        return

    # --------------------------------------------------------
    # REPORTS BEFORE AI
    # --------------------------------------------------------

    if await handle_report_command(
        chat_id,
        text
    ):

        return

    # --------------------------------------------------------
    # ACTIVE SESSION
    # --------------------------------------------------------

    session = get_session(
        chat_id
    )

    if session:

        await handle_active_session(
            chat_id,
            text,
            session
        )

        return

    # --------------------------------------------------------
    # AI
    # --------------------------------------------------------

    await send_message(
        chat_id,
        "⏳ جارٍ فهم البيانات..."
    )

    ai_data = await extract_data_with_ai(
        text
    )

    if not ai_data:

        await send_message(
            chat_id,
            """
لم أستطع فهم الرسالة بشكل موثوق.

جرب مثلاً:

"اليوم اشتغل أحمد ومحمد 8 ساعات وصرفت 10 دنانير"

أو:

"3 عمال: أحمد، محمد، خالد، 7 ساعات، مصاريف 15"
"""
        )

        return

    # --------------------------------------------------------
    # VALIDATION ERROR
    # --------------------------------------------------------

    if ai_data.get(
        "_validation_error"
    ):

        await send_message(
            chat_id,
            "⚠️ "
            + ai_data["_validation_error"]
        )

        return

    # --------------------------------------------------------
    # BUILD DATA
    # --------------------------------------------------------

    data = build_confirmation_data(
        ai_data
    )

    names = data.get(
        "workers",
        []
    )

    count = data.get(
        "workers_count"
    )

    # --------------------------------------------------------
    # NO WORKERS
    # --------------------------------------------------------

    if not count and not names:

        await send_message(
            chat_id,
            "كم عدد العمال؟"
        )

        save_session(
            chat_id,
            "ASK_WORKERS_COUNT",
            data
        )

        return

    # --------------------------------------------------------
    # COUNT BUT NO NAMES
    # --------------------------------------------------------

    if count and not names:

        save_session(
            chat_id,
            "ASK_NAMES",
            data
        )

        await send_message(
            chat_id,
            f"👷 فهمت أن عدد العمال {count}.\n"
            "اكتب أسماء العمال مفصولة بفواصل."
        )

        return

    # --------------------------------------------------------
    # NAMES BUT HOURS MISSING
    # --------------------------------------------------------

    missing_hours = [
        w["name"]
        for w in names
        if w.get("hours") is None
    ]

    if missing_hours:

        save_session(
            chat_id,
            "ASK_HOURS",
            data
        )

        await send_message(
            chat_id,
            "⏱️ كم ساعة عمل لكل عامل؟\n"
            "يمكنك كتابة رقم واحد إذا كانت الساعات متساوية."
        )

        return

    # --------------------------------------------------------
    # ALL DATA PRESENT
    # --------------------------------------------------------

    save_session(
        chat_id,
        "AWAITING_CONFIRMATION",
        data
    )

    await send_message(
        chat_id,
        build_confirmation_message(
            data
        )
    )


# ============================================================
# TELEGRAM POLLING
# ============================================================

async def delete_webhook():

    if not TELEGRAM_BOT_TOKEN:
        return

    url = (
        f"https://api.telegram.org/"
        f"bot{TELEGRAM_BOT_TOKEN}/deleteWebhook"
    )

    try:

        async with httpx.AsyncClient() as client:

            await client.get(
                url,
                params={
                    "drop_pending_updates": False
                },
                timeout=15
            )

    except Exception as e:

        logger.error(
            "Webhook deletion failed: %s",
            e
        )


async def polling_loop():

    await delete_webhook()

    offset = 0

    logger.info(
        "🚀 Medjol Farm Manager V2 started"
    )

    while True:

        try:

            url = (
                f"https://api.telegram.org/"
                f"bot{TELEGRAM_BOT_TOKEN}/getUpdates"
            )

            async with httpx.AsyncClient() as client:

                response = await client.get(
                    url,
                    params={
                        "offset": offset,
                        "timeout": 30,
                        "allowed_updates": json.dumps(
                            ["message"]
                        )
                    },
                    timeout=40
                )

            if response.status_code != 200:

                logger.error(
                    "Polling HTTP error: %s",
                    response.status_code
                )

                await asyncio.sleep(3)
                continue

            updates = response.json().get(
                "result",
                []
            )

            for update in updates:

                offset = (
                    update["update_id"]
                    + 1
                )

                message = update.get(
                    "message"
                )

                if not message:
                    continue

                if "text" not in message:
                    continue

                chat = message.get(
                    "chat",
                    {}
                )

                chat_id = chat.get(
                    "id"
                )

                text = message.get(
                    "text",
                    ""
                ).strip()

                if not chat_id or not text:
                    continue

                try:

                    await handle_incoming_message(
                        chat_id,
                        text
                    )

                except Exception as e:

                    logger.exception(
                        "Message handling error: %s",
                        e
                    )

                    await send_message(
                        chat_id,
                        "❌ حدث خطأ غير متوقع. "
                        "لم يتم حفظ هذه العملية."
                    )

        except Exception as e:

            logger.exception(
                "Polling exception: %s",
                e
            )

            await asyncio.sleep(5)


# ============================================================
# FASTAPI
# ============================================================

@app.on_event("startup")
async def startup():

    if not TELEGRAM_BOT_TOKEN:

        logger.error(
            "❌ TELEGRAM_BOT_TOKEN غير موجود"
        )

        return

    asyncio.create_task(
        polling_loop()
    )

    logger.info(
        "✅ Bot polling started"
    )


@app.get("/")
async def root():

    return {
        "status": "running",
        "name": "Medjol Farm Manager",
        "version": "2.0",
        "mode": "polling"
    }


@app.get("/health")
async def health():

    return {
        "status": "healthy",
        "time": now_local().isoformat()
    }


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    uvicorn.run(
        app,
        host="0.0.0.0",
        port=PORT
    )
