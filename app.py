import os
import re
import json
import html
import sqlite3
import logging
import asyncio
from io import BytesIO
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import httpx
import uvicorn
from fastapi import FastAPI
from openpyxl import Workbook


# ============================================================
# CONFIG
# ============================================================

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY", "").strip()
OPENROUTER_MODEL = os.getenv(
    "OPENROUTER_MODEL",
    "meta-llama/llama-3.3-70b-instruct"
).strip()

PORT = int(os.getenv("PORT", "8000"))
DB_NAME = os.getenv("DB_NAME", "farm.db")
TIMEZONE = os.getenv("TIMEZONE", "Asia/Amman")

# ثابت وغير قابل للتغيير من المستخدم
FIXED_HOURLY_RATE = 1.5

MAX_WORKERS = 100
MAX_HOURS_PER_WORKER = 24
MAX_EXPENSE = 1_000_000
MAX_TELEGRAM_TEXT = 3900


# ============================================================
# LOGGING
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

logger = logging.getLogger("farm_bot")


# ============================================================
# FASTAPI
# ============================================================

app = FastAPI(title="Farm Worker Management Bot")


@app.get("/")
def root():
    return {
        "status": "ok",
        "service": "farm-worker-bot"
    }


@app.get("/health")
def health():
    return {
        "status": "healthy"
    }


# ============================================================
# DATABASE
# ============================================================

def db_connect():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = db_connect()

    conn.executescript("""
    CREATE TABLE IF NOT EXISTS workers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        phone TEXT,
        hourly_rate REAL NOT NULL DEFAULT 1.5,
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS daily_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT NOT NULL UNIQUE,
        notes TEXT,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS work_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        daily_log_id INTEGER NOT NULL,
        worker_id INTEGER,
        hours REAL NOT NULL,
        task TEXT,
        notes TEXT,
        worker_name TEXT,
        UNIQUE(daily_log_id, worker_id, worker_name),
        FOREIGN KEY(daily_log_id) REFERENCES daily_logs(id)
            ON DELETE CASCADE,
        FOREIGN KEY(worker_id) REFERENCES workers(id)
            ON DELETE SET NULL
    );

    CREATE TABLE IF NOT EXISTS expenses (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        daily_log_id INTEGER NOT NULL,
        category TEXT NOT NULL,
        amount REAL NOT NULL,
        notes TEXT,
        FOREIGN KEY(daily_log_id) REFERENCES daily_logs(id)
            ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS sessions (
        chat_id INTEGER PRIMARY KEY,
        step TEXT NOT NULL,
        data TEXT NOT NULL,
        updated_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT
    );
    """)

    # توحيد الأجور القديمة على السعر الثابت
    conn.execute(
        "UPDATE workers SET hourly_rate = ?",
        (FIXED_HOURLY_RATE,)
    )

    conn.commit()
    conn.close()


# ============================================================
# DATE / TIME
# ============================================================

def now_local():
    return datetime.now(ZoneInfo(TIMEZONE))


def today_date():
    return now_local().date()


def today_str():
    return today_date().isoformat()


# ============================================================
# TEXT NORMALIZATION
# ============================================================

ARABIC_DIGITS = str.maketrans(
    "٠١٢٣٤٥٦٧٨٩",
    "0123456789"
)


def normalize_text(text):
    if not text:
        return ""

    text = str(text).strip()
    text = text.translate(ARABIC_DIGITS)

    replacements = {
        "أ": "ا",
        "إ": "ا",
        "آ": "ا",
        "ة": "ه",
        "ى": "ي",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    text = re.sub(r"\s+", " ", text)

    return text.strip().lower()


# ============================================================
# ARABIC NUMBERS
# ============================================================

ARABIC_NUMBERS = {
    "صفر": 0,
    "واحد": 1,
    "واحده": 1,
    "اثنان": 2,
    "اثنين": 2,
    "اثنتان": 2,
    "اثنتين": 2,
    "ثلاثه": 3,
    "ثلاث": 3,
    "اربعه": 4,
    "اربع": 4,
    "خمسه": 5,
    "خمس": 5,
    "سته": 6,
    "ست": 6,
    "سبعه": 7,
    "سبع": 7,
    "ثمانيه": 8,
    "ثمان": 8,
    "تسعه": 9,
    "تسع": 9,
    "عشر": 10,
    "عشره": 10,

    "احد عشر": 11,
    "اثنا عشر": 12,
    "اثني عشر": 12,
    "ثلاثه عشر": 13,
    "اربعه عشر": 14,
    "خمسه عشر": 15,
    "سته عشر": 16,
    "سبعه عشر": 17,
    "ثمانيه عشر": 18,
    "تسعه عشر": 19,

    "عشرون": 20,
    "ثلاثون": 30,
    "اربعون": 40,
    "خمسون": 50,
    "ستون": 60,
    "سبعون": 70,
    "ثمانون": 80,
    "تسعون": 90,
}


def arabic_number_from_text(text):
    text = normalize_text(text)

    if text in ARABIC_NUMBERS:
        return ARABIC_NUMBERS[text]

    # مثال: خمسة عشر
    words = text.split()

    if len(words) == 2:
        a = ARABIC_NUMBERS.get(words[0])
        b = ARABIC_NUMBERS.get(words[1])

        if a is not None and b is not None:
            if b == 10:
                return a + 10

            if a in (20, 30, 40, 50, 60, 70, 80, 90):
                return a + b

    return None


def number_from_text(text):
    if not text:
        return None

    text = str(text).translate(ARABIC_DIGITS)

    # رقم عشري أو صحيح
    match = re.search(r"(?<!\d)(\d+(?:[.,]\d+)?)(?!\d)", text)

    if match:
        try:
            return float(match.group(1).replace(",", "."))
        except ValueError:
            pass

    normalized = normalize_text(text)

    # أطول العبارات أولًا
    for phrase in sorted(
        ARABIC_NUMBERS.keys(),
        key=len,
        reverse=True
    ):
        if phrase in normalized:
            return float(ARABIC_NUMBERS[phrase])

    return None


# ============================================================
# DATE PARSING
# ============================================================

def parse_date_from_text(text):
    if not text:
        return None

    normalized = normalize_text(text)
    current = today_date()

    # اليوم
    if re.search(r"\bاليوم\b", normalized):
        return current.isoformat()

    # أمس / امبارح / البارحة
    if (
        "امس" in normalized
        or "امبارح" in normalized
        or "البارحه" in normalized
    ):
        return (current - timedelta(days=1)).isoformat()

    # أول أمس
    if "اول امس" in normalized:
        return (current - timedelta(days=2)).isoformat()

    # YYYY-MM-DD
    match = re.search(
        r"\b(20\d{2})[-/](\d{1,2})[-/](\d{1,2})\b",
        normalized
    )

    if match:
        try:
            year = int(match.group(1))
            month = int(match.group(2))
            day = int(match.group(3))

            return datetime(
                year, month, day
            ).date().isoformat()
        except ValueError:
            return None

    # DD/MM/YYYY أو DD-MM-YYYY
    match = re.search(
        r"\b(\d{1,2})[-/](\d{1,2})[-/](20\d{2})\b",
        normalized
    )

    if match:
        try:
            day = int(match.group(1))
            month = int(match.group(2))
            year = int(match.group(3))

            return datetime(
                year, month, day
            ).date().isoformat()
        except ValueError:
            return None

    # DD/MM أو DD-MM
    match = re.search(
        r"(?<!\d)(\d{1,2})[-/](\d{1,2})(?!\d)",
        normalized
    )

    if match:
        try:
            day = int(match.group(1))
            month = int(match.group(2))

            return datetime(
                current.year,
                month,
                day
            ).date().isoformat()
        except ValueError:
            return None

    return None


# ============================================================
# YES / NO
# ============================================================

YES_WORDS = {
    "نعم",
    "احفظ",
    "حفظ",
    "موافق",
    "تأكيد",
    "تاكيد",
    "صحيح",
    "نعم احفظ",
    "موافق احفظ",
    "/yes",
}

NO_WORDS = {
    "لا",
    "الغاء",
    "الغاء العملية",
    "إلغاء",
    "إلغاء العملية",
    "لا تحفظ",
    "/cancel",
}


def is_yes(text):
    return normalize_text(text) in {
        normalize_text(x) for x in YES_WORDS
    }


def is_no(text):
    return normalize_text(text) in {
        normalize_text(x) for x in NO_WORDS
    }


# ============================================================
# TELEGRAM
# ============================================================

TELEGRAM_API = (
    f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}"
)


async def telegram_request(method, payload=None, timeout=40):
    if not TELEGRAM_BOT_TOKEN:
        raise RuntimeError("TELEGRAM_BOT_TOKEN غير موجود")

    url = f"{TELEGRAM_API}/{method}"

    async with httpx.AsyncClient(timeout=timeout) as client:
        response = await client.post(
            url,
            json=payload or {}
        )

        response.raise_for_status()
        return response.json()


async def send_message(chat_id, text):
    chunks = split_message(text)

    for chunk in chunks:
        await telegram_request(
            "sendMessage",
            {
                "chat_id": chat_id,
                "text": chunk,
            }
        )


def split_message(text):
    if len(text) <= MAX_TELEGRAM_TEXT:
        return [text]

    chunks = []
    current = ""

    for line in text.splitlines(True):
        if len(current) + len(line) <= MAX_TELEGRAM_TEXT:
            current += line
        else:
            if current:
                chunks.append(current)

            current = line

    if current:
        chunks.append(current)

    return chunks


# ============================================================
# SESSION
# ============================================================

def save_session(chat_id, step, data):
    conn = db_connect()

    conn.execute("""
        INSERT INTO sessions(chat_id, step, data, updated_at)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(chat_id)
        DO UPDATE SET
            step = excluded.step,
            data = excluded.data,
            updated_at = excluded.updated_at
    """, (
        chat_id,
        step,
        json.dumps(data, ensure_ascii=False),
        datetime.utcnow().isoformat()
    ))

    conn.commit()
    conn.close()


def get_session(chat_id):
    conn = db_connect()

    row = conn.execute(
        "SELECT * FROM sessions WHERE chat_id = ?",
        (chat_id,)
    ).fetchone()

    conn.close()

    if not row:
        return None

    try:
        data = json.loads(row["data"])
    except Exception:
        data = {}

    return {
        "step": row["step"],
        "data": data
    }


def clear_session(chat_id):
    conn = db_connect()

    conn.execute(
        "DELETE FROM sessions WHERE chat_id = ?",
        (chat_id,)
    )

    conn.commit()
    conn.close()


# ============================================================
# WORKERS
# ============================================================

def get_active_workers():
    conn = db_connect()

    rows = conn.execute("""
        SELECT *
        FROM workers
        WHERE active = 1
        ORDER BY id
    """).fetchall()

    conn.close()

    return rows


def add_worker(name, phone=None):
    name = name.strip()

    if not name:
        return None

    conn = db_connect()

    cur = conn.execute("""
        INSERT INTO workers(
            name,
            phone,
            hourly_rate,
            active,
            created_at
        )
        VALUES (?, ?, ?, 1, ?)
    """, (
        name,
        phone,
        FIXED_HOURLY_RATE,
        datetime.utcnow().isoformat()
    ))

    conn.commit()

    worker_id = cur.lastrowid

    conn.close()

    return worker_id


def deactivate_worker(worker_id):
    conn = db_connect()

    conn.execute(
        "UPDATE workers SET active = 0 WHERE id = ?",
        (worker_id,)
    )

    conn.commit()
    conn.close()


# ============================================================
# AI EXTRACTION
# ============================================================

AI_SCHEMA = {
    "date": None,
    "workers_count": None,
    "workers_names": [],
    "total_hours": None,
    "breakfast_per_worker": None,
    "breakfast_total": None,
    "other_expenses": None,
    "notes": None,
}


def extract_json_from_response(text):
    if not text:
        return None

    text = text.strip()

    # إزالة markdown fences
    text = re.sub(
        r"^```(?:json)?\s*",
        "",
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r"\s*```$",
        "",
        text
    )

    try:
        return json.loads(text)
    except Exception:
        pass

    # البحث عن أول JSON object
    start = text.find("{")
    end = text.rfind("}")

    if start >= 0 and end > start:
        try:
            return json.loads(
                text[start:end + 1]
            )
        except Exception:
            pass

    return None


async def ai_extract(text):
    if not OPENROUTER_API_KEY:
        return None

    system_prompt = """
أنت نظام استخراج بيانات لمزرعة.

استخرج البيانات من رسالة المستخدم فقط.
لا تخمن أي معلومة غير موجودة بوضوح.

أعد JSON فقط.

الشكل المطلوب:

{
  "date": "YYYY-MM-DD أو null",
  "workers_count": number أو null,
  "workers_names": [],
  "total_hours": number أو null,
  "breakfast_per_worker": number أو null,
  "breakfast_total": number أو null,
  "other_expenses": number أو null,
  "notes": "string أو null"
}

قواعد مهمة:

1. أسماء العمال اختيارية.
2. إذا ذكر المستخدم أسماء العمال استخرجها.
3. إذا لم يذكر أسماء العمال، اجعل workers_names = [] ولا تطلب الأسماء.
4. إذا ذكر عدد العمال فقط، استخرج workers_count.
5. إذا قال "كل واحد 8 ساعات" أو "كل عامل 8 ساعات"،
   total_hours = 8.
   هذا يعني ساعات لكل عامل.
6. التاريخ:
   - اليوم = تاريخ اليوم.
   - أمس / امبارح / البارحة = اليوم السابق.
   - التاريخ الصريح استخرجه.
   - إذا لم يذكر تاريخًا، date = null.
7. الفطور:
   - "فطور لكل عامل 1 دينار"
     => breakfast_per_worker = 1
   - "بدل فطور لكل عامل دينار"
     => breakfast_per_worker = 1 إذا كانت القيمة واضحة.
   - إذا ذكر مجموع الفطور صراحة:
     => breakfast_total.
   - لا تضع نفس المبلغ في الحقلين.
8. مصاريف مثل:
   بنزين، مواصلات، شراء، صرفنا، مصاريف عامة
   تذهب إلى other_expenses.
9. إذا لم توجد مصاريف، استخدم null.
10. لا تعتبر أجر العمال مصروفًا.
11. لا تعتبر عدد العمال مبلغًا.
12. لا تعتبر الساعات مبلغًا.
13. لا تخترع أسماء.
14. إذا كانت الأسماء أقل من عدد العمال، لا تخترع بقية الأسماء.
15. إذا كانت الأسماء موجودة، workers_count يمكن استخراجه منها.
16. إذا قال المستخدم "5 عمال، أحمد ومحمد وخالد"
   workers_count = 5
   workers_names = ["أحمد","محمد","خالد"]
17. لا تغير أسماء العمال ولا تترجمها.
"""

    payload = {
        "model": OPENROUTER_MODEL,
        "messages": [
            {
                "role": "system",
                "content": system_prompt
            },
            {
                "role": "user",
                "content": text
            }
        ],
        "temperature": 0,
        "max_tokens": 1000,
        "response_format": {
            "type": "json_object"
        }
    }

    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
    }

    async with httpx.AsyncClient(timeout=45) as client:

        try:
            response = await client.post(
                "https://openrouter.ai/api/v1/chat/completions",
                headers=headers,
                json=payload
            )

            if response.status_code == 400:
                # بعض النماذج لا تدعم response_format
                payload.pop("response_format", None)

                response = await client.post(
                    "https://openrouter.ai/api/v1/chat/completions",
                    headers=headers,
                    json=payload
                )

            response.raise_for_status()

            result = response.json()

            content = (
                result.get("choices", [{}])[0]
                .get("message", {})
                .get("content", "")
            )

            return extract_json_from_response(content)

        except Exception as exc:
            logger.warning(
                "AI extraction failed: %s",
                exc
            )

            return None


# ============================================================
# DETERMINISTIC EXTRACTION
# ============================================================

def extract_worker_count(text):
    normalized = normalize_text(text)

    # 5 عمال
    match = re.search(
        r"(\d+(?:[.,]\d+)?)\s*(?:عامل|عمال|اشخاص|شخص|عاملين)",
        normalized
    )

    if match:
        return int(float(match.group(1)))

    # خمسة عمال
    for phrase, number in sorted(
        ARABIC_NUMBERS.items(),
        key=lambda x: len(x[0]),
        reverse=True
    ):
        if re.search(
            rf"\b{re.escape(phrase)}\s+(?:عامل|عمال|عاملين)\b",
            normalized
        ):
            return number

    return None


def extract_hours(text):
    normalized = normalize_text(text)

    patterns = [
        r"(\d+(?:[.,]\d+)?)\s*(?:ساعه|ساعات|ساعة)",
        r"(?:كل واحد|كل عامل|لكل عامل|للعامل).*?(\d+(?:[.,]\d+)?)\s*(?:ساعه|ساعات)",
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            normalized
        )

        if match:
            try:
                return float(
                    match.group(1).replace(",", ".")
                )
            except Exception:
                pass

    # كلمات مثل "كل واحد ثمان ساعات"
    for phrase, number in sorted(
        ARABIC_NUMBERS.items(),
        key=lambda x: len(x[0]),
        reverse=True
    ):
        pattern = (
            rf"(?:كل واحد|كل عامل|لكل عامل)"
            rf".*?{re.escape(phrase)}\s+"
            rf"(?:ساعه|ساعات)"
        )

        if re.search(pattern, normalized):
            return float(number)

    return None


def extract_breakfast_per_worker(text):
    normalized = normalize_text(text)

    breakfast_terms = (
        r"(?:فطور|الفطور|بدل فطور|بدل الفطور|وجبه فطور)"
    )

    worker_terms = (
        r"(?:لكل عامل|للعامل|لكل واحد|للعامل الواحد)"
    )

    pattern = (
        breakfast_terms
        + r".*?"
        + worker_terms
        + r".*?"
        + r"(\d+(?:[.,]\d+)?)"
    )

    match = re.search(pattern, normalized)

    if match:
        return float(
            match.group(1).replace(",", ".")
        )

    # الشكل العكسي:
    # لكل عامل فطور 1 دينار
    pattern = (
        worker_terms
        + r".*?"
        + breakfast_terms
        + r".*?"
        + r"(\d+(?:[.,]\d+)?)"
    )

    match = re.search(pattern, normalized)

    if match:
        return float(
            match.group(1).replace(",", ".")
        )

    return None


def extract_breakfast_total(text):
    normalized = normalize_text(text)

    patterns = [
        r"(?:مصاريف|مجموع|اجمالي|اجمالي مصاريف).*?"
        r"(?:فطور|الفطور).*?"
        r"(\d+(?:[.,]\d+)?)",

        r"(?:فطور|الفطور).*?"
        r"(?:مجموع|اجمالي).*?"
        r"(\d+(?:[.,]\d+)?)",
    ]

    for pattern in patterns:
        match = re.search(pattern, normalized)

        if match:
            try:
                return float(
                    match.group(1).replace(",", ".")
                )
            except Exception:
                pass

    return None


def extract_other_expenses(text):
    normalized = normalize_text(text)

    # أولًا نبحث عن مصاريف واضحة لا تخص الفطور
    patterns = [
        r"(?:بنزين|مواصلات|شراء|مشتريات|صرفنا|دفعت|مصروف|مصاريف عامة)"
        r".*?(\d+(?:[.,]\d+)?)",

        r"(?:مصاريف).*?(\d+(?:[.,]\d+)?)",
    ]

    for pattern in patterns:
        matches = list(
            re.finditer(pattern, normalized)
        )

        for match in matches:
            fragment = normalized[
                max(0, match.start() - 30):
                min(len(normalized), match.end() + 30)
            ]

            # لا نحسب مصروف الفطور كمصروف آخر
            if (
                "فطور" in fragment
                or "الفطور" in fragment
            ):
                continue

            try:
                return float(
                    match.group(1).replace(",", ".")
                )
            except Exception:
                pass

    return None


def extract_names(text):
    normalized = normalize_text(text)

    # نستخدم النص الأصلي حتى نحافظ على شكل الأسماء
    original = text

    patterns = [
        r"(?:العمال|عمال|اسماء العمال|اسماء)\s*[:：]\s*(.+)",
    ]

    names_part = None

    for pattern in patterns:
        match = re.search(
            pattern,
            original,
            flags=re.IGNORECASE
        )

        if match:
            names_part = match.group(1)
            break

    if names_part:
        names_part = re.split(
            r"\b(?:كل واحد|كل عامل|ساعات|ساعة|ساعه|فطور|مصاريف|بنزين)\b",
            names_part,
            maxsplit=1,
            flags=re.IGNORECASE
        )[0]

        raw_names = re.split(
            r"[,،&و]+",
            names_part
        )

        names = []

        for name in raw_names:
            name = name.strip(" .:-")

            if name and len(name) <= 60:
                names.append(name)

        return names

    return []


def deterministic_extract(text):
    data = dict(AI_SCHEMA)

    data["date"] = parse_date_from_text(text)
    data["workers_count"] = extract_worker_count(text)
    data["total_hours"] = extract_hours(text)
    data["breakfast_per_worker"] = (
        extract_breakfast_per_worker(text)
    )
    data["breakfast_total"] = (
        extract_breakfast_total(text)
    )
    data["other_expenses"] = (
        extract_other_expenses(text)
    )
    data["workers_names"] = extract_names(text)

    if data["workers_count"] is None:
        if data["workers_names"]:
            data["workers_count"] = len(
                data["workers_names"]
            )

    return data


# ============================================================
# MERGE AI + DETERMINISTIC
# ============================================================

def clean_ai_data(ai_data, fallback):
    if not isinstance(ai_data, dict):
        return fallback

    result = dict(AI_SCHEMA)

    for key in result:
        if key in ai_data:
            result[key] = ai_data[key]

    # التاريخ المحدد حتميًا أهم من AI
    deterministic_date = fallback.get("date")

    if deterministic_date:
        result["date"] = deterministic_date

    # إذا لم يستخرج AI العدد
    if result["workers_count"] is None:
        result["workers_count"] = fallback.get(
            "workers_count"
        )

    if not result.get("workers_names"):
        result["workers_names"] = (
            fallback.get("workers_names") or []
        )

    if result["total_hours"] is None:
        result["total_hours"] = fallback.get(
            "total_hours"
        )

    if result["breakfast_per_worker"] is None:
        result["breakfast_per_worker"] = fallback.get(
            "breakfast_per_worker"
        )

    if result["breakfast_total"] is None:
        result["breakfast_total"] = fallback.get(
            "breakfast_total"
        )

    if result["other_expenses"] is None:
        result["other_expenses"] = fallback.get(
            "other_expenses"
        )

    return result


# ============================================================
# VALIDATION
# ============================================================

def to_float(value):
    if value is None:
        return None

    try:
        return float(value)
    except Exception:
        return None


def validate_data(data):
    count = data.get("workers_count")
    hours = data.get("total_hours")

    if count is not None:
        try:
            count = int(float(count))
        except Exception:
            return False, "عدد العمال غير صحيح."

        data["workers_count"] = count

        if count < 1:
            return False, "عدد العمال يجب أن يكون أكبر من صفر."

        if count > MAX_WORKERS:
            return False, f"الحد الأقصى للعمال هو {MAX_WORKERS}."

    names = data.get("workers_names") or []

    if names:
        names = [
            str(x).strip()
            for x in names
            if str(x).strip()
        ]

        data["workers_names"] = names

        # إذا لم يوجد عدد، نستنتجه من الأسماء
        if count is None:
            count = len(names)
            data["workers_count"] = count

        # لا نرفض إذا الأسماء أقل من العدد
        # لأن الأسماء اختيارية أصلًا.

    if hours is not None:
        hours = to_float(hours)

        if hours is None:
            return False, "عدد الساعات غير صحيح."

        if hours < 0:
            return False, "الساعات لا يمكن أن تكون سالبة."

        if hours > MAX_HOURS_PER_WORKER:
            return False, (
                f"الساعات لكل عامل لا يمكن أن تتجاوز "
                f"{MAX_HOURS_PER_WORKER} ساعة."
            )

        data["total_hours"] = hours

    bpw = to_float(
        data.get("breakfast_per_worker")
    )

    bt = to_float(
        data.get("breakfast_total")
    )

    other = to_float(
        data.get("other_expenses")
    )

    if bpw is not None:
        if bpw < 0 or bpw > MAX_EXPENSE:
            return False, "قيمة بدل الفطور غير صحيحة."

    if bt is not None:
        if bt < 0 or bt > MAX_EXPENSE:
            return False, "إجمالي الفطور غير صحيح."

    if other is not None:
        if other < 0 or other > MAX_EXPENSE:
            return False, "المصاريف الأخرى غير صحيحة."

    # تعارض
    if (
        bpw is not None
        and bpw > 0
        and bt is not None
        and bt > 0
    ):
        return False, (
            "وجدت قيمة للفطور لكل عامل وقيمة إجمالية للفطور "
            "في نفس الوقت. اذكر واحدة فقط."
        )

    data["breakfast_per_worker"] = (
        bpw if bpw is not None else 0.0
    )

    data["breakfast_total"] = (
        bt if bt is not None else 0.0
    )

    data["other_expenses"] = (
        other if other is not None else 0.0
    )

    if not data.get("date"):
        data["date"] = today_str()

    return True, None


# ============================================================
# BUILD WORKERS
# ============================================================

def build_workers(data):
    count = int(data["workers_count"])
    names = data.get("workers_names") or []
    hours = float(data["total_hours"])

    workers = []

    for i in range(count):
        if i < len(names) and names[i]:
            name = str(names[i]).strip()
        else:
            # الاسم اختياري
            name = f"عامل {i + 1}"

        workers.append({
            "name": name,
            "hours": hours,
            "hourly_rate": FIXED_HOURLY_RATE,
            "task": "",
            "notes": "",
        })

    return workers


# ============================================================
# EXPENSE CALCULATION
# ============================================================

def calculate_breakfast(data):
    count = int(data["workers_count"])

    per_worker = float(
        data.get("breakfast_per_worker") or 0
    )

    total = float(
        data.get("breakfast_total") or 0
    )

    if per_worker > 0:
        return per_worker * count

    return total


def calculate_wages(workers):
    return sum(
        float(worker["hours"]) * FIXED_HOURLY_RATE
        for worker in workers
    )


def build_expense_rows(data):
    rows = []

    count = int(data["workers_count"])

    breakfast_per_worker = float(
        data.get("breakfast_per_worker") or 0
    )

    breakfast_total = float(
        data.get("breakfast_total") or 0
    )

    other_expenses = float(
        data.get("other_expenses") or 0
    )

    if breakfast_per_worker > 0:
        rows.append({
            "category": "فطور",
            "amount": breakfast_per_worker * count,
            "notes": (
                f"بدل فطور {breakfast_per_worker:.2f} "
                f"د.أ لكل عامل × {count}"
            )
        })

    elif breakfast_total > 0:
        rows.append({
            "category": "فطور",
            "amount": breakfast_total,
            "notes": "إجمالي بدل الفطور"
        })

    if other_expenses > 0:
        rows.append({
            "category": "مصاريف أخرى",
            "amount": other_expenses,
            "notes": ""
        })

    return rows


# ============================================================
# CONFIRMATION
# ============================================================

def build_confirmation_data(data):
    valid, error = validate_data(data)

    if not valid:
        return None, error

    workers = build_workers(data)

    wages = calculate_wages(workers)

    breakfast = calculate_breakfast(data)

    other = float(
        data.get("other_expenses") or 0
    )

    expenses_total = breakfast + other

    grand_total = wages + expenses_total

    result = dict(data)

    result["workers"] = workers
    result["wages"] = wages
    result["breakfast"] = breakfast
    result["other_expenses"] = other
    result["expenses_total"] = expenses_total
    result["grand_total"] = grand_total

    return result, None


def build_confirmation_message(data):
    result, error = build_confirmation_data(data)

    if error:
        return f"❌ {error}"

    workers = result["workers"]

    lines = [
        "📋 مراجعة اليومية",
        "",
        f"📅 التاريخ: {result['date']}",
        f"👷 عدد العمال: {len(workers)}",
        "",
    ]

    for index, worker in enumerate(workers, start=1):
        wage = (
            worker["hours"]
            * FIXED_HOURLY_RATE
        )

        lines.append(
            f"{index}. {worker['name']} — "
            f"{worker['hours']:g} ساعة × "
            f"{FIXED_HOURLY_RATE:.2f} = "
            f"{wage:.2f} د.أ"
        )

    lines.extend([
        "",
        f"💵 أجور العمال: "
        f"{result['wages']:.2f} د.أ",
    ])

    if result["breakfast"] > 0:
        per_worker = result.get(
            "breakfast_per_worker", 0
        )

        if per_worker > 0:
            lines.append(
                f"🍳 بدل الفطور: "
                f"{per_worker:.2f} × "
                f"{len(workers)} = "
                f"{result['breakfast']:.2f} د.أ"
            )
        else:
            lines.append(
                f"🍳 إجمالي الفطور: "
                f"{result['breakfast']:.2f} د.أ"
            )

    if result["other_expenses"] > 0:
        lines.append(
            f"💰 مصاريف أخرى: "
            f"{result['other_expenses']:.2f} د.أ"
        )

    lines.extend([
        f"💸 مجموع المصاريف: "
        f"{result['expenses_total']:.2f} د.أ",
        f"🏆 الإجمالي النهائي: "
        f"{result['grand_total']:.2f} د.أ",
        "",
        f"ℹ️ سعر الساعة ثابت: "
        f"{FIXED_HOURLY_RATE:.2f} د.أ",
        "ℹ️ الساعات محسوبة لكل عامل.",
    ])

    if result.get("notes"):
        lines.extend([
            "",
            f"📝 ملاحظات: {result['notes']}"
        ])

    lines.extend([
        "",
        "هل البيانات صحيحة؟",
        "اكتب: نعم للحفظ",
        "أو: تعديل",
        "أو: إلغاء"
    ])

    return "\n".join(lines)


# ============================================================
# DATABASE DAILY LOG
# ============================================================

def daily_exists(date):
    conn = db_connect()

    row = conn.execute(
        "SELECT id FROM daily_logs WHERE date = ?",
        (date,)
    ).fetchone()

    conn.close()

    return row is not None


def save_daily_data(data):
    result, error = build_confirmation_data(data)

    if error:
        return False, error

    date = result["date"]

    if daily_exists(date):
        return False, (
            f"يوجد سجل محفوظ مسبقًا بتاريخ {date}.\n"
            "لن أستبدله تلقائيًا حتى لا تضيع البيانات."
        )

    conn = db_connect()

    try:
        now = datetime.utcnow().isoformat()

        cursor = conn.execute("""
            INSERT INTO daily_logs(
                date,
                notes,
                created_at,
                updated_at
            )
            VALUES (?, ?, ?, ?)
        """, (
            date,
            result.get("notes") or "",
            now,
            now
        ))

        daily_log_id = cursor.lastrowid

        for worker in result["workers"]:

            # إذا كان الاسم "عامل X"، لا ننشئ عاملًا في قائمة العمال.
            worker_name = worker["name"]

            worker_id = None

            existing = conn.execute("""
                SELECT id
                FROM workers
                WHERE name = ?
                  AND active = 1
                LIMIT 1
            """, (
                worker_name,
            )).fetchone()

            if existing:
                worker_id = existing["id"]

            conn.execute("""
                INSERT INTO work_logs(
                    daily_log_id,
                    worker_id,
                    hours,
                    task,
                    notes,
                    worker_name
                )
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                daily_log_id,
                worker_id,
                worker["hours"],
                worker.get("task", ""),
                worker.get("notes", ""),
                worker_name,
            ))

        expense_rows = build_expense_rows(result)

        for expense in expense_rows:
            conn.execute("""
                INSERT INTO expenses(
                    daily_log_id,
                    category,
                    amount,
                    notes
                )
                VALUES (?, ?, ?, ?)
            """, (
                daily_log_id,
                expense["category"],
                expense["amount"],
                expense["notes"],
            ))

        conn.commit()

        return True, result

    except Exception as exc:
        conn.rollback()

        logger.exception(
            "Failed to save daily data"
        )

        return False, str(exc)

    finally:
        conn.close()


# ============================================================
# DAILY REPORT DATA
# ============================================================

def get_daily_log(date):
    conn = db_connect()

    daily = conn.execute("""
        SELECT *
        FROM daily_logs
        WHERE date = ?
    """, (
        date,
    )).fetchone()

    if not daily:
        conn.close()
        return None

    workers = conn.execute("""
        SELECT *
        FROM work_logs
        WHERE daily_log_id = ?
        ORDER BY id
    """, (
        daily["id"],
    )).fetchall()

    expenses = conn.execute("""
        SELECT *
        FROM expenses
        WHERE daily_log_id = ?
        ORDER BY id
    """, (
        daily["id"],
    )).fetchall()

    conn.close()

    return {
        "daily": daily,
        "workers": workers,
        "expenses": expenses,
    }


def get_logs_between(date_from, date_to):
    conn = db_connect()

    rows = conn.execute("""
        SELECT *
        FROM daily_logs
        WHERE date BETWEEN ? AND ?
        ORDER BY date
    """, (
        date_from,
        date_to,
    )).fetchall()

    result = []

    for daily in rows:

        workers = conn.execute("""
            SELECT *
            FROM work_logs
            WHERE daily_log_id = ?
            ORDER BY id
        """, (
            daily["id"],
        )).fetchall()

        expenses = conn.execute("""
            SELECT *
            FROM expenses
            WHERE daily_log_id = ?
            ORDER BY id
        """, (
            daily["id"],
        )).fetchall()

        result.append({
            "daily": daily,
            "workers": workers,
            "expenses": expenses,
        })

    conn.close()

    return result


# ============================================================
# REPORT CALCULATIONS
# ============================================================

def calculate_daily_total(log):
    wages = sum(
        float(row["hours"]) * FIXED_HOURLY_RATE
        for row in log["workers"]
    )

    breakfast = sum(
        float(row["amount"])
        for row in log["expenses"]
        if normalize_text(row["category"]) == "فطور"
    )

    other = sum(
        float(row["amount"])
        for row in log["expenses"]
        if normalize_text(row["category"]) != "فطور"
    )

    return {
        "wages": wages,
        "breakfast": breakfast,
        "other": other,
        "expenses": breakfast + other,
        "total": wages + breakfast + other,
    }


def format_daily_summary(log):
    date = log["daily"]["date"]

    totals = calculate_daily_total(log)

    lines = [
        "📊 تقرير اليومية",
        "",
        f"📅 التاريخ: {date}",
        f"👷 عدد العمال: {len(log['workers'])}",
        "",
    ]

    for i, worker in enumerate(
        log["workers"],
        start=1
    ):
        name = (
            worker["worker_name"]
            or f"عامل {i}"
        )

        hours = float(worker["hours"])

        wage = (
            hours * FIXED_HOURLY_RATE
        )

        lines.append(
            f"{i}. {name}: "
            f"{hours:g} ساعة = "
            f"{wage:.2f} د.أ"
        )

    lines.extend([
        "",
        f"💵 الأجور: "
        f"{totals['wages']:.2f} د.أ",
        f"🍳 الفطور: "
        f"{totals['breakfast']:.2f} د.أ",
        f"💰 المصاريف الأخرى: "
        f"{totals['other']:.2f} د.أ",
        f"💸 مجموع المصاريف: "
        f"{totals['expenses']:.2f} د.أ",
        f"🏆 الإجمالي: "
        f"{totals['total']:.2f} د.أ",
        "",
        f"سعر الساعة الثابت: "
        f"{FIXED_HOURLY_RATE:.2f} د.أ",
    ])

    if log["daily"]["notes"]:
        lines.extend([
            "",
            f"📝 {log['daily']['notes']}"
        ])

    return "\n".join(lines)


# ============================================================
# EXCEL REPORT
# ============================================================

def generate_excel_report(logs):
    wb = Workbook()

    ws = wb.active
    ws.title = "التقرير"

    ws.append([
        "التاريخ",
        "العامل",
        "الساعات",
        "سعر الساعة",
        "الأجر",
        "المهمة",
        "ملاحظات",
    ])

    for log in logs:
        date = log["daily"]["date"]

        for i, worker in enumerate(
            log["workers"],
            start=1
        ):
            name = (
                worker["worker_name"]
                or f"عامل {i}"
            )

            hours = float(worker["hours"])
            wage = hours * FIXED_HOURLY_RATE

            ws.append([
                date,
                name,
                hours,
                FIXED_HOURLY_RATE,
                wage,
                worker["task"] or "",
                worker["notes"] or "",
            ])

    expenses_ws = wb.create_sheet("المصاريف")

    expenses_ws.append([
        "التاريخ",
        "التصنيف",
        "المبلغ",
        "ملاحظات",
    ])

    for log in logs:
        for expense in log["expenses"]:
            expenses_ws.append([
                log["daily"]["date"],
                expense["category"],
                expense["amount"],
                expense["notes"] or "",
            ])

    summary_ws = wb.create_sheet("الملخص")

    summary_ws.append([
        "التاريخ",
        "الأجور",
        "الفطور",
        "مصاريف أخرى",
        "مجموع المصاريف",
        "الإجمالي",
    ])

    for log in logs:
        totals = calculate_daily_total(log)

        summary_ws.append([
            log["daily"]["date"],
            totals["wages"],
            totals["breakfast"],
            totals["other"],
            totals["expenses"],
            totals["total"],
        ])

    output = BytesIO()

    wb.save(output)

    output.seek(0)

    return output.getvalue()


# ============================================================
# PDF REPORT
# ============================================================

def generate_pdf_report(logs):
    try:
        from weasyprint import HTML
    except Exception as exc:
        raise RuntimeError(
            "WeasyPrint غير مثبت أو غير متاح."
        ) from exc

    rows_html = ""

    for log in logs:
        date = html.escape(
            str(log["daily"]["date"])
        )

        totals = calculate_daily_total(log)

        rows_html += f"""
        <h2>التاريخ: {date}</h2>

        <table>
            <tr>
                <th>العامل</th>
                <th>الساعات</th>
                <th>سعر الساعة</th>
                <th>الأجر</th>
            </tr>
        """

        for i, worker in enumerate(
            log["workers"],
            start=1
        ):
            name = html.escape(
                str(
                    worker["worker_name"]
                    or f"عامل {i}"
                )
            )

            hours = float(worker["hours"])
            wage = hours * FIXED_HOURLY_RATE

            rows_html += f"""
            <tr>
                <td>{name}</td>
                <td>{hours:g}</td>
                <td>{FIXED_HOURLY_RATE:.2f}</td>
                <td>{wage:.2f}</td>
            </tr>
            """

        rows_html += f"""
        </table>

        <p>
        الأجور: {totals['wages']:.2f} د.أ<br>
        الفطور: {totals['breakfast']:.2f} د.أ<br>
        المصاريف الأخرى: {totals['other']:.2f} د.أ<br>
        مجموع المصاريف: {totals['expenses']:.2f} د.أ<br>
        <strong>الإجمالي: {totals['total']:.2f} د.أ</strong>
        </p>
        """

    document = f"""
    <!DOCTYPE html>
    <html dir="rtl">
    <head>
        <meta charset="utf-8">
        <style>
            body {{
                font-family: DejaVu Sans, sans-serif;
                direction: rtl;
                padding: 30px;
            }}

            h1 {{
                text-align: center;
            }}

            table {{
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 20px;
            }}

            th, td {{
                border: 1px solid #999;
                padding: 8px;
                text-align: center;
            }}

            th {{
                font-weight: bold;
            }}
        </style>
    </head>

    <body>
        <h1>تقرير المزرعة</h1>

        <p>
            سعر الساعة الثابت:
            {FIXED_HOURLY_RATE:.2f} د.أ
        </p>

        {rows_html}

    </body>
    </html>
    """

    return HTML(
        string=document
    ).write_pdf()


# ============================================================
# TELEGRAM COMMANDS
# ============================================================

async def show_workers(chat_id):
    workers = get_active_workers()

    if not workers:
        await send_message(
            chat_id,
            "لا يوجد عمال محفوظون حاليًا."
        )
        return

    lines = [
        "👷 قائمة العمال:",
        ""
    ]

    for i, worker in enumerate(
        workers,
        start=1
    ):
        lines.append(
            f"{i}. {worker['name']}"
        )

    lines.extend([
        "",
        f"💵 سعر الساعة ثابت: "
        f"{FIXED_HOURLY_RATE:.2f} د.أ"
    ])

    await send_message(
        chat_id,
        "\n".join(lines)
    )


async def show_help(chat_id):
    text = f"""
🤖 إدارة عمال المزرعة

يمكنك إرسال اليومية بشكل طبيعي.

مثال:

أمس 5 عمال، كل واحد 8 ساعات،
فطور لكل عامل 1 دينار،
وبنزين 5 دنانير

وسأحسب:

👷 5 عمال
⏱️ 8 ساعات لكل عامل
💵 الساعة = {FIXED_HOURLY_RATE:.2f} د.أ
🍳 الفطور = 5 د.أ
⛽ البنزين = 5 د.أ

أسماء العمال اختيارية.

مثال بأسماء:

اليوم أحمد ومحمد وخالد،
كل واحد 8 ساعات،
فطور لكل عامل 1 دينار

إذا لم تذكر أسماء، سأستخدم:
عامل 1، عامل 2، عامل 3...

الأوامر:

/start
/help
/workers
/report
/report_today
/report_week
/cancel

يمكنك أيضًا كتابة:
"تقرير اليوم"
"تقرير أمس"
"تقرير من 1/9 إلى 4/9"
""".strip()

    await send_message(
        chat_id,
        text
    )


async def show_today_report(chat_id):
    date = today_str()

    log = get_daily_log(date)

    if not log:
        await send_message(
            chat_id,
            f"لا يوجد سجل بتاريخ {date}."
        )
        return

    await send_message(
        chat_id,
        format_daily_summary(log)
    )


async def show_report_for_date(chat_id, date):
    log = get_daily_log(date)

    if not log:
        await send_message(
            chat_id,
            f"لا يوجد سجل بتاريخ {date}."
        )
        return

    await send_message(
        chat_id,
        format_daily_summary(log)
    )


# ============================================================
# REPORT COMMAND PARSING
# ============================================================

def parse_report_date(text):
    date = parse_date_from_text(text)

    if date:
        return date

    return None


# ============================================================
# START NEW DAILY ENTRY
# ============================================================

async def start_daily_entry(chat_id, text):
    fallback = deterministic_extract(text)

    ai_data = await ai_extract(text)

    data = clean_ai_data(
        ai_data,
        fallback
    )

    valid, error = validate_data(data)

    if not valid:
        await send_message(
            chat_id,
            f"❌ {error}"
        )
        return

    # ========================================================
    # أسماء العمال اختيارية
    # ========================================================

    if data.get("workers_count") is None:
        save_session(
            chat_id,
            "ASK_WORKERS_COUNT",
            data
        )

        await send_message(
            chat_id,
            "👷 كم عدد العمال؟"
        )

        return

    if data.get("total_hours") is None:
        save_session(
            chat_id,
            "ASK_HOURS",
            data
        )

        await send_message(
            chat_id,
            "⏱️ كم ساعة عمل لكل عامل؟"
        )

        return

    # المصاريف اختيارية.
    # لا نطلب أسماء العمال.

    if (
        data.get("breakfast_per_worker", 0) == 0
        and data.get("breakfast_total", 0) == 0
        and data.get("other_expenses", 0) == 0
    ):
        save_session(
            chat_id,
            "ASK_OPTIONAL_EXPENSES",
            data
        )

        await send_message(
            chat_id,
            "💰 هل توجد مصاريف إضافية أو بدل فطور؟\n\n"
            "إذا نعم، اكتبها مثل:\n"
            "فطور لكل عامل 1 دينار، وبنزين 5 دنانير\n\n"
            "وإذا لا توجد مصاريف، اكتب: لا"
        )

        return

    message = build_confirmation_message(data)

    save_session(
        chat_id,
        "CONFIRM",
        data
    )

    await send_message(
        chat_id,
        message
    )


# ============================================================
# SESSION HANDLER
# ============================================================

async def handle_session(chat_id, text, session):
    step = session["step"]
    data = session["data"]

    normalized = normalize_text(text)

    # الإلغاء يعمل في أي مرحلة
    if is_no(text) and step != "CONFIRM":
        clear_session(chat_id)

        await send_message(
            chat_id,
            "❌ تم إلغاء العملية."
        )

        return

    if step == "ASK_WORKERS_COUNT":

        value = number_from_text(text)

        if value is None:
            await send_message(
                chat_id,
                "❌ لم أفهم عدد العمال.\n"
                "مثال: 5 عمال"
            )
            return

        data["workers_count"] = int(value)

        if (
            data["workers_count"] < 1
            or data["workers_count"] > MAX_WORKERS
        ):
            await send_message(
                chat_id,
                f"❌ عدد العمال يجب أن يكون بين 1 و {MAX_WORKERS}."
            )
            return

        data["workers_names"] = (
            data.get("workers_names") or []
        )

        if data.get("total_hours") is None:
            save_session(
                chat_id,
                "ASK_HOURS",
                data
            )

            await send_message(
                chat_id,
                "⏱️ كم ساعة عمل لكل عامل؟"
            )

            return

        save_session(
            chat_id,
            "ASK_OPTIONAL_EXPENSES",
            data
        )

        await send_message(
            chat_id,
            "💰 هل توجد مصاريف إضافية أو بدل فطور؟\n"
            "اكتبها أو اكتب: لا"
        )

        return

    if step == "ASK_HOURS":

        value = number_from_text(text)

        if value is None:
            await send_message(
                chat_id,
                "❌ لم أفهم عدد الساعات.\n"
                "مثال: 8 ساعات"
            )
            return

        value = float(value)

        if value < 0 or value > MAX_HOURS_PER_WORKER:
            await send_message(
                chat_id,
                f"❌ الساعات يجب أن تكون بين 0 و "
                f"{MAX_HOURS_PER_WORKER}."
            )
            return

        data["total_hours"] = value

        save_session(
            chat_id,
            "ASK_OPTIONAL_EXPENSES",
            data
        )

        await send_message(
            chat_id,
            "💰 هل توجد مصاريف إضافية أو بدل فطور؟\n\n"
            "مثال:\n"
            "فطور لكل عامل 1 دينار، وبنزين 5 دنانير\n\n"
            "أو اكتب: لا"
        )

        return

    if step == "ASK_OPTIONAL_EXPENSES":

        if normalize_text(text) in {
            "لا",
            "لا يوجد",
            "لا توجد",
            "ما في",
            "مافي",
            "مافي مصاريف",
        }:
            data["breakfast_per_worker"] = 0
            data["breakfast_total"] = 0
            data["other_expenses"] = 0

            save_session(
                chat_id,
                "CONFIRM",
                data
            )

            await send_message(
                chat_id,
                build_confirmation_message(data)
            )

            return

        fallback = deterministic_extract(text)
        ai_data = await ai_extract(text)

        expense_data = clean_ai_data(
            ai_data,
            fallback
        )

        data["breakfast_per_worker"] = (
            expense_data.get(
                "breakfast_per_worker"
            ) or 0
        )

        data["breakfast_total"] = (
            expense_data.get(
                "breakfast_total"
            ) or 0
        )

        data["other_expenses"] = (
            expense_data.get(
                "other_expenses"
            ) or 0
        )

        valid, error = validate_data(data)

        if not valid:
            await send_message(
                chat_id,
                f"❌ {error}"
            )
            return

        save_session(
            chat_id,
            "CONFIRM",
            data
        )

        await send_message(
            chat_id,
            build_confirmation_message(data)
        )

        return

    if step == "CONFIRM":

        if is_yes(text):

            success, result = save_daily_data(data)

            if not success:
                await send_message(
                    chat_id,
                    f"❌ تعذر الحفظ:\n{result}"
                )
                clear_session(chat_id)
                return

            clear_session(chat_id)

            await send_message(
                chat_id,
                "✅ تم حفظ اليومية بنجاح.\n\n"
                + format_saved_confirmation(result)
            )

            return

        if normalize_text(text) in {
            "تعديل",
            "عدل",
            "تعديل البيانات",
        }:
            clear_session(chat_id)

            await send_message(
                chat_id,
                "✏️ أرسل اليومية من جديد بالبيانات الصحيحة."
            )

            return

        if is_no(text):
            clear_session(chat_id)

            await send_message(
                chat_id,
                "❌ تم إلغاء الحفظ."
            )

            return

        await send_message(
            chat_id,
            "اكتب فقط:\n"
            "✅ نعم للحفظ\n"
            "✏️ تعديل\n"
            "❌ إلغاء"
        )

        return


# ============================================================
# SAVED MESSAGE
# ============================================================

def format_saved_confirmation(data):
    return (
        f"📅 {data['date']}\n"
        f"👷 العمال: {len(data['workers'])}\n"
        f"💵 الأجور: {data['wages']:.2f} د.أ\n"
        f"🍳 الفطور: {data['breakfast']:.2f} د.أ\n"
        f"💰 مصاريف أخرى: {data['other_expenses']:.2f} د.أ\n"
        f"🏆 الإجمالي: {data['grand_total']:.2f} د.أ"
    )


# ============================================================
# GENERAL TEXT COMMANDS
# ============================================================

async def handle_general_text(chat_id, text):

    normalized = normalize_text(text)

    # تقارير
    if normalized in {
        "تقرير",
        "تقرير اليوم",
        "تقرير اليوميه",
        "تقرير اليومية",
        "تقرير اليوم",
    }:
        await show_today_report(chat_id)
        return

    # تقرير أمس
    if (
        normalized == "تقرير امس"
        or normalized == "تقرير امبارح"
        or normalized == "تقرير البارحه"
    ):
        date = (
            today_date()
            - timedelta(days=1)
        ).isoformat()

        await show_report_for_date(
            chat_id,
            date
        )

        return

    # تقرير بتاريخ
    if "تقرير" in normalized:
        date = parse_report_date(text)

        if date:
            await show_report_for_date(
                chat_id,
                date
            )
            return

    # العمال
    if normalized in {
        "العمال",
        "قائمة العمال",
        "اسماء العمال",
        "اسماء",
    }:
        await show_workers(chat_id)
        return

    # بداية تسجيل واضحة
    await start_daily_entry(
        chat_id,
        text
    )


# ============================================================
# COMMAND HANDLER
# ============================================================

async def handle_command(chat_id, text):

    command = text.strip().split()[0].lower()

    if command == "/start":
        await send_message(
            chat_id,
            "👋 أهلاً بك في نظام إدارة عمال المزرعة.\n\n"
            "أرسل اليومية بشكل طبيعي.\n\n"
            "مثال:\n"
            "أمس 5 عمال، كل واحد 8 ساعات، "
            "فطور لكل عامل 1 دينار، وبنزين 5 دنانير\n\n"
            "أسماء العمال اختيارية.\n"
            "اكتب /help للمساعدة."
        )
        return True

    if command == "/help":
        await show_help(chat_id)
        return True

    if command == "/cancel":
        clear_session(chat_id)

        await send_message(
            chat_id,
            "❌ تم إلغاء العملية."
        )

        return True

    if command == "/workers":
        await show_workers(chat_id)
        return True

    if command == "/report_today":
        await show_today_report(chat_id)
        return True

    if command == "/report":
        parts = text.strip().split(maxsplit=1)

        if len(parts) == 2:
            date = parse_report_date(
                parts[1]
            )

            if date:
                await show_report_for_date(
                    chat_id,
                    date
                )
                return True

        await show_today_report(chat_id)
        return True

    return False


# ============================================================
# UPDATE HANDLER
# ============================================================

async def handle_update(update):

    message = update.get("message")

    if not message:
        return

    chat = message.get("chat") or {}
    chat_id = chat.get("id")

    if chat_id is None:
        return

    text = message.get("text")

    if not text:
        return

    text = text.strip()

    # الأوامر أولًا
    if text.startswith("/"):
        handled = await handle_command(
            chat_id,
            text
        )

        if handled:
            return

    # الجلسة الحالية
    session = get_session(chat_id)

    if session:

        # help أثناء الجلسة
        if normalize_text(text) in {
            "مساعدة",
            "ساعدني",
            "help",
        }:
            await show_help(chat_id)
            return

        await handle_session(
            chat_id,
            text,
            session
        )

        return

    await handle_general_text(
        chat_id,
        text
    )


# ============================================================
# TELEGRAM POLLING
# ============================================================

async def polling_loop():

    if not TELEGRAM_BOT_TOKEN:
        logger.error(
            "TELEGRAM_BOT_TOKEN غير موجود."
        )
        return

    init_db()

    logger.info(
        "Bot started. Fixed hourly rate = %.2f JD",
        FIXED_HOURLY_RATE
    )

    offset = None

    # إزالة webhook إن وجد
    try:
        await telegram_request(
            "deleteWebhook",
            {
                "drop_pending_updates": False
            }
        )
    except Exception as exc:
        logger.warning(
            "deleteWebhook failed: %s",
            exc
        )

    while True:

        try:
            payload = {
                "timeout": 30,
                "allowed_updates": ["message"],
            }

            if offset is not None:
                payload["offset"] = offset

            result = await telegram_request(
                "getUpdates",
                payload,
                timeout=40
            )

            if not result.get("ok"):
                await asyncio.sleep(3)
                continue

            updates = result.get(
                "result",
                []
            )

            for update in updates:

                update_id = update.get(
                    "update_id"
                )

                if update_id is not None:
                    offset = update_id + 1

                try:
                    await handle_update(update)

                except Exception:
                    logger.exception(
                        "Error handling update"
                    )

        except asyncio.CancelledError:
            raise

        except Exception as exc:
            logger.exception(
                "Polling error: %s",
                exc
            )

            await asyncio.sleep(5)


# ============================================================
# STARTUP
# ============================================================

@app.on_event("startup")
async def startup_event():

    init_db()

    if TELEGRAM_BOT_TOKEN:
        asyncio.create_task(
            polling_loop()
        )
    else:
        logger.warning(
            "TELEGRAM_BOT_TOKEN غير موجود."
        )


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":

    init_db()

    uvicorn.run(
        app,
        host="0.0.0.0",
        port=PORT
    )
