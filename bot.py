import argparse
import asyncio
import datetime
import hashlib
import html
import json
import logging
import os
import re
import shutil
import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

import aiosqlite
import openpyxl
from aiogram import Bot, Dispatcher, F
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ChatMemberStatus, ParseMode
from aiogram.exceptions import TelegramAPIError, TelegramBadRequest, TelegramForbiddenError
from aiogram.filters import CommandStart, ExceptionTypeFilter
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    CallbackQuery,
    ErrorEvent,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Message,
)
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.interval import IntervalTrigger
from dotenv import load_dotenv

# ==============================================================================
# 1. КОНФИГУРАЦИЯ
# ==============================================================================
load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger("olympiad_bot")


def _parse_int_env(name: str, default: int) -> int:
    raw = os.getenv(name, "").strip()
    if not raw:
        return default
    try:
        return int(raw)
    except ValueError:
        logger.warning("Некорректное число в %s, используется значение по умолчанию.", name)
        return default


def normalize_channel_id(raw_id: int) -> int:
    s = str(raw_id)
    if raw_id > 0 and not s.startswith("-100"):
        return int(f"-100{raw_id}")
    return raw_id


BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
ADMIN_ID = _parse_int_env("ADMIN_ID", 5250536869)
RAW_CHANNEL_ID = _parse_int_env("CHANNEL_ID", 3399018036)
CHANNEL_CHAT_ID = normalize_channel_id(RAW_CHANNEL_ID)
CHANNEL_USERNAME = os.getenv("CHANNEL_USERNAME", "amiriskanderscience").lstrip("@").strip()
CHANNEL_URL = os.getenv("CHANNEL_URL", f"https://t.me/{CHANNEL_USERNAME}").strip()

TZ_NAME = os.getenv("TZ", "Europe/Moscow").strip() or "Europe/Moscow"
try:
    APP_TZ = ZoneInfo(TZ_NAME)
except Exception:
    TZ_NAME = "Europe/Moscow"
    APP_TZ = ZoneInfo("Europe/Moscow")

BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = Path(os.getenv("EXCEL_PATH", str(BASE_DIR / "2027.xlsx")))
DB_PATH = Path(os.getenv("DB_PATH", str(BASE_DIR / "olympiads.db")))
BACKUP_DIR = Path(os.getenv("BACKUP_DIR", str(BASE_DIR / "backups")))

MIN_GRADE = 7
MAX_GRADE = 11
GRADES_RANGE = list(range(MIN_GRADE, MAX_GRADE + 1))

REMINDER_OFFSETS = (7, 3, 1, 0)
URGENT_REG_DAYS = 7  # Срочная рассылка, если до конца регистрации осталось < 7 дней (0..6)
DEFAULT_SUBJECTS = ["Химия", "Биология", "Математика", "Литература", "Физика"]
REG_STAGE_CODES = ("registration", "final_reg")

STAGE_COLUMN_MAP = {
    "регистрация": ("registration", "Регистрация", 10),
    "i тур": ("round_1", "I тур", 20),
    "1 тур": ("round_1", "I тур", 20),
    "отборочный этап": ("round_1", "Отборочный этап", 20),
    "регистрация на закл": ("final_reg", "Регистрация на финал", 30),
    "регистрация на заключительный этап": ("final_reg", "Регистрация на финал", 30),
    "ii тур": ("round_2", "II тур / Финал", 40),
    "заключительный этап": ("final_stage", "Заключительный этап", 40),
    "iii тур": ("round_3", "III тур", 50),
    "iv тур": ("round_4", "IV тур", 60),
}

MONTHS_RU = {
    "январ": 1, "феврал": 2, "ферал": 2, "март": 3, "апрел": 4,
    "май": 5, "мая": 5, "июн": 6, "июл": 7, "август": 8,
    "сентябр": 9, "октбяр": 10, "октябр": 10, "ноябр": 11, "декабр": 12,
}

AMBIGUOUS_KEYWORDS = (
    "планируется", "предварительно", "запланировано", "середина",
    "конец", "начало", "в течение", "или", "ручная проверка",
    "дист", "очно", "заочно", "волна", "тур", "по заявкам", "выборочно",
)
# Совпадение только с начала слова, чтобы «или» не срабатывало внутри других слов
AMBIGUOUS_RE = re.compile("|".join(r"(?<!\w)" + re.escape(k) for k in AMBIGUOUS_KEYWORDS))

STAGE_PARTICIPATION_CHOICES = [
    "Регистрация подана",
    "Отборочный этап пройден",
    "Регистрация на финал пройдена",
    "Заключительный этап пройден",
    "Призёр / Победитель",
]


def is_admin(user_id: int | None) -> bool:
    return user_id is not None and int(user_id) == int(ADMIN_ID)


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value.strip().lower().replace("ё", "е"))


def esc(value: Any) -> str:
    """Экранирование текста для parse_mode=HTML (названия, заметки и т.п. могут содержать < > &)."""
    return html.escape("" if value is None else str(value), quote=False)


def join_limited(lines: list[str], limit: int = 3900) -> str:
    """Склеивает строки, не превышая лимит Telegram и не обрезая HTML-теги посередине."""
    out: list[str] = []
    total = 0
    for ln in lines:
        if total + len(ln) + 1 > limit:
            out.append("…")
            break
        out.append(ln)
        total += len(ln) + 1
    return "\n".join(out)


def make_olympiad_key(subject: str, title: str) -> str:
    raw = f"{normalize_text(subject)}||{normalize_text(title)}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def fmt_date_ru(iso_str: str | None) -> str:
    if not iso_str:
        return "не указана"
    try:
        return datetime.date.fromisoformat(iso_str).strftime("%d.%m.%Y")
    except ValueError:
        return esc(iso_str)


def fmt_date_range_ru(start_iso: str | None, end_iso: str | None, raw_text: str) -> str:
    raw = esc(raw_text)
    if start_iso and end_iso:
        if start_iso == end_iso:
            return f"{fmt_date_ru(start_iso)} ({raw})"
        return f"{fmt_date_ru(start_iso)} — {fmt_date_ru(end_iso)} ({raw})"
    if start_iso:
        return f"{fmt_date_ru(start_iso)} ({raw})"
    return f"{raw} [требует проверки]"


# ==============================================================================
# 2. БАЗА ДАННЫХ И МИГРАЦИИ
# ==============================================================================
SCHEMA_SQL = """
PRAGMA foreign_keys = ON;
PRAGMA journal_mode = WAL;

CREATE TABLE IF NOT EXISTS schema_migrations (
    version INTEGER PRIMARY KEY,
    applied_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    telegram_id INTEGER NOT NULL UNIQUE,
    full_name TEXT NOT NULL,
    username TEXT,
    grade INTEGER CHECK (grade IS NULL OR (grade >= 7 AND grade <= 11)),
    selected_subjects_json TEXT NOT NULL DEFAULT '[]',
    is_subscribed INTEGER NOT NULL DEFAULT 0,
    last_sub_check TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_users_telegram_id ON users(telegram_id);

CREATE TABLE IF NOT EXISTS olympiads (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    subject TEXT NOT NULL,
    title TEXT NOT NULL,
    normalized_key TEXT NOT NULL UNIQUE,
    level TEXT,
    grades_raw TEXT,
    grades_json TEXT NOT NULL DEFAULT '[]',
    grades_known INTEGER NOT NULL DEFAULT 0,
    url TEXT,
    notes TEXT,
    source_sheet TEXT NOT NULL,
    source_row INTEGER NOT NULL,
    needs_review INTEGER NOT NULL DEFAULT 0,
    is_active INTEGER NOT NULL DEFAULT 1,
    updated_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_olympiads_subject ON olympiads(subject);
CREATE INDEX IF NOT EXISTS idx_olympiads_active ON olympiads(is_active);

CREATE TABLE IF NOT EXISTS stages (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    olympiad_id INTEGER NOT NULL,
    stage_code TEXT NOT NULL,
    stage_name TEXT NOT NULL,
    sort_order INTEGER NOT NULL DEFAULT 50,
    raw_date_text TEXT NOT NULL,
    parsed_start_date TEXT,
    parsed_end_date TEXT,
    is_ambiguous INTEGER NOT NULL DEFAULT 0,
    needs_review INTEGER NOT NULL DEFAULT 0,
    ambiguity_reason TEXT,
    is_manual INTEGER NOT NULL DEFAULT 0,
    UNIQUE(olympiad_id, stage_code),
    FOREIGN KEY (olympiad_id) REFERENCES olympiads(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_stages_olympiad_id ON stages(olympiad_id);

CREATE TABLE IF NOT EXISTS registrations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    olympiad_id INTEGER NOT NULL,
    created_at TEXT NOT NULL,
    UNIQUE(user_id, olympiad_id),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE,
    FOREIGN KEY (olympiad_id) REFERENCES olympiads(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS progress (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    stage_id INTEGER NOT NULL,
    is_completed INTEGER NOT NULL DEFAULT 0,
    completed_at TEXT,
    updated_at TEXT NOT NULL,
    UNIQUE(user_id, stage_id),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE,
    FOREIGN KEY (stage_id) REFERENCES stages(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS progress_audit (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    stage_id INTEGER NOT NULL,
    olympiad_title TEXT NOT NULL,
    stage_name TEXT NOT NULL,
    new_status INTEGER NOT NULL,
    changed_at TEXT NOT NULL,
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    olympiad_id INTEGER,
    olympiad_title TEXT NOT NULL,
    subject TEXT NOT NULL,
    participation_year INTEGER NOT NULL,
    stage_reached TEXT NOT NULL,
    result_text TEXT,
    note_text TEXT,
    is_final INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL,
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE,
    FOREIGN KEY (olympiad_id) REFERENCES olympiads(id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS reminder_settings (
    user_id INTEGER PRIMARY KEY,
    enabled INTEGER NOT NULL DEFAULT 1,
    updated_at TEXT NOT NULL,
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS reminders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    stage_id INTEGER NOT NULL,
    target_date TEXT NOT NULL,
    offset_days INTEGER NOT NULL,
    sent_at TEXT NOT NULL,
    UNIQUE(user_id, stage_id, target_date, offset_days),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE,
    FOREIGN KEY (stage_id) REFERENCES stages(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS import_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    admin_id INTEGER NOT NULL,
    source_file TEXT NOT NULL,
    backup_path TEXT,
    found_olympiads INTEGER NOT NULL DEFAULT 0,
    imported_stages INTEGER NOT NULL DEFAULT 0,
    found_links INTEGER NOT NULL DEFAULT 0,
    skipped_rows INTEGER NOT NULL DEFAULT 0,
    errors_count INTEGER NOT NULL DEFAULT 0,
    ambiguous_count INTEGER NOT NULL DEFAULT 0,
    review_count INTEGER NOT NULL DEFAULT 0,
    details_json TEXT,
    created_at TEXT NOT NULL
);
"""


async def get_db(db_path: Path = DB_PATH) -> aiosqlite.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = await aiosqlite.connect(str(db_path))
    conn.row_factory = aiosqlite.Row
    await conn.execute("PRAGMA foreign_keys = ON;")
    await conn.execute("PRAGMA journal_mode = WAL;")
    return conn


async def init_db(db_path: Path = DB_PATH) -> None:
    now_iso = datetime.datetime.now(APP_TZ).isoformat()
    conn = await get_db(db_path)
    try:
        await conn.executescript(SCHEMA_SQL)

        cur_cols = await conn.execute("PRAGMA table_info(users)")
        col_names = {r["name"] for r in await cur_cols.fetchall()}
        if "selected_subjects_json" not in col_names:
            await conn.execute("ALTER TABLE users ADD COLUMN selected_subjects_json TEXT NOT NULL DEFAULT '[]'")

        cur_cols = await conn.execute("PRAGMA table_info(stages)")
        col_names = {r["name"] for r in await cur_cols.fetchall()}
        if "is_manual" not in col_names:
            await conn.execute("ALTER TABLE stages ADD COLUMN is_manual INTEGER NOT NULL DEFAULT 0")

        await conn.execute("UPDATE users SET grade = 7 WHERE grade IS NOT NULL AND grade < 7")

        for old_tbl in ("favorites", "google_accounts", "google_events"):
            cur_t = await conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (old_tbl,))
            if await cur_t.fetchone():
                if old_tbl == "favorites":
                    await conn.execute(
                        """
                        INSERT OR IGNORE INTO registrations (user_id, olympiad_id, created_at)
                        SELECT user_id, olympiad_id, created_at FROM favorites
                        """
                    )
                await conn.execute(f"DROP TABLE {old_tbl}")

        cur = await conn.execute("SELECT version FROM schema_migrations WHERE version = 6")
        if not await cur.fetchone():
            await conn.execute(
                "INSERT OR REPLACE INTO schema_migrations (version, applied_at) VALUES (?, ?)",
                (6, now_iso),
            )
        await conn.commit()
    finally:
        await conn.close()


def create_db_backup(db_path: Path = DB_PATH, backup_dir: Path = BACKUP_DIR) -> str | None:
    if not db_path.exists():
        return None
    backup_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.datetime.now(APP_TZ).strftime("%Y%m%d_%H%M%S")
    dest = backup_dir / f"{db_path.stem}_backup_{ts}.db"

    src_conn = sqlite3.connect(str(db_path))
    dst_conn = sqlite3.connect(str(dest))
    try:
        src_conn.backup(dst_conn)
    finally:
        dst_conn.close()
        src_conn.close()
    return str(dest)


# ==============================================================================
# 3. РАЗБОР EXCEL И ИМПОРТ
# ==============================================================================
def _month_from_word(word: str) -> int | None:
    w = normalize_text(word)
    for prefix, m in MONTHS_RU.items():
        if w.startswith(prefix):
            return m
    return None


def _infer_school_year(month: int, explicit_year: int | None) -> tuple[int | None, bool]:
    if explicit_year is not None:
        return explicit_year, False
    if 8 <= month <= 12:
        return 2026, True
    if 1 <= month <= 7:
        return 2027, True
    return None, True


def parse_grades_value(raw_val: Any) -> dict[str, Any]:
    if raw_val is None or str(raw_val).strip() == "":
        return {
            "grades": [],
            "grades_raw": "Не указаны в источнике",
            "grades_known": 0,
            "ambiguous": False,
            "reason": None,
        }

    if isinstance(raw_val, (datetime.datetime, datetime.date)):
        # Excel мог превратить диапазон вида «5-11» в дату (5 ноября / 11 мая и т.п.)
        a, b = sorted((raw_val.month, raw_val.day))
        if 1 <= a <= b <= 11:
            grades = [g for g in range(a, b + 1) if MIN_GRADE <= g <= MAX_GRADE]
            return {
                "grades": grades,
                "grades_raw": f"{a}-{b} (в Excel было автопреобразовано в дату {raw_val:%d.%m.%Y})",
                "grades_known": 1 if grades else 0,
                "ambiguous": True,
                "reason": f"Excel преобразовал диапазон классов {a}-{b} в дату {raw_val:%Y-%m-%d}",
            }
        return {
            "grades": [],
            "grades_raw": str(raw_val),
            "grades_known": 0,
            "ambiguous": True,
            "reason": f"Неоднозначное значение в столбце классов: {raw_val}",
        }

    raw_str = str(raw_val).strip()
    norm = raw_str.replace("–", "-").replace("—", "-")
    grades_set: set[int] = set()
    ambiguous = False
    bad_parts: list[str] = []

    for chunk in re.split(r"[,;]+", norm):
        part = chunk.strip()
        if not part:
            continue
        m_rng = re.fullmatch(r"(\d{1,2})\s*-\s*(\d{1,2})", part)
        if m_rng:
            a, b = int(m_rng.group(1)), int(m_rng.group(2))
            if 1 <= a <= b <= 11:
                for g in range(a, b + 1):
                    if MIN_GRADE <= g <= MAX_GRADE:
                        grades_set.add(g)
            else:
                ambiguous = True
                bad_parts.append(part)
        elif re.fullmatch(r"\d{1,2}", part):
            g = int(part)
            if 1 <= g <= 11:
                if MIN_GRADE <= g <= MAX_GRADE:
                    grades_set.add(g)
            else:
                ambiguous = True
                bad_parts.append(part)
        else:
            ambiguous = True
            bad_parts.append(part)

    sorted_grades = sorted(grades_set)
    return {
        "grades": sorted_grades,
        "grades_raw": raw_str,
        "grades_known": 1 if sorted_grades else 0,
        "ambiguous": ambiguous,
        "reason": f"Нестандартный формат классов: {', '.join(bad_parts)}" if bad_parts else None,
    }


_RE_TWO_MONTHS = re.compile(
    r"(\d{1,2})\s+([а-яё]+)\s*-\s*(\d{1,2})\s+([а-яё]+)(?:\s+(\d{4}))?", re.IGNORECASE
)
_RE_SAME_MONTH = re.compile(r"(\d{1,2})\s*-\s*(\d{1,2})\s+([а-яё]+)(?:\s+(\d{4}))?", re.IGNORECASE)
_RE_SINGLE = re.compile(r"(?:^|[^\d-])(\d{1,2})\s+([а-яё]+)(?:\s+(\d{4}))?", re.IGNORECASE)
_RE_DOT = re.compile(r"\b(\d{1,2})\.(\d{1,2})(?:\.(\d{4}))?\b")


def _parse_text_dates(clean: str) -> tuple[str | None, str | None, bool, str | None]:
    """Возвращает (start_iso, end_iso, год_выведен_автоматически, ошибка)."""
    error = "дата не распознана в надёжном формате"

    m = _RE_TWO_MONTHS.search(clean)
    if m:
        d1_s, m1_w, d2_s, m2_w, y_s = m.groups()
        m1, m2 = _month_from_word(m1_w), _month_from_word(m2_w)
        if m1 and m2:
            y_exp = int(y_s) if y_s else None
            y1, inf1 = _infer_school_year(m1, y_exp)
            y2, inf2 = _infer_school_year(m2, y_exp)
            if y_exp is None and m1 > m2:
                y1, y2 = 2026, 2027
            try:
                return (
                    datetime.date(y1, m1, int(d1_s)).isoformat(),
                    datetime.date(y2, m2, int(d2_s)).isoformat(),
                    inf1 or inf2,
                    None,
                )
            except (ValueError, TypeError) as exc:
                return None, None, False, f"некорректная дата: {exc}"
        error = "не распознан месяц"

    m = _RE_SAME_MONTH.search(clean)
    if m:
        d1_s, d2_s, m_w, y_s = m.groups()
        mon = _month_from_word(m_w)
        if mon:
            y, inf = _infer_school_year(mon, int(y_s) if y_s else None)
            try:
                return (
                    datetime.date(y, mon, int(d1_s)).isoformat(),
                    datetime.date(y, mon, int(d2_s)).isoformat(),
                    inf,
                    None,
                )
            except (ValueError, TypeError) as exc:
                return None, None, False, f"некорректная дата: {exc}"
        error = "не распознан месяц"

    m = _RE_SINGLE.search(clean)
    if m:
        d_s, m_w, y_s = m.groups()
        mon = _month_from_word(m_w)
        if mon:
            y, inf = _infer_school_year(mon, int(y_s) if y_s else None)
            try:
                iso = datetime.date(y, mon, int(d_s)).isoformat()
                return iso, iso, inf, None
            except (ValueError, TypeError) as exc:
                return None, None, False, f"некорректная дата: {exc}"
        error = "не распознан месяц"

    m = _RE_DOT.search(clean)
    if m:
        d_s, m_s, y_s = m.groups()
        mon = int(m_s)
        y, inf = _infer_school_year(mon, int(y_s) if y_s else None)
        try:
            iso = datetime.date(y, mon, int(d_s)).isoformat()
            return iso, iso, inf, None
        except (ValueError, TypeError) as exc:
            return None, None, False, f"некорректная дата: {exc}"

    return None, None, False, error


def parse_date_expression(raw_val: Any) -> dict[str, Any] | None:
    if raw_val is None:
        return None

    if isinstance(raw_val, (datetime.datetime, datetime.date)):
        d = raw_val.date() if isinstance(raw_val, datetime.datetime) else raw_val
        iso = d.isoformat()
        return {
            "raw_text": d.strftime("%d.%m.%Y"),
            "start_date": iso,
            "end_date": iso,
            "is_ambiguous": 0,
            "needs_review": 0,
            "reason": None,
        }

    raw_text = str(raw_val).strip()
    if not raw_text:
        return None

    reasons: list[str] = []
    is_ambiguous = False
    needs_review = False
    lower_txt = normalize_text(raw_text)

    if "*" in raw_text:
        is_ambiguous = True
        needs_review = True
        reasons.append("содержит примечание '*'")

    if AMBIGUOUS_RE.search(lower_txt):
        is_ambiguous = True
        needs_review = True
        reasons.append("приблизительная или многовариантная формулировка")

    start_iso: str | None = None
    end_iso: str | None = None

    if not is_ambiguous:
        clean = raw_text.replace("–", "-").replace("—", "-")
        start_iso, end_iso, inferred, error = _parse_text_dates(clean)
        if error or not start_iso:
            is_ambiguous = True
            needs_review = True
            reasons.append(error or "дата не распознана в надёжном формате")
        elif inferred:
            needs_review = True
            reasons.append("год соотнесён с учебным годом 2026/2027")

    return {
        "raw_text": raw_text,
        "start_date": start_iso,
        "end_date": end_iso,
        "is_ambiguous": 1 if is_ambiguous else 0,
        "needs_review": 1 if needs_review else 0,
        "reason": "; ".join(reasons) if reasons else None,
    }


def _extract_url_and_notes(ws: Any, row_idx: int, skip_cols: set[str], max_col: int) -> tuple[str | None, str | None]:
    url: str | None = None
    note_parts: list[str] = []

    for c in range(1, max_col + 1):
        col_let = openpyxl.utils.get_column_letter(c)
        if col_let in skip_cols:
            continue
        cell = ws.cell(row_idx, c)
        val = cell.value
        hyper = cell.hyperlink.target if cell.hyperlink else None
        if hyper and not url:
            url = hyper.strip()
        if val is not None and str(val).strip():
            txt = str(val).strip()
            url_match = re.search(r"(https?://\S+|www\.\S+)", txt)
            if url_match and not url:
                found_u = url_match.group(1).rstrip(".,;)")
                url = f"https://{found_u}" if found_u.startswith("www.") else found_u
                remainder = txt.replace(url_match.group(1), "").strip(" \n\r\t|;*")
                if remainder:
                    note_parts.append(remainder)
            else:
                note_parts.append(txt)

    return url, ("\n".join(note_parts) if note_parts else None)


async def import_excel_to_db(
    excel_path: Path = EXCEL_PATH,
    db_path: Path = DB_PATH,
    admin_id: int = ADMIN_ID,
    make_backup: bool = True,
) -> dict[str, Any]:
    await init_db(db_path)
    backup_path = create_db_backup(db_path) if make_backup else None

    report: dict[str, Any] = {
        "source_file": str(excel_path),
        "backup_path": backup_path,
        "found_olympiads": 0,
        "imported_stages": 0,
        "found_links": 0,
        "skipped_rows": 0,
        "errors_count": 0,
        "ambiguous_dates": [],
        "review_rows": [],
        "errors": [],
    }

    if not excel_path.exists():
        report["errors_count"] += 1
        report["errors"].append(f"Файл {excel_path} не найден.")
        return report

    try:
        wb = openpyxl.load_workbook(str(excel_path), data_only=True)
    except Exception as exc:
        logger.exception("Ошибка открытия Excel: %s", exc)
        report["errors_count"] += 1
        report["errors"].append(f"Ошибка открытия Excel: {exc}")
        return report

    now_iso = datetime.datetime.now(APP_TZ).isoformat()
    seen_keys: set[str] = set()

    conn = await get_db(db_path)
    try:
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            subject = sheet_name.strip()
            max_col = min(ws.max_column, 30)

            headers_by_col: dict[str, str] = {}
            for c in range(1, max_col + 1):
                h_val = ws.cell(1, c).value
                if h_val is not None and str(h_val).strip():
                    headers_by_col[openpyxl.utils.get_column_letter(c)] = str(h_val).strip()

            level_col: str | None = None
            grades_col: str | None = None
            stage_cols: list[tuple[str, str, str, int]] = []

            for col_let, h_text in headers_by_col.items():
                h_norm = normalize_text(h_text)
                if h_norm == "уровень":
                    level_col = col_let
                elif "класс" in h_norm:
                    grades_col = col_let
                elif h_norm in STAGE_COLUMN_MAP:
                    s_code, s_label, s_order = STAGE_COLUMN_MAP[h_norm]
                    stage_cols.append((col_let, s_code, s_label, s_order))

            skip_cols: set[str] = {"A"} | {c for c in (level_col, grades_col) if c} | {sc[0] for sc in stage_cols}

            for r in range(2, ws.max_row + 1):
                row_vals = [ws.cell(r, c).value for c in range(1, max_col + 1)]
                if all(v is None or str(v).strip() == "" for v in row_vals):
                    continue

                raw_title = ws.cell(r, 1).value
                title = str(raw_title).strip() if raw_title is not None else ""
                if not title or "подготовка к олимпиадам" in normalize_text(title):
                    report["skipped_rows"] += 1
                    continue

                try:
                    level_val = ws[f"{level_col}{r}"].value if level_col else None
                    level_str = str(level_val).strip() if level_val is not None and str(level_val).strip() else "Не указан"

                    grades_raw_cell = ws[f"{grades_col}{r}"].value if grades_col else None
                    g_info = parse_grades_value(grades_raw_cell)

                    url, notes = _extract_url_and_notes(ws, r, skip_cols, max_col)
                    if url:
                        report["found_links"] += 1

                    row_review_reasons: list[str] = []
                    if g_info["ambiguous"] and g_info["reason"]:
                        row_review_reasons.append(g_info["reason"])

                    ol_key = make_olympiad_key(subject, title)
                    seen_keys.add(ol_key)

                    stage_map: dict[str, dict[str, Any]] = {}
                    for col_let, s_code, s_label, s_order in stage_cols:
                        parsed_d = parse_date_expression(ws[f"{col_let}{r}"].value)
                        if not parsed_d:
                            continue
                        stage_map[s_code] = {
                            "stage_code": s_code,
                            "stage_name": s_label,
                            "sort_order": s_order,
                            "raw_date_text": parsed_d["raw_text"],
                            "parsed_start_date": parsed_d["start_date"],
                            "parsed_end_date": parsed_d["end_date"],
                            "is_ambiguous": parsed_d["is_ambiguous"],
                            "needs_review": parsed_d["needs_review"],
                            "ambiguity_reason": parsed_d["reason"],
                        }
                        if parsed_d["is_ambiguous"]:
                            report["ambiguous_dates"].append(
                                f"[{subject}] {title} — {s_label}: «{parsed_d['raw_text']}» ({parsed_d['reason']})"
                            )
                        if parsed_d["needs_review"]:
                            row_review_reasons.append(f"{s_label}: {parsed_d['reason']}")
                    stage_payloads = list(stage_map.values())

                    ol_needs_review = 1 if (row_review_reasons or not g_info["grades_known"]) else 0
                    if ol_needs_review:
                        reason_summary = "; ".join(row_review_reasons) if row_review_reasons else "классы не указаны в источнике"
                        report["review_rows"].append(
                            f"[{subject} | {sheet_name}!A{r}] {title} — {reason_summary}"
                        )

                    await conn.execute(
                        """
                        INSERT INTO olympiads (
                            subject, title, normalized_key, level, grades_raw, grades_json,
                            grades_known, url, notes, source_sheet, source_row,
                            needs_review, is_active, updated_at
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?)
                        ON CONFLICT(normalized_key) DO UPDATE SET
                            subject = excluded.subject,
                            title = excluded.title,
                            level = excluded.level,
                            grades_raw = excluded.grades_raw,
                            grades_json = excluded.grades_json,
                            grades_known = excluded.grades_known,
                            url = excluded.url,
                            notes = excluded.notes,
                            source_sheet = excluded.source_sheet,
                            source_row = excluded.source_row,
                            needs_review = excluded.needs_review,
                            is_active = 1,
                            updated_at = excluded.updated_at
                        """,
                        (
                            subject,
                            title,
                            ol_key,
                            level_str,
                            g_info["grades_raw"],
                            json.dumps(g_info["grades"], ensure_ascii=False),
                            g_info["grades_known"],
                            url,
                            notes,
                            sheet_name,
                            r,
                            ol_needs_review,
                            now_iso,
                        ),
                    )

                    cur_ol = await conn.execute("SELECT id FROM olympiads WHERE normalized_key = ?", (ol_key,))
                    ol_row = await cur_ol.fetchone()
                    olympiad_id = int(ol_row["id"])

                    active_codes = [sp["stage_code"] for sp in stage_payloads]
                    if active_codes:
                        placeholders = ",".join("?" for _ in active_codes)
                        await conn.execute(
                            f"DELETE FROM stages WHERE olympiad_id = ? AND stage_code NOT IN ({placeholders})",
                            [olympiad_id, *active_codes],
                        )
                    else:
                        await conn.execute("DELETE FROM stages WHERE olympiad_id = ?", (olympiad_id,))

                    for sp in stage_payloads:
                        # Ручные правки админа сохраняются, пока исходный текст даты в Excel не изменился
                        await conn.execute(
                            """
                            INSERT INTO stages (
                                olympiad_id, stage_code, stage_name, sort_order,
                                raw_date_text, parsed_start_date, parsed_end_date,
                                is_ambiguous, needs_review, ambiguity_reason, is_manual
                            )
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0)
                            ON CONFLICT(olympiad_id, stage_code) DO UPDATE SET
                                stage_name = excluded.stage_name,
                                sort_order = excluded.sort_order,
                                parsed_start_date = CASE WHEN stages.is_manual = 1 AND stages.raw_date_text = excluded.raw_date_text
                                    THEN stages.parsed_start_date ELSE excluded.parsed_start_date END,
                                parsed_end_date = CASE WHEN stages.is_manual = 1 AND stages.raw_date_text = excluded.raw_date_text
                                    THEN stages.parsed_end_date ELSE excluded.parsed_end_date END,
                                is_ambiguous = CASE WHEN stages.is_manual = 1 AND stages.raw_date_text = excluded.raw_date_text
                                    THEN stages.is_ambiguous ELSE excluded.is_ambiguous END,
                                needs_review = CASE WHEN stages.is_manual = 1 AND stages.raw_date_text = excluded.raw_date_text
                                    THEN stages.needs_review ELSE excluded.needs_review END,
                                ambiguity_reason = CASE WHEN stages.is_manual = 1 AND stages.raw_date_text = excluded.raw_date_text
                                    THEN stages.ambiguity_reason ELSE excluded.ambiguity_reason END,
                                is_manual = CASE WHEN stages.is_manual = 1 AND stages.raw_date_text = excluded.raw_date_text
                                    THEN 1 ELSE 0 END,
                                raw_date_text = excluded.raw_date_text
                            """,
                            (
                                olympiad_id,
                                sp["stage_code"],
                                sp["stage_name"],
                                sp["sort_order"],
                                sp["raw_date_text"],
                                sp["parsed_start_date"],
                                sp["parsed_end_date"],
                                sp["is_ambiguous"],
                                sp["needs_review"],
                                sp["ambiguity_reason"],
                            ),
                        )
                        report["imported_stages"] += 1

                    report["found_olympiads"] += 1
                except Exception as exc:
                    logger.exception("Ошибка импорта строки %s!%s: %s", sheet_name, r, exc)
                    report["errors_count"] += 1
                    report["errors"].append(f"[{sheet_name}!A{r}] {title}: {exc}")

        cur_all = await conn.execute("SELECT id, normalized_key FROM olympiads")
        for item in await cur_all.fetchall():
            if item["normalized_key"] not in seen_keys:
                await conn.execute(
                    "UPDATE olympiads SET is_active = 0, updated_at = ? WHERE id = ?",
                    (now_iso, item["id"]),
                )

        await conn.execute(
            """
            INSERT INTO import_logs (
                admin_id, source_file, backup_path, found_olympiads, imported_stages,
                found_links, skipped_rows, errors_count, ambiguous_count, review_count,
                details_json, created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                admin_id,
                str(excel_path),
                backup_path,
                report["found_olympiads"],
                report["imported_stages"],
                report["found_links"],
                report["skipped_rows"],
                report["errors_count"],
                len(report["ambiguous_dates"]),
                len(report["review_rows"]),
                json.dumps(
                    {
                        "ambiguous_dates": report["ambiguous_dates"][:100],
                        "review_rows": report["review_rows"][:100],
                        "errors": report["errors"][:50],
                    },
                    ensure_ascii=False,
                ),
                now_iso,
            ),
        )
        await conn.commit()
    finally:
        await conn.close()
        wb.close()

    return report


# ==============================================================================
# 4. ПОЛЬЗОВАТЕЛИ, ПОДПИСКА, ПРЕДМЕТЫ И КАТАЛОГ
# ==============================================================================
async def get_all_subjects(db_path: Path = DB_PATH) -> list[str]:
    conn = await get_db(db_path)
    try:
        cur = await conn.execute("SELECT DISTINCT subject FROM olympiads WHERE is_active = 1")
        found = [r["subject"] for r in await cur.fetchall()]
        ordered: list[str] = [s for s in DEFAULT_SUBJECTS if s in found]
        for s in sorted(found, key=lambda x: x.lower()):
            if s not in ordered:
                ordered.append(s)
        return ordered
    finally:
        await conn.close()


def _parse_user_subjects(user_row: dict[str, Any], available_subjects: list[str]) -> list[str]:
    raw_json = user_row.get("selected_subjects_json") or "[]"
    try:
        saved = json.loads(raw_json)
    except Exception:
        saved = []
    if not isinstance(saved, list):
        saved = []
    valid = [s for s in saved if isinstance(s, str)]
    if valid:
        return valid
    return list(available_subjects)


async def get_user(telegram_id: int, db_path: Path = DB_PATH) -> dict[str, Any] | None:
    """Читает пользователя, НЕ перезаписывая его данные."""
    conn = await get_db(db_path)
    try:
        cur = await conn.execute("SELECT * FROM users WHERE telegram_id = ?", (telegram_id,))
        row = await cur.fetchone()
    finally:
        await conn.close()
    if not row:
        return None
    user_dict = dict(row)
    avail = await get_all_subjects(db_path)
    user_dict["selected_subjects"] = _parse_user_subjects(user_dict, avail)
    return user_dict


async def ensure_user(
    telegram_id: int,
    full_name: str,
    username: str | None,
    db_path: Path = DB_PATH,
) -> dict[str, Any]:
    now_iso = datetime.datetime.now(APP_TZ).isoformat()
    safe_name = (full_name or "Пользователь").strip()[:200] or "Пользователь"
    safe_username = username.strip() if username else None

    conn = await get_db(db_path)
    try:
        await conn.execute(
            """
            INSERT INTO users (telegram_id, full_name, username, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(telegram_id) DO UPDATE SET
                full_name = excluded.full_name,
                username = excluded.username,
                updated_at = excluded.updated_at
            """,
            (telegram_id, safe_name, safe_username, now_iso, now_iso),
        )
        await conn.execute(
            """
            INSERT OR IGNORE INTO reminder_settings (user_id, enabled, updated_at)
            SELECT id, 1, ? FROM users WHERE telegram_id = ?
            """,
            (now_iso, telegram_id),
        )
        await conn.commit()
    finally:
        await conn.close()
    return await get_user(telegram_id, db_path) or {}


async def set_user_grade(telegram_id: int, grade: int, db_path: Path = DB_PATH) -> None:
    if not (MIN_GRADE <= grade <= MAX_GRADE):
        raise ValueError(f"Класс должен быть в диапазоне от {MIN_GRADE} до {MAX_GRADE}.")
    now_iso = datetime.datetime.now(APP_TZ).isoformat()
    conn = await get_db(db_path)
    try:
        await conn.execute(
            "UPDATE users SET grade = ?, updated_at = ? WHERE telegram_id = ?",
            (grade, now_iso, telegram_id),
        )
        await conn.commit()
    finally:
        await conn.close()


async def toggle_user_subject(telegram_id: int, subject: str, db_path: Path = DB_PATH) -> list[str]:
    user = await get_user(telegram_id, db_path)
    avail = await get_all_subjects(db_path)
    current = list((user or {}).get("selected_subjects") or avail)

    if subject in current:
        current.remove(subject)
        if not current:
            current = list(avail)
    else:
        current.append(subject)

    ordered = [s for s in avail if s in current] + [s for s in current if s not in avail]
    now_iso = datetime.datetime.now(APP_TZ).isoformat()
    conn = await get_db(db_path)
    try:
        await conn.execute(
            "UPDATE users SET selected_subjects_json = ?, updated_at = ? WHERE telegram_id = ?",
            (json.dumps(ordered, ensure_ascii=False), now_iso, telegram_id),
        )
        await conn.commit()
    finally:
        await conn.close()
    return ordered


async def set_all_user_subjects(telegram_id: int, db_path: Path = DB_PATH) -> list[str]:
    avail = await get_all_subjects(db_path)
    now_iso = datetime.datetime.now(APP_TZ).isoformat()
    conn = await get_db(db_path)
    try:
        await conn.execute(
            "UPDATE users SET selected_subjects_json = ?, updated_at = ? WHERE telegram_id = ?",
            (json.dumps(avail, ensure_ascii=False), now_iso, telegram_id),
        )
        await conn.commit()
    finally:
        await conn.close()
    return avail


async def check_channel_subscription(bot: Bot, telegram_id: int, db_path: Path = DB_PATH) -> tuple[bool, str | None]:
    now_iso = datetime.datetime.now(APP_TZ).isoformat()
    try:
        member = await bot.get_chat_member(chat_id=CHANNEL_CHAT_ID, user_id=telegram_id)
        subscribed = member.status in {
            ChatMemberStatus.CREATOR,
            ChatMemberStatus.ADMINISTRATOR,
            ChatMemberStatus.MEMBER,
        }
        conn = await get_db(db_path)
        try:
            await conn.execute(
                "UPDATE users SET is_subscribed = ?, last_sub_check = ?, updated_at = ? WHERE telegram_id = ?",
                (1 if subscribed else 0, now_iso, now_iso, telegram_id),
            )
            await conn.commit()
        finally:
            await conn.close()

        if subscribed:
            return True, None
        return False, f"Подписка на канал @{CHANNEL_USERNAME} не найдена."

    except TelegramForbiddenError:
        return False, (
            f"Не удалось проверить подписку: добавьте бота в канал @{CHANNEL_USERNAME} "
            "как администратора."
        )
    except TelegramBadRequest as exc:
        logger.error("Ошибка getChatMember: %s", exc)
        return False, "Ошибка проверки подписки через Telegram API."
    except TelegramAPIError as exc:
        logger.error("Ошибка Telegram API: %s", exc)
        return False, "Временная ошибка Telegram API. Попробуйте позже."
    except Exception as exc:
        logger.exception("Неожиданная ошибка проверки подписки: %s", exc)
        return False, "Ошибка при проверке подписки."


async def get_olympiads_for_user(
    subjects: list[str] | None,
    grade: int | None,
    ignore_grade: bool = False,
    db_path: Path = DB_PATH,
) -> list[dict[str, Any]]:
    conn = await get_db(db_path)
    try:
        sql = "SELECT * FROM olympiads WHERE is_active = 1"
        params: list[Any] = []
        if subjects:
            placeholders = ",".join("?" for _ in subjects)
            sql += f" AND subject IN ({placeholders})"
            params.extend(subjects)
        sql += " ORDER BY subject, level, title COLLATE NOCASE"
        cur = await conn.execute(sql, params)
        rows = [dict(r) for r in await cur.fetchall()]

        if ignore_grade or grade is None:
            return rows

        result: list[dict[str, Any]] = []
        for r in rows:
            if r.get("grades_known"):
                g_list = json.loads(r.get("grades_json") or "[]")
                if grade in g_list:
                    result.append(r)
            else:
                result.append(r)
        return result
    finally:
        await conn.close()


async def search_olympiads(query_text: str, db_path: Path = DB_PATH) -> list[dict[str, Any]]:
    q = normalize_text(query_text)
    if not q:
        return []
    conn = await get_db(db_path)
    try:
        cur = await conn.execute("SELECT * FROM olympiads WHERE is_active = 1 ORDER BY subject, title")
        rows = [dict(r) for r in await cur.fetchall()]
        results: list[dict[str, Any]] = []
        for r in rows:
            haystack = " ".join(
                filter(
                    None,
                    [r.get("title"), r.get("subject"), r.get("level"), r.get("notes"), r.get("grades_raw")],
                )
            )
            matched = q in normalize_text(haystack)
            if not matched and q.isdigit() and MIN_GRADE <= int(q) <= MAX_GRADE:
                matched = int(q) in json.loads(r.get("grades_json") or "[]")
            if matched:
                results.append(r)
        return results
    finally:
        await conn.close()


async def get_olympiad_card_data(olympiad_id: int, user_db_id: int, db_path: Path = DB_PATH) -> dict[str, Any] | None:
    conn = await get_db(db_path)
    try:
        cur = await conn.execute("SELECT * FROM olympiads WHERE id = ?", (olympiad_id,))
        ol_row = await cur.fetchone()
        if not ol_row:
            return None
        ol = dict(ol_row)

        cur_reg = await conn.execute(
            "SELECT 1 FROM registrations WHERE user_id = ? AND olympiad_id = ?",
            (user_db_id, olympiad_id),
        )
        ol["is_registered"] = (await cur_reg.fetchone()) is not None

        cur_st = await conn.execute(
            """
            SELECT s.*, COALESCE(p.is_completed, 0) AS is_completed
            FROM stages s
            LEFT JOIN progress p ON p.stage_id = s.id AND p.user_id = ?
            WHERE s.olympiad_id = ?
            ORDER BY s.sort_order, s.id
            """,
            (user_db_id, olympiad_id),
        )
        ol["stages"] = [dict(s) for s in await cur_st.fetchall()]
        return ol
    finally:
        await conn.close()


def format_olympiad_card_text(ol: dict[str, Any]) -> str:
    if ol.get("grades_known"):
        grades_display = ol.get("grades_raw") or ", ".join(map(str, json.loads(ol.get("grades_json") or "[]")))
    else:
        grades_display = "Не указаны в источнике"

    stage_lines: list[str] = []
    for st in ol.get("stages", []):
        mark = "✅" if st.get("is_completed") else "⬜"
        d_str = fmt_date_range_ru(st.get("parsed_start_date"), st.get("parsed_end_date"), st.get("raw_date_text", ""))
        stage_lines.append(f"{mark} <b>{esc(st['stage_name'])}</b>: {d_str}")

    lines = [
        f"🏅 <b>{esc(ol['title'])}</b>",
        f"📌 <b>{esc(ol['subject'])}</b> | Уровень: <b>{esc(ol.get('level') or '-')}</b> | Классы: <b>{esc(grades_display)}</b>",
    ]
    if ol.get("is_registered"):
        lines.append("📝 <b>Статус:</b> Вы зарегистрированы")
    if not ol.get("is_active", 1):
        lines.append("⚠️ <i>Неактивна в последнем обновлении таблицы.</i>")

    lines.append("\n<b>Этапы и даты:</b>")
    lines.extend(stage_lines if stage_lines else ["• Даты не указаны"])

    if ol.get("url"):
        lines.append(f"\n🔗 Сайт: {esc(ol['url'])}")
    if ol.get("notes"):
        notes = str(ol["notes"])
        if len(notes) > 1200:
            notes = notes[:1200] + "…"
        lines.append(f"📝 {esc(notes)}")
    return join_limited(lines, 4000)


# ==============================================================================
# 5. РАССЫЛКА ПО СТАТУСУ РЕГИСТРАЦИИ ПРИ ВЫБОРЕ ПРЕДМЕТОВ (ИДЁТ / ЗАКОНЧИЛАСЬ)
# ==============================================================================
async def build_registration_split_digest(
    subjects: list[str],
    grade: int | None,
    today_date: datetime.date | None = None,
    db_path: Path = DB_PATH,
) -> dict[str, list[dict[str, Any]]]:
    """
    Делит олимпиады по выбранным предметам и классу на две группы:
    - active_reg: хотя бы один этап регистрации ещё не завершён к ref_day;
    - closed_reg: все этапы регистрации с известными датами уже завершены.
    """
    ref_day = today_date or datetime.datetime.now(APP_TZ).date()
    ols = await get_olympiads_for_user(subjects=subjects, grade=grade, ignore_grade=False, db_path=db_path)

    active_reg: list[dict[str, Any]] = []
    closed_reg: list[dict[str, Any]] = []

    conn = await get_db(db_path)
    try:
        for ol in ols:
            cur_s = await conn.execute(
                """
                SELECT * FROM stages
                WHERE olympiad_id = ? AND stage_code IN ('registration', 'final_reg')
                ORDER BY sort_order, id
                """,
                (ol["id"],),
            )
            candidates: list[tuple[dict[str, Any], str | None, str, datetime.date]] = []
            for s in [dict(r) for r in await cur_s.fetchall()]:
                start_str = s.get("parsed_start_date")
                end_str = s.get("parsed_end_date") or start_str
                if not end_str:
                    continue
                try:
                    end_d = datetime.date.fromisoformat(end_str)
                except ValueError:
                    continue
                candidates.append((s, start_str, end_str, end_d))
            if not candidates:
                continue

            open_c = [c for c in candidates if c[3] >= ref_day]
            chosen = open_c[0] if open_c else candidates[-1]
            st, start_str, end_str, _ = chosen

            entry = {
                "id": ol["id"],
                "title": ol["title"],
                "subject": ol["subject"],
                "level": ol.get("level"),
                "url": ol.get("url"),
                "raw_date_text": st["raw_date_text"],
                "start_date": start_str,
                "end_date": end_str,
            }
            (active_reg if open_c else closed_reg).append(entry)
    finally:
        await conn.close()

    return {"active_reg": active_reg, "closed_reg": closed_reg}


async def send_split_registration_broadcasts(
    message_or_bot: Message | Bot,
    chat_id: int,
    subjects: list[str],
    grade: int | None,
    today_date: datetime.date | None = None,
    db_path: Path = DB_PATH,
) -> None:
    """Отправляет два сообщения: регистрация идёт / регистрация закончилась."""
    digest = await build_registration_split_digest(
        subjects=subjects,
        grade=grade,
        today_date=today_date,
        db_path=db_path,
    )
    active_list = digest["active_reg"]
    closed_list = digest["closed_reg"]

    subj_title = esc(", ".join(subjects)) if subjects else "все предметы"
    grade_title = f"{grade} кл." if grade else "7–11 кл."

    lines_open = [f"🟢 <b>Регистрация идёт</b> ({subj_title}, {grade_title}):\n"]
    if not active_list:
        lines_open.append("• Сейчас нет олимпиад с активной регистрацией.")
    else:
        for item in active_list[:20]:
            period = fmt_date_range_ru(item.get("start_date"), item.get("end_date"), item["raw_date_text"])
            lines_open.append(f"• [{esc(item['subject'])}] <b>{esc(item['title'])}</b> — {period}")
        if len(active_list) > 20:
            lines_open.append(f"<i>…и ещё {len(active_list) - 20} (открой Каталог)</i>")

    lines_closed = [f"🔴 <b>Регистрация закончилась</b> ({subj_title}, {grade_title}):\n"]
    if not closed_list:
        lines_closed.append("• Нет олимпиад с завершённой регистрацией.")
    else:
        for item in closed_list[:20]:
            period = fmt_date_range_ru(item.get("start_date"), item.get("end_date"), item["raw_date_text"])
            lines_closed.append(f"• [{esc(item['subject'])}] <b>{esc(item['title'])}</b> — {period}")
        if len(closed_list) > 20:
            lines_closed.append(f"<i>…и ещё {len(closed_list) - 20}</i>")

    msg_open = join_limited(lines_open)
    msg_closed = join_limited(lines_closed)

    if isinstance(message_or_bot, Message):
        await message_or_bot.answer(msg_open)
        await message_or_bot.answer(msg_closed)
    else:
        await message_or_bot.send_message(chat_id=chat_id, text=msg_open)
        await message_or_bot.send_message(chat_id=chat_id, text=msg_closed)


# ==============================================================================
# 6. МОИ ОЛИМПИАДЫ (РЕГИСТРАЦИИ), ЭТАПЫ, ИСТОРИЯ И СТАТИСТИКА
# ==============================================================================
async def toggle_olympiad_registration(user_db_id: int, olympiad_id: int, db_path: Path = DB_PATH) -> bool:
    now_iso = datetime.datetime.now(APP_TZ).isoformat()
    conn = await get_db(db_path)
    try:
        cur_info = await conn.execute("SELECT title FROM olympiads WHERE id = ?", (olympiad_id,))
        ol_row = await cur_info.fetchone()
        ol_title = ol_row["title"] if ol_row else f"Олимпиада #{olympiad_id}"

        cur_reg = await conn.execute(
            "SELECT id FROM registrations WHERE user_id = ? AND olympiad_id = ?",
            (user_db_id, olympiad_id),
        )
        reg_row = await cur_reg.fetchone()
        new_state = reg_row is None

        if new_state:
            await conn.execute(
                "INSERT OR IGNORE INTO registrations (user_id, olympiad_id, created_at) VALUES (?, ?, ?)",
                (user_db_id, olympiad_id, now_iso),
            )
        else:
            await conn.execute(
                "DELETE FROM registrations WHERE user_id = ? AND olympiad_id = ?",
                (user_db_id, olympiad_id),
            )

        cur_st = await conn.execute(
            "SELECT id, stage_name FROM stages WHERE olympiad_id = ? AND stage_code = 'registration' LIMIT 1",
            (olympiad_id,),
        )
        st = await cur_st.fetchone()
        if st:
            stage_id = int(st["id"])
            await conn.execute(
                """
                INSERT INTO progress (user_id, stage_id, is_completed, completed_at, updated_at)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(user_id, stage_id) DO UPDATE SET
                    is_completed = excluded.is_completed,
                    completed_at = excluded.completed_at,
                    updated_at = excluded.updated_at
                """,
                (user_db_id, stage_id, 1 if new_state else 0, now_iso if new_state else None, now_iso),
            )
            await conn.execute(
                """
                INSERT INTO progress_audit (user_id, stage_id, olympiad_title, stage_name, new_status, changed_at)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (user_db_id, stage_id, ol_title, st["stage_name"], 1 if new_state else 0, now_iso),
            )

        await conn.commit()
        return new_state
    finally:
        await conn.close()


async def get_user_registered_olympiads(user_db_id: int, db_path: Path = DB_PATH) -> list[dict[str, Any]]:
    conn = await get_db(db_path)
    try:
        cur = await conn.execute(
            """
            SELECT o.* FROM registrations r
            JOIN olympiads o ON o.id = r.olympiad_id
            WHERE r.user_id = ?
            ORDER BY o.subject, o.title
            """,
            (user_db_id,),
        )
        items = [dict(r) for r in await cur.fetchall()]
        for item in items:
            cur_s = await conn.execute(
                """
                SELECT s.*, COALESCE(p.is_completed, 0) AS is_completed
                FROM stages s
                LEFT JOIN progress p ON p.stage_id = s.id AND p.user_id = ?
                WHERE s.olympiad_id = ?
                ORDER BY s.sort_order, s.id
                """,
                (user_db_id, item["id"]),
            )
            item["stages"] = [dict(s) for s in await cur_s.fetchall()]
        return items
    finally:
        await conn.close()


async def toggle_stage_progress(user_db_id: int, stage_id: int, db_path: Path = DB_PATH) -> bool:
    now_iso = datetime.datetime.now(APP_TZ).isoformat()
    conn = await get_db(db_path)
    try:
        cur_info = await conn.execute(
            """
            SELECT s.id, s.olympiad_id, s.stage_code, s.stage_name, o.title AS olympiad_title
            FROM stages s JOIN olympiads o ON o.id = s.olympiad_id
            WHERE s.id = ?
            """,
            (stage_id,),
        )
        info = await cur_info.fetchone()
        if not info:
            raise ValueError(f"Этап {stage_id} не найден.")

        cur_p = await conn.execute(
            "SELECT is_completed FROM progress WHERE user_id = ? AND stage_id = ?",
            (user_db_id, stage_id),
        )
        p_row = await cur_p.fetchone()
        new_state = not bool(p_row["is_completed"]) if p_row else True
        completed_at = now_iso if new_state else None

        await conn.execute(
            """
            INSERT INTO progress (user_id, stage_id, is_completed, completed_at, updated_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(user_id, stage_id) DO UPDATE SET
                is_completed = excluded.is_completed,
                completed_at = excluded.completed_at,
                updated_at = excluded.updated_at
            """,
            (user_db_id, stage_id, 1 if new_state else 0, completed_at, now_iso),
        )
        await conn.execute(
            """
            INSERT INTO progress_audit (user_id, stage_id, olympiad_title, stage_name, new_status, changed_at)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (user_db_id, stage_id, info["olympiad_title"], info["stage_name"], 1 if new_state else 0, now_iso),
        )

        if new_state:
            await conn.execute(
                "INSERT OR IGNORE INTO registrations (user_id, olympiad_id, created_at) VALUES (?, ?, ?)",
                (user_db_id, int(info["olympiad_id"]), now_iso),
            )
        elif info["stage_code"] == "registration":
            await conn.execute(
                "DELETE FROM registrations WHERE user_id = ? AND olympiad_id = ?",
                (user_db_id, int(info["olympiad_id"])),
            )

        await conn.commit()
        return new_state
    finally:
        await conn.close()


async def add_history_entry(
    user_db_id: int,
    olympiad_title: str,
    subject: str,
    participation_year: int,
    stage_reached: str,
    result_text: str | None = None,
    note_text: str | None = None,
    is_final: bool = False,
    olympiad_id: int | None = None,
    db_path: Path = DB_PATH,
) -> int:
    clean_title = olympiad_title.strip()
    clean_subject = subject.strip()
    clean_stage = stage_reached.strip()
    if not clean_title or not clean_subject or not clean_stage:
        raise ValueError("Заполните название, предмет и этап.")
    if not (1990 <= participation_year <= 2100):
        raise ValueError("Некорректный год.")

    now_iso = datetime.datetime.now(APP_TZ).isoformat()
    conn = await get_db(db_path)
    try:
        cur = await conn.execute(
            """
            INSERT INTO history (
                user_id, olympiad_id, olympiad_title, subject, participation_year,
                stage_reached, result_text, note_text, is_final, created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                user_db_id,
                olympiad_id,
                clean_title,
                clean_subject,
                participation_year,
                clean_stage,
                result_text.strip() if result_text and result_text.strip() else None,
                note_text.strip() if note_text and note_text.strip() else None,
                1 if is_final else 0,
                now_iso,
            ),
        )
        await conn.commit()
        return int(cur.lastrowid)
    finally:
        await conn.close()


async def get_user_history(user_db_id: int, db_path: Path = DB_PATH) -> list[dict[str, Any]]:
    conn = await get_db(db_path)
    try:
        cur = await conn.execute(
            "SELECT * FROM history WHERE user_id = ? ORDER BY participation_year DESC, id DESC",
            (user_db_id,),
        )
        return [dict(r) for r in await cur.fetchall()]
    finally:
        await conn.close()


async def _count(conn: aiosqlite.Connection, sql: str, params: tuple = ()) -> int:
    cur = await conn.execute(sql, params)
    row = await cur.fetchone()
    return int(row["c"]) if row else 0


async def get_user_statistics(user_db_id: int, db_path: Path = DB_PATH) -> dict[str, Any]:
    conn = await get_db(db_path)
    try:
        catalog_total = await _count(conn, "SELECT COUNT(*) AS c FROM olympiads WHERE is_active = 1")
        registered_count = await _count(conn, "SELECT COUNT(*) AS c FROM registrations WHERE user_id = ?", (user_db_id,))
        completed_stages = await _count(
            conn, "SELECT COUNT(*) AS c FROM progress WHERE user_id = ? AND is_completed = 1", (user_db_id,)
        )
        cur_h = await conn.execute("SELECT * FROM history WHERE user_id = ?", (user_db_id,))
        hist_rows = [dict(r) for r in await cur_h.fetchall()]

        distinct_titles = {normalize_text(h["olympiad_title"]) for h in hist_rows if h.get("olympiad_title")}
        by_subject: dict[str, int] = {}
        by_year: dict[int, int] = {}
        for h in hist_rows:
            by_subject[h["subject"]] = by_subject.get(h["subject"], 0) + 1
            by_year[int(h["participation_year"])] = by_year.get(int(h["participation_year"]), 0) + 1

        return {
            "catalog_total": catalog_total,
            "registered_count": registered_count,
            "completed_stages": completed_stages,
            "history_count": len(hist_rows),
            "participated_olympiads_count": len(distinct_titles),
            "finals_count": sum(1 for h in hist_rows if h.get("is_final")),
            "by_subject": dict(sorted(by_subject.items(), key=lambda x: (-x[1], x[0]))),
            "by_year": dict(sorted(by_year.items(), key=lambda x: x[0], reverse=True)),
        }
    finally:
        await conn.close()


# ==============================================================================
# 7. СРОЧНЫЕ (< 7 ДН. ДО РЕГИСТРАЦИИ) И ПЛАНОВЫЕ НАПОМИНАНИЯ (APSCHEDULER)
# ==============================================================================
async def set_reminders_enabled(user_db_id: int, enabled: bool, db_path: Path = DB_PATH) -> None:
    now_iso = datetime.datetime.now(APP_TZ).isoformat()
    conn = await get_db(db_path)
    try:
        await conn.execute(
            """
            INSERT INTO reminder_settings (user_id, enabled, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(user_id) DO UPDATE SET enabled = excluded.enabled, updated_at = excluded.updated_at
            """,
            (user_db_id, 1 if enabled else 0, now_iso),
        )
        await conn.commit()
    finally:
        await conn.close()


async def is_reminders_enabled(user_db_id: int, db_path: Path = DB_PATH) -> bool:
    conn = await get_db(db_path)
    try:
        cur = await conn.execute("SELECT enabled FROM reminder_settings WHERE user_id = ?", (user_db_id,))
        row = await cur.fetchone()
        return bool(row["enabled"]) if row is not None else True
    finally:
        await conn.close()


async def collect_due_reminders(today_date: datetime.date | None = None, db_path: Path = DB_PATH) -> list[dict[str, Any]]:
    """
    Формирует уведомления по выбранным предметам пользователя (и его классу 7-11):
    1) СРОЧНАЯ рассылка, если до дедлайна регистрации осталось меньше 7 дней (0 <= delta_days < 7);
    2) Плановые напоминания за 7, 3, 1 день и в день события по турам и этапам.
    """
    ref_day = today_date or datetime.datetime.now(APP_TZ).date()
    due_items: list[dict[str, Any]] = []

    avail_subjects = await get_all_subjects(db_path)
    conn = await get_db(db_path)
    try:
        cur_users = await conn.execute(
            """
            SELECT u.id AS user_db_id, u.telegram_id, u.grade, u.selected_subjects_json
            FROM users u
            LEFT JOIN reminder_settings rs ON rs.user_id = u.id
            WHERE u.is_subscribed = 1
              AND COALESCE(rs.enabled, 1) = 1
            """
        )
        users = [dict(u) for u in await cur_users.fetchall()]
        if not users:
            return []

        cur_stages = await conn.execute(
            """
            SELECT
                s.id AS stage_id, s.stage_code, s.stage_name, s.raw_date_text,
                s.parsed_start_date, s.parsed_end_date, s.is_ambiguous,
                o.id AS olympiad_id, o.title AS olympiad_title, o.subject, o.url,
                o.grades_json, o.grades_known
            FROM stages s
            JOIN olympiads o ON o.id = s.olympiad_id
            WHERE o.is_active = 1
              AND s.is_ambiguous = 0
              AND s.parsed_start_date IS NOT NULL
            """
        )
        stages = [dict(s) for s in await cur_stages.fetchall()]

        upcoming_stages: list[dict[str, Any]] = []
        for st in stages:
            is_reg = st["stage_code"] in REG_STAGE_CODES
            target_str = st["parsed_end_date"] if (is_reg and st["parsed_end_date"]) else st["parsed_start_date"]
            if not target_str:
                continue
            try:
                target_day = datetime.date.fromisoformat(target_str)
            except ValueError:
                continue

            delta_days = (target_day - ref_day).days
            is_urgent_reg = is_reg and (0 <= delta_days < URGENT_REG_DAYS)

            if not is_urgent_reg and delta_days not in REMINDER_OFFSETS:
                continue

            upcoming_stages.append(
                {
                    **st,
                    "target_date": target_day.isoformat(),
                    "offset_days": delta_days,
                    "is_registration_deadline": is_reg,
                    "is_urgent_registration": is_urgent_reg,
                    "grades_list": json.loads(st.get("grades_json") or "[]") if st.get("grades_known") else None,
                }
            )

        if not upcoming_stages:
            return []

        for u in users:
            user_subjects = set(_parse_user_subjects(u, avail_subjects))
            user_grade = u.get("grade")

            cur_hist_subj = await conn.execute(
                """
                SELECT DISTINCT o.subject
                FROM registrations r
                JOIN olympiads o ON o.id = r.olympiad_id
                WHERE r.user_id = ?
                """,
                (u["user_db_id"],),
            )
            registered_subjects = {r["subject"] for r in await cur_hist_subj.fetchall()}
            monitored_subjects = user_subjects | registered_subjects

            for st in upcoming_stages:
                if st["subject"] not in monitored_subjects:
                    continue
                if user_grade is not None and st["grades_list"] is not None:
                    if user_grade not in st["grades_list"]:
                        continue

                cur_prog = await conn.execute(
                    "SELECT is_completed FROM progress WHERE user_id = ? AND stage_id = ?",
                    (u["user_db_id"], st["stage_id"]),
                )
                p_row = await cur_prog.fetchone()
                if p_row and p_row["is_completed"]:
                    continue

                cur_sent = await conn.execute(
                    "SELECT 1 FROM reminders WHERE user_id = ? AND stage_id = ? AND target_date = ? AND offset_days = ?",
                    (u["user_db_id"], st["stage_id"], st["target_date"], st["offset_days"]),
                )
                if await cur_sent.fetchone():
                    continue

                cur_reg = await conn.execute(
                    "SELECT 1 FROM registrations WHERE user_id = ? AND olympiad_id = ?",
                    (u["user_db_id"], st["olympiad_id"]),
                )
                is_reg_user = (await cur_reg.fetchone()) is not None

                due_items.append(
                    {
                        "user_db_id": u["user_db_id"],
                        "telegram_id": u["telegram_id"],
                        "stage_id": st["stage_id"],
                        "stage_code": st["stage_code"],
                        "stage_name": st["stage_name"],
                        "raw_date_text": st["raw_date_text"],
                        "olympiad_id": st["olympiad_id"],
                        "olympiad_title": st["olympiad_title"],
                        "subject": st["subject"],
                        "url": st["url"],
                        "target_date": st["target_date"],
                        "offset_days": st["offset_days"],
                        "is_registration_deadline": st["is_registration_deadline"],
                        "is_urgent_registration": st["is_urgent_registration"],
                        "is_user_registered": is_reg_user,
                    }
                )
    finally:
        await conn.close()
    return due_items


async def mark_reminder_sent(user_db_id: int, stage_id: int, target_date: str, offset_days: int, db_path: Path = DB_PATH) -> None:
    now_iso = datetime.datetime.now(APP_TZ).isoformat()
    conn = await get_db(db_path)
    try:
        await conn.execute(
            "INSERT OR IGNORE INTO reminders (user_id, stage_id, target_date, offset_days, sent_at) VALUES (?, ?, ?, ?, ?)",
            (user_db_id, stage_id, target_date, offset_days, now_iso),
        )
        await conn.commit()
    finally:
        await conn.close()


def format_reminder_message(item: dict[str, Any]) -> str:
    offset = item["offset_days"]
    when_label = "СЕГОДНЯ" if offset == 0 else ("завтра (остался 1 день)" if offset == 1 else f"осталось {offset} дн.")
    reg_badge = " · вы зарегистрированы ✅" if item.get("is_user_registered") else ""

    if item.get("is_urgent_registration"):
        header = f"🚨 <b>СРОЧНАЯ РАССЫЛКА: до конца регистрации {when_label}!</b>"
        kind_label = "Дедлайн регистрации"
    else:
        header = f"⏰ <b>Напоминание ({when_label})</b>"
        kind_label = "Окончание регистрации" if item.get("is_registration_deadline") else "Тур / этап"

    lines = [
        header,
        f"🏅 <b>{esc(item['olympiad_title'])}</b> ({esc(item['subject'])}{reg_badge})",
        f"📌 {kind_label}: <b>{esc(item['stage_name'])}</b>",
        f"📅 {fmt_date_ru(item['target_date'])} ({esc(item['raw_date_text'])})",
    ]
    if item.get("url"):
        lines.append(f"🔗 {esc(item['url'])}")
    return "\n".join(lines)


async def run_reminders_tick(bot: Bot | None, db_path: Path = DB_PATH) -> int:
    try:
        due = await collect_due_reminders(db_path=db_path)
    except Exception as exc:
        logger.exception("Ошибка сбора напоминаний: %s", exc)
        return 0

    sent_count = 0
    for item in due:
        text = format_reminder_message(item)
        kb = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="📋 Открыть олимпиаду", callback_data=f"ol:view:{item['olympiad_id']}:menu:main")]
            ]
        )
        if bot is not None:
            try:
                await bot.send_message(chat_id=item["telegram_id"], text=text, reply_markup=kb)
            except TelegramForbiddenError:
                # Пользователь заблокировал бота — помечаем как отправленное, чтобы не повторять каждые 15 минут
                await mark_reminder_sent(item["user_db_id"], item["stage_id"], item["target_date"], item["offset_days"], db_path=db_path)
                continue
            except TelegramAPIError as exc:
                logger.error("Ошибка отправки напоминания %s: %s", item["telegram_id"], exc)
                continue
        await mark_reminder_sent(item["user_db_id"], item["stage_id"], item["target_date"], item["offset_days"], db_path=db_path)
        sent_count += 1
        if bot is not None:
            await asyncio.sleep(0.04)
    return sent_count


# ==============================================================================
# 7.1. РАССЫЛКА ПРИ ЗАПУСКЕ БОТА
# ==============================================================================
STARTUP_DIGEST_ENABLED = os.getenv("STARTUP_DIGEST", "1").strip() not in ("0", "false", "False", "no")


async def build_upcoming_digest(
    subjects: list[str],
    grade: int | None,
    ref_day: datetime.date,
    days: int = URGENT_REG_DAYS,
    db_path: Path = DB_PATH,
) -> list[dict[str, Any]]:
    """
    Этапы, до которых осталось меньше `days` дней (0..days-1):
    - для регистрации — её дедлайн (конец регистрации);
    - для остальных этапов (туры, финал) — дата начала.
    """
    ols = await get_olympiads_for_user(subjects=subjects, grade=grade, ignore_grade=False, db_path=db_path)
    ids = {o["id"]: o for o in ols}
    if not ids:
        return []

    conn = await get_db(db_path)
    try:
        placeholders = ",".join("?" for _ in ids)
        cur = await conn.execute(
            f"""
            SELECT * FROM stages
            WHERE olympiad_id IN ({placeholders})
              AND is_ambiguous = 0 AND parsed_start_date IS NOT NULL
            ORDER BY olympiad_id, sort_order
            """,
            list(ids.keys()),
        )
        rows = [dict(r) for r in await cur.fetchall()]
    finally:
        await conn.close()

    result: list[dict[str, Any]] = []
    for st in rows:
        is_reg = st["stage_code"] in REG_STAGE_CODES
        target_str = st["parsed_end_date"] if (is_reg and st["parsed_end_date"]) else st["parsed_start_date"]
        try:
            target_day = datetime.date.fromisoformat(target_str)
        except (TypeError, ValueError):
            continue
        left = (target_day - ref_day).days
        if not (0 <= left < days):
            continue
        ol = ids[st["olympiad_id"]]
        result.append(
            {
                "title": ol["title"],
                "subject": ol["subject"],
                "url": ol.get("url"),
                "stage_name": st["stage_name"],
                "is_reg": is_reg,
                "target_date": target_day.isoformat(),
                "days_left": left,
                "raw_date_text": st["raw_date_text"],
            }
        )
    result.sort(key=lambda x: (x["days_left"], x["subject"], x["title"]))
    return result


async def send_startup_digest(bot: Bot, db_path: Path = DB_PATH) -> int:
    """
    При запуске каждому подписанному пользователю (по его предметам и классу) отправляет:
    1) олимпиады, у которых регистрация уже закончилась;
    2) то, что начинается / закрывается меньше чем через 7 дней.
    """
    ref_day = datetime.datetime.now(APP_TZ).date()
    avail = await get_all_subjects(db_path)

    conn = await get_db(db_path)
    try:
        cur = await conn.execute(
            "SELECT telegram_id, grade, selected_subjects_json FROM users WHERE is_subscribed = 1"
        )
        users = [dict(r) for r in await cur.fetchall()]
    finally:
        await conn.close()

    sent_users = 0
    for u in users:
        subjects = _parse_user_subjects(u, avail)
        grade = u.get("grade")
        try:
            digest = await build_registration_split_digest(subjects, grade, ref_day, db_path)
            upcoming = await build_upcoming_digest(subjects, grade, ref_day, db_path=db_path)
        except Exception as exc:
            logger.exception("Ошибка формирования стартовой рассылки для %s: %s", u["telegram_id"], exc)
            continue

        subj_title = esc(", ".join(subjects)) if subjects else "все предметы"
        grade_title = f"{grade} кл." if grade else "7–11 кл."

        lines_soon = [f"🚨 <b>Скоро — меньше {URGENT_REG_DAYS} дн.</b> ({subj_title}, {grade_title}):\n"]
        if not upcoming:
            lines_soon.append("• Нет этапов, до которых осталось меньше 7 дней.")
        for it in upcoming[:25]:
            when = "СЕГОДНЯ" if it["days_left"] == 0 else ("завтра" if it["days_left"] == 1 else f"через {it['days_left']} дн.")
            kind = "конец регистрации" if it["is_reg"] else "начало"
            lines_soon.append(
                f"• [{esc(it['subject'])}] <b>{esc(it['title'])}</b> — {esc(it['stage_name'])} "
                f"({kind}): {fmt_date_ru(it['target_date'])}, <b>{when}</b>"
            )
        if len(upcoming) > 25:
            lines_soon.append(f"<i>…и ещё {len(upcoming) - 25}</i>")

        closed = digest["closed_reg"]
        lines_closed = [f"🔴 <b>Регистрация закончилась</b> ({subj_title}, {grade_title}):\n"]
        if not closed:
            lines_closed.append("• Нет олимпиад с завершённой регистрацией.")
        for it in closed[:25]:
            period = fmt_date_range_ru(it.get("start_date"), it.get("end_date"), it["raw_date_text"])
            lines_closed.append(f"• [{esc(it['subject'])}] <b>{esc(it['title'])}</b> — {period}")
        if len(closed) > 25:
            lines_closed.append(f"<i>…и ещё {len(closed) - 25}</i>")

        try:
            await bot.send_message(chat_id=u["telegram_id"], text=join_limited(lines_soon))
            await bot.send_message(chat_id=u["telegram_id"], text=join_limited(lines_closed))
            sent_users += 1
        except TelegramForbiddenError:
            continue
        except TelegramAPIError as exc:
            logger.error("Стартовая рассылка не доставлена %s: %s", u["telegram_id"], exc)
        await asyncio.sleep(0.05)

    logger.info("Стартовая рассылка отправлена пользователям: %s из %s", sent_users, len(users))
    return sent_users


# ==============================================================================
# 8. АДМИН-ФУНКЦИИ
# ==============================================================================
async def get_admin_stats(db_path: Path = DB_PATH) -> dict[str, Any]:
    conn = await get_db(db_path)
    try:
        return {
            "users_total": await _count(conn, "SELECT COUNT(*) AS c FROM users"),
            "users_subscribed": await _count(conn, "SELECT COUNT(*) AS c FROM users WHERE is_subscribed = 1"),
            "olympiads_active": await _count(conn, "SELECT COUNT(*) AS c FROM olympiads WHERE is_active = 1"),
            "stages_total": await _count(conn, "SELECT COUNT(*) AS c FROM stages"),
            "ambiguous_stages": await _count(conn, "SELECT COUNT(*) AS c FROM stages WHERE is_ambiguous = 1"),
            "review_olympiads": await _count(conn, "SELECT COUNT(*) AS c FROM olympiads WHERE needs_review = 1"),
        }
    finally:
        await conn.close()


async def get_ambiguous_stages_list(limit: int = 10, db_path: Path = DB_PATH) -> list[dict[str, Any]]:
    conn = await get_db(db_path)
    try:
        cur = await conn.execute(
            """
            SELECT s.*, o.title AS olympiad_title, o.subject
            FROM stages s JOIN olympiads o ON o.id = s.olympiad_id
            WHERE s.is_ambiguous = 1
            ORDER BY o.subject, o.title
            LIMIT ?
            """,
            (limit,),
        )
        return [dict(r) for r in await cur.fetchall()]
    finally:
        await conn.close()


async def fix_stage_date_manual(stage_id: int, iso_date_str: str, db_path: Path = DB_PATH) -> tuple[bool, str]:
    try:
        parsed_d = datetime.date.fromisoformat(iso_date_str.strip())
    except ValueError:
        return False, "Формат даты: ГГГГ-ММ-ДД (например, 2027-02-14)."

    iso = parsed_d.isoformat()
    conn = await get_db(db_path)
    try:
        cur = await conn.execute("SELECT id, olympiad_id FROM stages WHERE id = ?", (stage_id,))
        row = await cur.fetchone()
        if not row:
            return False, "Этап не найден."
        await conn.execute(
            """
            UPDATE stages
            SET parsed_start_date = ?, parsed_end_date = ?, is_ambiguous = 0, needs_review = 0,
                ambiguity_reason = NULL, is_manual = 1
            WHERE id = ?
            """,
            (iso, iso, stage_id),
        )
        rem_cnt = await _count(
            conn,
            "SELECT COUNT(*) AS c FROM stages WHERE olympiad_id = ? AND (is_ambiguous = 1 OR needs_review = 1)",
            (row["olympiad_id"],),
        )
        if rem_cnt == 0:
            await conn.execute("UPDATE olympiads SET needs_review = 0 WHERE id = ?", (row["olympiad_id"],))
        await conn.commit()
    finally:
        await conn.close()
    return True, f"Дата этапа #{stage_id} обновлена на {parsed_d.strftime('%d.%m.%Y')}."


# ==============================================================================
# 9. ИНТЕРФЕЙС И FSM
# ==============================================================================
class UserStates(StatesGroup):
    waiting_search_query = State()
    waiting_history_title = State()
    waiting_history_subject = State()
    waiting_history_year = State()
    waiting_history_stage = State()
    waiting_history_result = State()
    waiting_history_note = State()
    waiting_admin_fix_date = State()
    waiting_admin_broadcast = State()


def main_menu_kb(telegram_id: int) -> InlineKeyboardMarkup:
    rows = [
        [
            InlineKeyboardButton(text="📚 Каталог", callback_data="menu:catalog:0"),
            InlineKeyboardButton(text="🔎 Поиск", callback_data="menu:search"),
        ],
        [
            InlineKeyboardButton(text="📝 Мои олимпиады", callback_data="menu:my"),
            InlineKeyboardButton(text="🗂 История", callback_data="menu:history"),
        ],
        [
            InlineKeyboardButton(text="👤 Профиль и рассылка", callback_data="menu:profile"),
            InlineKeyboardButton(text="ℹ️ Помощь", callback_data="menu:help"),
        ],
    ]
    if is_admin(telegram_id):
        rows.append([InlineKeyboardButton(text="🛠 Админ", callback_data="menu:admin")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def subscription_required_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="📢 Подписаться", url=CHANNEL_URL)],
            [InlineKeyboardButton(text="✅ Проверить подписку", callback_data="sub:check")],
        ]
    )


def back_kb(cb: str = "menu:main", text: str = "⬅️ Назад") -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=text, callback_data=cb)]])


def grade_kb(back_cb: str = "menu:main") -> InlineKeyboardMarkup:
    row = [InlineKeyboardButton(text=f"{g} кл.", callback_data=f"grade:set:{g}") for g in GRADES_RANGE]
    return InlineKeyboardMarkup(
        inline_keyboard=[
            row,
            [InlineKeyboardButton(text="⬅️ Назад", callback_data=back_cb)],
        ]
    )


def subjects_multi_kb(available: list[str], selected: list[str], back_cb: str = "menu:catalog:0") -> InlineKeyboardMarkup:
    # В callback_data передаётся индекс предмета (лимит Telegram — 64 байта, кириллица занимает 2 байта на символ)
    rows: list[list[InlineKeyboardButton]] = []
    for idx, subj in enumerate(available):
        mark = "✅" if subj in selected else "⬜"
        rows.append([InlineKeyboardButton(text=f"{mark} {subj}", callback_data=f"subj:toggle:{idx}")])
    rows.append([InlineKeyboardButton(text="☑️ Выбрать все", callback_data="subj:all")])
    rows.append([InlineKeyboardButton(text="✔️ Готово", callback_data=back_cb)])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def olympiad_card_kb(ol_id: int, is_registered: bool, stages: list[dict[str, Any]], back_cb: str = "menu:catalog:0") -> InlineKeyboardMarkup:
    rows: list[list[InlineKeyboardButton]] = [
        [
            InlineKeyboardButton(
                text="✅ Зарегистрирован" if is_registered else "📝 Я зарегистрировался",
                callback_data=f"reg:toggle:{ol_id}:{back_cb}",
            )
        ]
    ]
    for st in stages:
        icon = "✅" if st.get("is_completed") else "⬜"
        rows.append([InlineKeyboardButton(text=f"{icon} {st['stage_name']}", callback_data=f"stage:toggle:{st['id']}:{ol_id}")])
    rows.append([InlineKeyboardButton(text="⬅️ Назад", callback_data=back_cb)])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def admin_panel_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="📥 Импорт 2027.xlsx", callback_data="admin:import")],
            [InlineKeyboardButton(text="⚠️ Неоднозначные даты", callback_data="admin:ambiguous")],
            [InlineKeyboardButton(text="💾 Бэкап БД", callback_data="admin:backup")],
            [InlineKeyboardButton(text="📣 Рассылка", callback_data="admin:broadcast")],
            [InlineKeyboardButton(text="⬅️ Меню", callback_data="menu:main")],
        ]
    )


# ==============================================================================
# 10. ОБРАБОТЧИКИ
# ==============================================================================
dp = Dispatcher(storage=MemoryStorage())


@dp.errors(ExceptionTypeFilter(TelegramBadRequest))
async def on_bad_request(event: ErrorEvent) -> bool:
    """Безобидные ошибки Telegram («message is not modified», «query is too old») не должны ронять обработку."""
    logger.warning("TelegramBadRequest (игнорируется): %s", event.exception)
    return True


async def guard_access(bot: Bot, telegram_id: int, full_name: str, username: str | None) -> tuple[bool, dict[str, Any], str | None]:
    user = await ensure_user(telegram_id, full_name, username)
    ok, err_msg = await check_channel_subscription(bot, telegram_id)
    if not ok:
        return False, user, err_msg
    return True, user, None


@dp.message(CommandStart())
async def cmd_start(message: Message, bot: Bot, state: FSMContext) -> None:
    await state.clear()
    tg_user = message.from_user
    if not tg_user:
        return

    user = await ensure_user(tg_user.id, tg_user.full_name, tg_user.username)
    sub_ok, err_msg = await check_channel_subscription(bot, tg_user.id)
    if not sub_ok:
        await message.answer(
            f"Для доступа подпишитесь на <b>@{esc(CHANNEL_USERNAME)}</b> ({esc(CHANNEL_URL)}).\n\n{esc(err_msg or '')}".strip(),
            reply_markup=subscription_required_kb(),
        )
        return

    if user.get("grade") is None:
        await message.answer("Выбери класс (7–11):", reply_markup=grade_kb("menu:main"))
        return

    sel = user.get("selected_subjects") or await get_all_subjects()
    await message.answer(
        f"Привет, <b>{esc(tg_user.full_name)}</b>! Класс: <b>{user['grade']}</b>.\n"
        f"Предметы: <b>{esc(', '.join(sel))}</b>",
        reply_markup=main_menu_kb(tg_user.id),
    )


@dp.callback_query(F.data == "sub:check")
async def cb_sub_check(callback: CallbackQuery, bot: Bot) -> None:
    tg_user = callback.from_user
    if not tg_user or not callback.message:
        return
    user = await ensure_user(tg_user.id, tg_user.full_name, tg_user.username)
    sub_ok, err_msg = await check_channel_subscription(bot, tg_user.id)
    if not sub_ok:
        await callback.answer("Подписка не найдена", show_alert=True)
        await callback.message.edit_text(f"❌ {esc(err_msg or 'Подпишитесь на канал.')}", reply_markup=subscription_required_kb())
        return

    await callback.answer("Подписка подтверждена!")
    if user.get("grade") is None:
        await callback.message.edit_text("✅ Подписка подтверждена! Выбери класс (7–11):", reply_markup=grade_kb("menu:main"))
    else:
        await callback.message.edit_text("✅ Подписка подтверждена!", reply_markup=main_menu_kb(tg_user.id))


@dp.callback_query(F.data == "menu:main")
async def cb_menu_main(callback: CallbackQuery, bot: Bot, state: FSMContext) -> None:
    await state.clear()
    tg_user = callback.from_user
    if not tg_user or not callback.message:
        return
    ok, user, err = await guard_access(bot, tg_user.id, tg_user.full_name, tg_user.username)
    if not ok:
        await callback.answer("Нужна подписка", show_alert=True)
        await callback.message.edit_text(esc(err or "Подпишитесь на канал."), reply_markup=subscription_required_kb())
        return
    await callback.answer()
    g_txt = f"{user['grade']} кл." if user.get("grade") else "не выбран"
    sel = ", ".join(user.get("selected_subjects") or [])
    await callback.message.edit_text(
        f"<b>Меню</b> ({g_txt})\nПредметы: <b>{esc(sel)}</b>",
        reply_markup=main_menu_kb(tg_user.id),
    )


@dp.callback_query(F.data.startswith("menu:catalog:"))
async def cb_catalog(callback: CallbackQuery, bot: Bot) -> None:
    tg_user = callback.from_user
    if not tg_user or not callback.message:
        return
    ok, user, err = await guard_access(bot, tg_user.id, tg_user.full_name, tg_user.username)
    if not ok:
        await callback.answer("Нужна подписка", show_alert=True)
        await callback.message.edit_text(esc(err or "Подпишитесь на канал."), reply_markup=subscription_required_kb())
        return
    await callback.answer()
    ignore_grade = callback.data.split(":")[2] == "1"
    selected_subjects = user.get("selected_subjects") or await get_all_subjects()
    grade = user.get("grade")

    items = await get_olympiads_for_user(subjects=selected_subjects, grade=grade, ignore_grade=ignore_grade)
    rows: list[list[InlineKeyboardButton]] = [
        [InlineKeyboardButton(text=f"[{ol['subject'][:5]}] {ol['title'][:42]}", callback_data=f"ol:view:{ol['id']}:{callback.data}")]
        for ol in items[:25]
    ]
    rows.append([InlineKeyboardButton(text="🎯 Выбрать предметы", callback_data="subj:picker")])
    if grade:
        alt_cb = "menu:catalog:0" if ignore_grade else "menu:catalog:1"
        alt_txt = f"Только {grade} класс" if ignore_grade else "Показать все классы (7–11)"
        rows.append([InlineKeyboardButton(text=alt_txt, callback_data=alt_cb)])
    rows.append([InlineKeyboardButton(text="⬅️ Назад", callback_data="menu:main")])

    subj_label = ", ".join(selected_subjects)
    grade_label = "все классы (7–11)" if (ignore_grade or not grade) else f"{grade} класс"
    extra = f"\n<i>Показаны первые 25 — уточните через «🔎 Поиск».</i>" if len(items) > 25 else ""
    await callback.message.edit_text(
        f"📚 <b>Каталог</b> | Предметы: <b>{esc(subj_label)}</b> | {grade_label}\nНайдено: <b>{len(items)}</b>{extra}",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=rows),
    )


@dp.callback_query(F.data == "subj:picker")
async def cb_subj_picker(callback: CallbackQuery, bot: Bot) -> None:
    tg_user = callback.from_user
    if not tg_user or not callback.message:
        return
    ok, user, err = await guard_access(bot, tg_user.id, tg_user.full_name, tg_user.username)
    if not ok:
        await callback.answer("Нужна подписка", show_alert=True)
        await callback.message.edit_text(esc(err or "Подпишитесь на канал."), reply_markup=subscription_required_kb())
        return
    await callback.answer()
    avail = await get_all_subjects()
    selected = user.get("selected_subjects") or avail
    await callback.message.edit_text(
        "Выберите предметы (при выборе сразу придут списки олимпиад по регистрации):",
        reply_markup=subjects_multi_kb(avail, selected, back_cb="menu:catalog:0"),
    )


@dp.callback_query(F.data.startswith("subj:toggle:"))
async def cb_subj_toggle(callback: CallbackQuery, bot: Bot) -> None:
    tg_user = callback.from_user
    if not tg_user or not callback.message:
        return
    ok, user, _ = await guard_access(bot, tg_user.id, tg_user.full_name, tg_user.username)
    if not ok:
        await callback.answer("Нужна подписка", show_alert=True)
        return
    avail = await get_all_subjects()
    try:
        idx = int(callback.data.split(":")[2])
        subj = avail[idx]
    except (ValueError, IndexError):
        await callback.answer("Список предметов обновился, откройте выбор заново", show_alert=True)
        return
    selected = await toggle_user_subject(tg_user.id, subj)
    await callback.answer("Сохранено")
    await callback.message.edit_reply_markup(reply_markup=subjects_multi_kb(avail, selected, back_cb="menu:catalog:0"))
    await send_split_registration_broadcasts(
        message_or_bot=bot,
        chat_id=callback.message.chat.id,
        subjects=selected,
        grade=user.get("grade"),
    )


@dp.callback_query(F.data == "subj:all")
async def cb_subj_all(callback: CallbackQuery, bot: Bot) -> None:
    tg_user = callback.from_user
    if not tg_user or not callback.message:
        return
    ok, user, _ = await guard_access(bot, tg_user.id, tg_user.full_name, tg_user.username)
    if not ok:
        await callback.answer("Нужна подписка", show_alert=True)
        return
    selected = await set_all_user_subjects(tg_user.id)
    avail = await get_all_subjects()
    await callback.answer("Выбраны все предметы")
    await callback.message.edit_reply_markup(reply_markup=subjects_multi_kb(avail, selected, back_cb="menu:catalog:0"))
    await send_split_registration_broadcasts(
        message_or_bot=bot,
        chat_id=callback.message.chat.id,
        subjects=selected,
        grade=user.get("grade"),
    )


@dp.callback_query(F.data.startswith("ol:view:"))
async def cb_ol_view(callback: CallbackQuery, bot: Bot) -> None:
    tg_user = callback.from_user
    if not tg_user or not callback.message:
        return
    ok, user, err = await guard_access(bot, tg_user.id, tg_user.full_name, tg_user.username)
    if not ok:
        await callback.answer("Нужна подписка", show_alert=True)
        await callback.message.edit_text(esc(err or "Подпишитесь на канал."), reply_markup=subscription_required_kb())
        return
    await callback.answer()
    _, _, ol_s, back_cb = callback.data.split(":", 3)
    ol_id = int(ol_s)
    card = await get_olympiad_card_data(ol_id, int(user["id"]))
    if not card:
        await callback.message.edit_text("Олимпиада не найдена.", reply_markup=main_menu_kb(tg_user.id))
        return
    await callback.message.edit_text(
        format_olympiad_card_text(card),
        reply_markup=olympiad_card_kb(ol_id, card["is_registered"], card["stages"], back_cb=back_cb),
    )


@dp.callback_query(F.data.startswith("reg:toggle:"))
async def cb_reg_toggle(callback: CallbackQuery, bot: Bot) -> None:
    tg_user = callback.from_user
    if not tg_user or not callback.message:
        return
    ok, user, _ = await guard_access(bot, tg_user.id, tg_user.full_name, tg_user.username)
    if not ok:
        await callback.answer("Нужна подписка", show_alert=True)
        return
    _, _, ol_s, back_cb = callback.data.split(":", 3)
    ol_id = int(ol_s)
    registered = await toggle_olympiad_registration(int(user["id"]), ol_id)
    await callback.answer("Добавлено в «Мои олимпиады» ✅" if registered else "Отметка регистрации снята")
    card = await get_olympiad_card_data(ol_id, int(user["id"]))
    if card:
        await callback.message.edit_text(
            format_olympiad_card_text(card),
            reply_markup=olympiad_card_kb(ol_id, card["is_registered"], card["stages"], back_cb=back_cb),
        )


@dp.callback_query(F.data.startswith("stage:toggle:"))
async def cb_stage_toggle(callback: CallbackQuery, bot: Bot) -> None:
    tg_user = callback.from_user
    if not tg_user or not callback.message:
        return
    ok, user, _ = await guard_access(bot, tg_user.id, tg_user.full_name, tg_user.username)
    if not ok:
        await callback.answer("Нужна подписка", show_alert=True)
        return
    _, _, st_s, ol_s = callback.data.split(":")
    await toggle_stage_progress(int(user["id"]), int(st_s))
    await callback.answer("Обновлено")
    card = await get_olympiad_card_data(int(ol_s), int(user["id"]))
    if card:
        await callback.message.edit_text(
            format_olympiad_card_text(card),
            reply_markup=olympiad_card_kb(int(ol_s), card["is_registered"], card["stages"], back_cb="menu:my"),
        )


@dp.callback_query(F.data == "menu:search")
async def cb_menu_search(callback: CallbackQuery, bot: Bot, state: FSMContext) -> None:
    tg_user = callback.from_user
    if not tg_user or not callback.message:
        return
    ok, _, _ = await guard_access(bot, tg_user.id, tg_user.full_name, tg_user.username)
    if not ok:
        await callback.answer("Нужна подписка", show_alert=True)
        return
    await callback.answer()
    await state.set_state(UserStates.waiting_search_query)
    await callback.message.edit_text(
        "🔎 Введите название, предмет, уровень, класс (7–11) или слово:",
        reply_markup=back_kb("menu:main", "⬅️ Отмена"),
    )


@dp.message(UserStates.waiting_search_query)
async def msg_search_query(message: Message, bot: Bot, state: FSMContext) -> None:
    tg_user = message.from_user
    if not tg_user or not message.text:
        return
    ok, _, _ = await guard_access(bot, tg_user.id, tg_user.full_name, tg_user.username)
    if not ok:
        await state.clear()
        return
    q = message.text.strip()
    await state.clear()
    matches = await search_olympiads(q)
    if not matches:
        await message.answer(f"Ничего не найдено по запросу «{esc(q)}».", reply_markup=main_menu_kb(tg_user.id))
        return
    rows = [
        [InlineKeyboardButton(text=f"[{m['subject'][:5]}] {m['title'][:42]}", callback_data=f"ol:view:{m['id']}:menu:main")]
        for m in matches[:25]
    ]
    rows.append([InlineKeyboardButton(text="⬅️ Меню", callback_data="menu:main")])
    await message.answer(f"Найдено: <b>{len(matches)}</b>", reply_markup=InlineKeyboardMarkup(inline_keyboard=rows))


@dp.callback_query(F.data == "menu:my")
async def cb_menu_my(callback: CallbackQuery, bot: Bot) -> None:
    tg_user = callback.from_user
    if not tg_user or not callback.message:
        return
    ok, user, _ = await guard_access(bot, tg_user.id, tg_user.full_name, tg_user.username)
    if not ok:
        await callback.answer("Нужна подписка", show_alert=True)
        return
    await callback.answer()
    regs = await get_user_registered_olympiads(int(user["id"]))
    if not regs:
        await callback.message.edit_text(
            "📝 Вы ещё не отметили ни одной олимпиады.\nОткройте каталог и нажмите «📝 Я зарегистрировался» в нужной карточке.",
            reply_markup=InlineKeyboardMarkup(
                inline_keyboard=[
                    [InlineKeyboardButton(text="📚 Каталог", callback_data="menu:catalog:0")],
                    [InlineKeyboardButton(text="⬅️ Назад", callback_data="menu:main")],
                ]
            ),
        )
        return

    lines = [f"📝 <b>Мои олимпиады</b> (зарегистрирован: <b>{len(regs)}</b>)\n"]
    rows: list[list[InlineKeyboardButton]] = []
    for ol in regs[:30]:
        done = sum(1 for s in ol.get("stages", []) if s.get("is_completed"))
        total_s = len(ol.get("stages", []))
        next_st = next((s for s in ol.get("stages", []) if not s.get("is_completed")), None)
        next_txt = f" | далее: {esc(next_st['stage_name'])} ({esc(next_st['raw_date_text'])})" if next_st else ""
        lines.append(f"✅ <b>{esc(ol['title'])}</b> ({done}/{total_s}){next_txt}")
        rows.append([InlineKeyboardButton(text=f"Этапы: {ol['title'][:40]}", callback_data=f"ol:view:{ol['id']}:menu:my")])
    rows.append([InlineKeyboardButton(text="⬅️ Назад", callback_data="menu:main")])
    await callback.message.edit_text(join_limited(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=rows))


@dp.callback_query(F.data == "menu:history")
async def cb_menu_history(callback: CallbackQuery, bot: Bot) -> None:
    tg_user = callback.from_user
    if not tg_user or not callback.message:
        return
    ok, user, _ = await guard_access(bot, tg_user.id, tg_user.full_name, tg_user.username)
    if not ok:
        await callback.answer("Нужна подписка", show_alert=True)
        return
    await callback.answer()
    entries = await get_user_history(int(user["id"]))
    lines = ["🗂 <b>История участия</b>\n"]
    if not entries:
        lines.append("Записей пока нет.")
    else:
        for h in entries[:15]:
            fin = " · финал" if h.get("is_final") else ""
            res = f" · {esc(h['result_text'])}" if h.get("result_text") else ""
            lines.append(
                f"• <b>{h['participation_year']}</b> — {esc(h['olympiad_title'])} ({esc(h['subject'])}, {esc(h['stage_reached'])}{fin}{res})"
            )
    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="➕ Добавить участие", callback_data="hist:add")],
            [InlineKeyboardButton(text="⬅️ Назад", callback_data="menu:main")],
        ]
    )
    await callback.message.edit_text(join_limited(lines), reply_markup=kb)


@dp.callback_query(F.data == "hist:add")
async def cb_hist_add(callback: CallbackQuery, bot: Bot, state: FSMContext) -> None:
    tg_user = callback.from_user
    if not tg_user or not callback.message:
        return
    ok, _, _ = await guard_access(bot, tg_user.id, tg_user.full_name, tg_user.username)
    if not ok:
        await callback.answer("Нужна подписка", show_alert=True)
        return
    await callback.answer()
    await state.set_state(UserStates.waiting_history_title)
    await callback.message.edit_text("Введите <b>название олимпиады</b>:")


@dp.message(UserStates.waiting_history_title)
async def msg_hist_title(message: Message, state: FSMContext) -> None:
    if not message.text or not message.text.strip():
        await message.answer("Введите название:")
        return
    await state.update_data(hist_title=message.text.strip())
    await state.set_state(UserStates.waiting_history_subject)
    await message.answer("Введите <b>предмет</b>:")


@dp.message(UserStates.waiting_history_subject)
async def msg_hist_subject(message: Message, state: FSMContext) -> None:
    if not message.text or not message.text.strip():
        await message.answer("Введите предмет:")
        return
    await state.update_data(hist_subject=message.text.strip())
    await state.set_state(UserStates.waiting_history_year)
    await message.answer("Введите <b>год участия</b> (например, 2025):")


@dp.message(UserStates.waiting_history_year)
async def msg_hist_year(message: Message, state: FSMContext) -> None:
    txt = (message.text or "").strip()
    if not txt.isdigit() or not (1990 <= int(txt) <= 2100):
        await message.answer("Введите год числом (например, 2025):")
        return
    await state.update_data(hist_year=int(txt))
    await state.set_state(UserStates.waiting_history_stage)
    await message.answer(
        "Выберите или введите <b>этап</b>:",
        reply_markup=InlineKeyboardMarkup(
            inline_keyboard=[[InlineKeyboardButton(text=ch, callback_data=f"hist:stage:{i}")] for i, ch in enumerate(STAGE_PARTICIPATION_CHOICES)]
        ),
    )


@dp.callback_query(UserStates.waiting_history_stage, F.data.startswith("hist:stage:"))
async def cb_hist_stage(callback: CallbackQuery, state: FSMContext) -> None:
    try:
        chosen = STAGE_PARTICIPATION_CHOICES[int(callback.data.split(":")[2])]
    except (ValueError, IndexError):
        await callback.answer("Некорректный выбор", show_alert=True)
        return
    is_final = "финал" in chosen.lower() or "заключительный" in chosen.lower()
    await callback.answer()
    await state.update_data(hist_stage=chosen, hist_is_final=is_final)
    await state.set_state(UserStates.waiting_history_result)
    await callback.message.edit_text("Укажите <b>результат</b> (или «-»):")


@dp.message(UserStates.waiting_history_stage)
async def msg_hist_stage(message: Message, state: FSMContext) -> None:
    txt = (message.text or "").strip()
    if not txt:
        await message.answer("Укажите этап:")
        return
    is_final = "финал" in txt.lower() or "заключ" in txt.lower()
    await state.update_data(hist_stage=txt, hist_is_final=is_final)
    await state.set_state(UserStates.waiting_history_result)
    await message.answer("Укажите <b>результат</b> (или «-»):")


@dp.message(UserStates.waiting_history_result)
async def msg_hist_result(message: Message, state: FSMContext) -> None:
    txt = (message.text or "").strip()
    await state.update_data(hist_result=None if txt in ("", "-") else txt)
    await state.set_state(UserStates.waiting_history_note)
    await message.answer("Добавьте <b>заметку</b> (или «-»):")


@dp.message(UserStates.waiting_history_note)
async def msg_hist_note(message: Message, bot: Bot, state: FSMContext) -> None:
    tg_user = message.from_user
    if not tg_user:
        return
    ok, user, _ = await guard_access(bot, tg_user.id, tg_user.full_name, tg_user.username)
    if not ok:
        await state.clear()
        return
    txt = (message.text or "").strip()
    data = await state.get_data()
    await state.clear()
    try:
        await add_history_entry(
            user_db_id=int(user["id"]),
            olympiad_title=data["hist_title"],
            subject=data["hist_subject"],
            participation_year=int(data["hist_year"]),
            stage_reached=data["hist_stage"],
            result_text=data.get("hist_result"),
            note_text=None if txt in ("", "-") else txt,
            is_final=bool(data.get("hist_is_final")),
        )
    except (KeyError, ValueError) as exc:
        await message.answer(f"❌ Не удалось сохранить: {esc(exc)}", reply_markup=main_menu_kb(tg_user.id))
        return
    await message.answer("✅ Сохранено в историю!", reply_markup=main_menu_kb(tg_user.id))


async def render_profile(callback: CallbackQuery, user: dict[str, Any]) -> None:
    st = await get_user_statistics(int(user["id"]))
    rem_on = await is_reminders_enabled(int(user["id"]))

    subj_stat = ", ".join(f"{k}: {v}" for k, v in st["by_subject"].items()) or "нет"
    year_stat = ", ".join(f"{k}—{v}" for k, v in st["by_year"].items()) or "нет"
    sel_subjs = ", ".join(user.get("selected_subjects") or [])

    text = (
        "👤 <b>Профиль и рассылка</b>\n"
        f"Класс: <b>{user.get('grade') or 'не выбран'}</b>\n"
        f"Предметы: <b>{esc(sel_subjs)}</b>\n"
        f"Срочные (&lt;7 дн. до регистрации) и плановые напоминания: <b>{'вкл ✅' if rem_on else 'выкл ❌'}</b> ({esc(TZ_NAME)})\n\n"
        f"Каталог: <b>{st['catalog_total']}</b> · Зарегистрирован: <b>{st['registered_count']}</b> · Этапов пройдено: <b>{st['completed_stages']}</b>\n"
        f"Участий в истории: <b>{st['history_count']}</b> (олимпиад: <b>{st['participated_olympiads_count']}</b>, финалов: <b>{st['finals_count']}</b>)\n"
        f"По предметам: {esc(subj_stat)}\n"
        f"По годам: {esc(year_stat)}"
    )
    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="🎓 Класс (7–11)", callback_data="prof:grade"),
                InlineKeyboardButton(text="🎯 Предметы", callback_data="subj:picker"),
            ],
            [InlineKeyboardButton(text="🔔 Напоминания: вкл/выкл", callback_data="rem:toggle")],
            [InlineKeyboardButton(text="⬅️ Назад", callback_data="menu:main")],
        ]
    )
    await callback.message.edit_text(text[:4000], reply_markup=kb)


@dp.callback_query(F.data == "menu:profile")
async def cb_menu_profile(callback: CallbackQuery, bot: Bot) -> None:
    tg_user = callback.from_user
    if not tg_user or not callback.message:
        return
    ok, user, _ = await guard_access(bot, tg_user.id, tg_user.full_name, tg_user.username)
    if not ok:
        await callback.answer("Нужна подписка", show_alert=True)
        return
    await callback.answer()
    await render_profile(callback, user)


@dp.callback_query(F.data == "prof:grade")
async def cb_prof_grade(callback: CallbackQuery, bot: Bot) -> None:
    tg_user = callback.from_user
    if not tg_user or not callback.message:
        return
    ok, _, _ = await guard_access(bot, tg_user.id, tg_user.full_name, tg_user.username)
    if not ok:
        await callback.answer("Нужна подписка", show_alert=True)
        return
    await callback.answer()
    await callback.message.edit_text("Выберите класс (7–11):", reply_markup=grade_kb("menu:profile"))


@dp.callback_query(F.data.startswith("grade:set:"))
async def cb_grade_set(callback: CallbackQuery, bot: Bot) -> None:
    tg_user = callback.from_user
    if not tg_user or not callback.message:
        return
    ok, _, _ = await guard_access(bot, tg_user.id, tg_user.full_name, tg_user.username)
    if not ok:
        await callback.answer("Нужна подписка", show_alert=True)
        return
    try:
        grade_num = int(callback.data.split(":")[2])
        await set_user_grade(tg_user.id, grade_num)
    except ValueError:
        await callback.answer("Некорректный класс", show_alert=True)
        return
    await callback.answer(f"Класс: {grade_num}")
    await callback.message.edit_text(f"✅ Выбран <b>{grade_num} класс</b>.", reply_markup=main_menu_kb(tg_user.id))


@dp.callback_query(F.data == "rem:toggle")
async def cb_rem_toggle(callback: CallbackQuery, bot: Bot) -> None:
    tg_user = callback.from_user
    if not tg_user or not callback.message:
        return
    ok, user, _ = await guard_access(bot, tg_user.id, tg_user.full_name, tg_user.username)
    if not ok:
        await callback.answer("Нужна подписка", show_alert=True)
        return
    cur = await is_reminders_enabled(int(user["id"]))
    await set_reminders_enabled(int(user["id"]), not cur)
    await callback.answer("Напоминания выключены" if cur else "Напоминания включены")
    await render_profile(callback, user)


@dp.callback_query(F.data == "menu:help")
async def cb_menu_help(callback: CallbackQuery, bot: Bot) -> None:
    tg_user = callback.from_user
    if not tg_user or not callback.message:
        return
    ok, _, _ = await guard_access(bot, tg_user.id, tg_user.full_name, tg_user.username)
    if not ok:
        await callback.answer("Нужна подписка", show_alert=True)
        return
    await callback.answer()
    await callback.message.edit_text(
        "ℹ️ <b>Как пользоваться</b>\n"
        "• При выборе предметов бот присылает два отдельных списка: где регистрация идёт и где закончилась.\n"
        "• Если до конца регистрации по выбранному предмету осталось меньше 7 дней, бот присылает срочную рассылку.\n"
        "• В карточке олимпиады кнопка <b>«📝 Я зарегистрировался»</b> добавляет её в <b>Мои олимпиады</b>.",
        reply_markup=back_kb("menu:main"),
    )


# --- АДМИН ---
@dp.callback_query(F.data == "menu:admin")
async def cb_menu_admin(callback: CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    tg_user = callback.from_user
    if not tg_user or not is_admin(tg_user.id) or not callback.message:
        await callback.answer("Нет доступа", show_alert=True)
        return
    await callback.answer()
    st = await get_admin_stats()
    await callback.message.edit_text(
        f"🛠 <b>Админ</b> | Польз.: <b>{st['users_total']}</b> | Олимпиад: <b>{st['olympiads_active']}</b> | Неоднозн.: <b>{st['ambiguous_stages']}</b>",
        reply_markup=admin_panel_kb(),
    )


@dp.callback_query(F.data == "admin:import")
async def cb_admin_import(callback: CallbackQuery) -> None:
    tg_user = callback.from_user
    if not tg_user or not is_admin(tg_user.id) or not callback.message:
        await callback.answer("Нет доступа", show_alert=True)
        return
    await callback.answer("Импорт...")
    rep = await import_excel_to_db(EXCEL_PATH, DB_PATH, tg_user.id, make_backup=True)
    text = (
        f"📥 Олимпиад: <b>{rep['found_olympiads']}</b> · Этапов: <b>{rep['imported_stages']}</b> · Ссылок: <b>{rep['found_links']}</b>\n"
        f"Пропущено: <b>{rep['skipped_rows']}</b> · Ошибок: <b>{rep['errors_count']}</b> · Неоднозначных дат: <b>{len(rep['ambiguous_dates'])}</b>"
    )
    if rep["errors"]:
        text += "\n\n" + esc("\n".join(rep["errors"][:3]))[:600]
    await callback.message.edit_text(text, reply_markup=back_kb("menu:admin"))


@dp.callback_query(F.data == "admin:ambiguous")
async def cb_admin_ambiguous(callback: CallbackQuery) -> None:
    tg_user = callback.from_user
    if not tg_user or not is_admin(tg_user.id) or not callback.message:
        await callback.answer("Нет доступа", show_alert=True)
        return
    await callback.answer()
    items = await get_ambiguous_stages_list(limit=8)
    if not items:
        await callback.message.edit_text("Неоднозначных этапов нет.", reply_markup=back_kb("menu:admin"))
        return
    lines = ["⚠️ <b>Неоднозначные этапы:</b>"]
    rows: list[list[InlineKeyboardButton]] = []
    for it in items:
        lines.append(
            f"#{it['id']} [{esc(it['subject'])}] {esc(it['olympiad_title'][:25])} — {esc(it['stage_name'])}: «{esc(it['raw_date_text'])}»"
        )
        rows.append([InlineKeyboardButton(text=f"Дата для #{it['id']}", callback_data=f"admin:fix:{it['id']}")])
    rows.append([InlineKeyboardButton(text="⬅️ Назад", callback_data="menu:admin")])
    await callback.message.edit_text(join_limited(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=rows))


@dp.callback_query(F.data.startswith("admin:fix:"))
async def cb_admin_fix_prompt(callback: CallbackQuery, state: FSMContext) -> None:
    tg_user = callback.from_user
    if not tg_user or not is_admin(tg_user.id) or not callback.message:
        await callback.answer("Нет доступа", show_alert=True)
        return
    st_id = int(callback.data.split(":")[2])
    await callback.answer()
    await state.set_state(UserStates.waiting_admin_fix_date)
    await state.update_data(fix_stage_id=st_id)
    await callback.message.edit_text(f"Введите дату для #{st_id} в формате ГГГГ-ММ-ДД:")


@dp.message(UserStates.waiting_admin_fix_date)
async def msg_admin_fix_date(message: Message, state: FSMContext) -> None:
    tg_user = message.from_user
    if not tg_user or not is_admin(tg_user.id):
        await state.clear()
        return
    st_id = int((await state.get_data())["fix_stage_id"])
    await state.clear()
    ok, msg = await fix_stage_date_manual(st_id, message.text or "")
    await message.answer(("✅ " if ok else "❌ ") + esc(msg), reply_markup=back_kb("menu:admin", "⬅️ Админ"))


@dp.callback_query(F.data == "admin:backup")
async def cb_admin_backup(callback: CallbackQuery) -> None:
    tg_user = callback.from_user
    if not tg_user or not is_admin(tg_user.id) or not callback.message:
        await callback.answer("Нет доступа", show_alert=True)
        return
    b_path = create_db_backup(DB_PATH, BACKUP_DIR)
    await callback.answer("Готово", show_alert=True)
    await callback.message.edit_text(f"💾 Бэкап: <code>{esc(b_path or 'нет БД')}</code>", reply_markup=back_kb("menu:admin"))


@dp.callback_query(F.data == "admin:broadcast")
async def cb_admin_broadcast_prompt(callback: CallbackQuery, state: FSMContext) -> None:
    tg_user = callback.from_user
    if not tg_user or not is_admin(tg_user.id) or not callback.message:
        await callback.answer("Нет доступа", show_alert=True)
        return
    await callback.answer()
    await state.set_state(UserStates.waiting_admin_broadcast)
    await callback.message.edit_text("Введите текст рассылки (или /cancel):")


@dp.message(UserStates.waiting_admin_broadcast)
async def msg_admin_broadcast(message: Message, bot: Bot, state: FSMContext) -> None:
    tg_user = message.from_user
    if not tg_user or not is_admin(tg_user.id):
        await state.clear()
        return
    text = (message.text or "").strip()
    await state.clear()
    if not text or text == "/cancel":
        await message.answer("Отменено.", reply_markup=admin_panel_kb())
        return
    conn = await get_db(DB_PATH)
    try:
        cur = await conn.execute("SELECT telegram_id FROM users WHERE is_subscribed = 1")
        user_ids = [int(r["telegram_id"]) for r in await cur.fetchall()]
    finally:
        await conn.close()
    sent = failed = 0
    safe_text = f"📢 {esc(text)}"[:4000]
    for uid in user_ids:
        try:
            await bot.send_message(chat_id=uid, text=safe_text)
            sent += 1
        except TelegramAPIError:
            failed += 1
        await asyncio.sleep(0.05)
    await message.answer(f"Отправлено: {sent}, ошибок: {failed}.", reply_markup=admin_panel_kb())


# ==============================================================================
# 11. АВТОТЕСТЫ
# ==============================================================================
class OlympiadBotLocalTests(unittest.IsolatedAsyncioTestCase):
    async def asyncSetUp(self) -> None:
        self.tmp_dir = Path(tempfile.mkdtemp(prefix="olympiad_test_"))
        self.db_path = self.tmp_dir / "test.db"
        await init_db(self.db_path)

    async def asyncTearDown(self) -> None:
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_00_parsers(self) -> None:
        d = parse_date_expression("15 октября - 20 ноября")
        self.assertEqual((d["start_date"], d["end_date"], d["is_ambiguous"]), ("2026-10-15", "2026-11-20", 0))
        d = parse_date_expression("1-5 февраля 2027")
        self.assertEqual((d["start_date"], d["end_date"]), ("2027-02-01", "2027-02-05"))
        d = parse_date_expression("середина ноября")
        self.assertEqual(d["is_ambiguous"], 1)
        d = parse_date_expression("31.13.2026")
        self.assertEqual(d["is_ambiguous"], 1)
        g = parse_grades_value("5-11")
        self.assertEqual(g["grades"], [7, 8, 9, 10, 11])
        g = parse_grades_value(datetime.datetime(2026, 11, 5))
        self.assertEqual(g["grades"], [7, 8, 9, 10, 11])

    async def test_01_urgent_reg_under_7_days(self) -> None:
        if not EXCEL_PATH.exists():
            self.skipTest("Нет 2027.xlsx")
        rep = await import_excel_to_db(EXCEL_PATH, self.db_path, ADMIN_ID, make_backup=False)
        self.assertGreater(rep["found_olympiads"], 80)

        await ensure_user(50001, "Тест Срочной", "test_urg", db_path=self.db_path)
        await set_user_grade(50001, 11, db_path=self.db_path)
        conn = await get_db(self.db_path)
        try:
            await conn.execute(
                "UPDATE users SET is_subscribed = 1, selected_subjects_json = ? WHERE telegram_id = 50001",
                (json.dumps(["Химия"], ensure_ascii=False),),
            )
            await conn.commit()  # без commit изменения терялись при закрытии соединения
            cur = await conn.execute(
                """
                SELECT s.id, s.parsed_start_date, s.parsed_end_date
                FROM stages s JOIN olympiads o ON o.id = s.olympiad_id
                WHERE o.subject = 'Химия' AND s.stage_code = 'registration'
                  AND s.is_ambiguous = 0 AND s.parsed_end_date IS NOT NULL
                  AND (o.grades_known = 0 OR o.grades_json LIKE '%11%')
                LIMIT 1
                """
            )
            st = await cur.fetchone()
        finally:
            await conn.close()

        self.assertIsNotNone(st)
        ref_day = datetime.date.fromisoformat(st["parsed_end_date"]) - datetime.timedelta(days=5)
        due = await collect_due_reminders(today_date=ref_day, db_path=self.db_path)
        matched = [d for d in due if d["telegram_id"] == 50001 and d["stage_id"] == st["id"]]
        self.assertTrue(matched and matched[0]["is_urgent_registration"])


def run_local_tests() -> None:
    suite = unittest.defaultTestLoader.loadTestsFromTestCase(OlympiadBotLocalTests)
    res = unittest.TextTestRunner(verbosity=2).run(suite)
    if not res.wasSuccessful():
        sys.exit(1)


# ==============================================================================
# 12. ЗАПУСК
# ==============================================================================
async def main_async() -> None:
    if not BOT_TOKEN:
        logger.critical("Не задан BOT_TOKEN в .env")
        sys.exit(1)

    await init_db(DB_PATH)
    create_db_backup(DB_PATH, BACKUP_DIR)

    if EXCEL_PATH.exists():
        rep = await import_excel_to_db(EXCEL_PATH, DB_PATH, ADMIN_ID, make_backup=False)
        logger.info(
            "Импорт Excel: олимпиад=%s, этапов=%s, ошибок=%s",
            rep["found_olympiads"], rep["imported_stages"], rep["errors_count"],
        )
    else:
        logger.warning("Файл %s не найден — импорт пропущен.", EXCEL_PATH)

    bot = Bot(
        token=BOT_TOKEN,
        default=DefaultBotProperties(parse_mode=ParseMode.HTML, link_preview_is_disabled=True),
    )
    scheduler = AsyncIOScheduler(timezone=APP_TZ)
    scheduler.add_job(
        run_reminders_tick,
        trigger=IntervalTrigger(minutes=15, timezone=APP_TZ),
        args=[bot, DB_PATH],
        id="olympiad_reminders_job",
        replace_existing=True,
        max_instances=1,
        coalesce=True,
    )
    scheduler.start()

    try:
        if STARTUP_DIGEST_ENABLED:
            try:
                await send_startup_digest(bot, DB_PATH)
            except Exception as exc:
                logger.exception("Ошибка стартовой рассылки: %s", exc)
        await run_reminders_tick(bot=bot, db_path=DB_PATH)
        logger.info("Запуск polling бота...")
        await dp.start_polling(bot, allowed_updates=dp.resolve_used_update_types())
    finally:
        if scheduler.running:
            scheduler.shutdown(wait=False)
        await bot.session.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--test", action="store_true")
    args = parser.parse_args()
    if args.test:
        run_local_tests()
    else:
        try:
            asyncio.run(main_async())
        except KeyboardInterrupt:
            logger.info("Остановлено.")