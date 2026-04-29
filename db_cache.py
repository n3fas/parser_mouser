import sqlite3
import json
import threading
from pathlib import Path
from typing import Optional, Dict, Any, List
from contextlib import contextmanager

DB_PATH = Path(__file__).parent / "mouser_cache.db"
_db_lock = threading.Lock()

@contextmanager
def db_session():
    """Контекстный менеджер для безопасной работы с SQLite."""
    with _db_lock:
        conn = sqlite3.connect(DB_PATH, check_same_thread=False)
        try:
            _init_db(conn)
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

def _init_db(conn):
    """Инициализация таблиц, если они не существуют."""
    statements = [
        "CREATE TABLE IF NOT EXISTS part_cache (pn TEXT PRIMARY KEY, data TEXT, timestamp DATETIME DEFAULT CURRENT_TIMESTAMP)",
        "CREATE TABLE IF NOT EXISTS user_history (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, pn TEXT, timestamp DATETIME DEFAULT CURRENT_TIMESTAMP)",
        "CREATE TABLE IF NOT EXISTS docs_categories (name TEXT PRIMARY KEY)",
        "CREATE TABLE IF NOT EXISTS labeling_categories (name TEXT PRIMARY KEY)",
        "CREATE TABLE IF NOT EXISTS ip_brands (name TEXT PRIMARY KEY)",
        "CREATE TABLE IF NOT EXISTS allowed_users (user_id INTEGER PRIMARY KEY)",
        "CREATE TABLE IF NOT EXISTS translations (text_en TEXT PRIMARY KEY, text_ru TEXT)"
    ]
    for stmt in statements:
        conn.execute(stmt)

# --- Управление пользователями ---

def add_user(user_id: int) -> bool:
    with db_session() as conn:
        try:
            conn.execute("INSERT INTO allowed_users (user_id) VALUES (?)", (user_id,))
            return True
        except sqlite3.IntegrityError:
            return False

def remove_user(user_id: int) -> bool:
    with db_session() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM allowed_users WHERE user_id = ?", (user_id,))
        return cursor.rowcount > 0

def is_user_allowed(user_id: int) -> bool:
    with db_session() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT 1 FROM allowed_users WHERE user_id = ?", (user_id,))
        return cursor.fetchone() is not None

def get_all_users() -> List[int]:
    with db_session() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT user_id FROM allowed_users ORDER BY user_id")
        return [r[0] for r in cursor.fetchall()]

# --- ТРОИС (IP Brands) ---

def clear_ip_brands():
    with db_session() as conn:
        conn.execute("DELETE FROM ip_brands")

def add_ip_brand(name: str) -> bool:
    if not name: return False
    name_clean = name.strip().upper()
    with db_session() as conn:
        try:
            conn.execute("INSERT INTO ip_brands (name) VALUES (?)", (name_clean,))
            return True
        except sqlite3.IntegrityError:
            return False

def is_ip_brand(name: str) -> bool:
    if not name: return False
    name_clean = name.strip().upper()
    with db_session() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT 1 FROM ip_brands WHERE name = ?", (name_clean,))
        return cursor.fetchone() is not None

def get_ip_brands_count() -> int:
    with db_session() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM ip_brands")
        return cursor.fetchone()[0]

def get_all_ip_brands() -> List[str]:
    with db_session() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM ip_brands ORDER BY name")
        return [r[0] for r in cursor.fetchall()]

# --- Категории (Документы и Маркировка) ---

def add_docs_category(name: str) -> bool:
    if not name: return False
    with db_session() as conn:
        try:
            conn.execute("INSERT INTO docs_categories (name) VALUES (?)", (name.strip(),))
            return True
        except sqlite3.IntegrityError:
            return False

def remove_docs_category(name: str) -> bool:
    if not name: return False
    with db_session() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM docs_categories WHERE name = ? COLLATE NOCASE", (name.strip(),))
        return cursor.rowcount > 0

def get_docs_categories(limit: int = 0, offset: int = 0) -> List[str]:
    with db_session() as conn:
        cursor = conn.cursor()
        query = "SELECT name FROM docs_categories ORDER BY name"
        if limit > 0:
            query += f" LIMIT {limit} OFFSET {offset}"
        cursor.execute(query)
        return [r[0] for r in cursor.fetchall()]

def get_docs_categories_count() -> int:
    with db_session() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM docs_categories")
        return cursor.fetchone()[0]

def is_docs_category(category: str) -> bool:
    if not category: return False
    cat_clean = category.strip().lower()
    with db_session() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT 1 FROM docs_categories WHERE name = ? COLLATE NOCASE", (cat_clean,))
        return cursor.fetchone() is not None

def add_labeling_category(name: str) -> bool:
    if not name: return False
    with db_session() as conn:
        try:
            conn.execute("INSERT INTO labeling_categories (name) VALUES (?)", (name.strip(),))
            return True
        except sqlite3.IntegrityError:
            return False

def remove_labeling_category(name: str) -> bool:
    if not name: return False
    with db_session() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM labeling_categories WHERE name = ? COLLATE NOCASE", (name.strip(),))
        return cursor.rowcount > 0

def get_labeling_categories(limit: int = 0, offset: int = 0) -> List[str]:
    with db_session() as conn:
        cursor = conn.cursor()
        query = "SELECT name FROM labeling_categories ORDER BY name"
        if limit > 0:
            query += f" LIMIT {limit} OFFSET {offset}"
        cursor.execute(query)
        return [r[0] for r in cursor.fetchall()]

def get_labeling_categories_count() -> int:
    with db_session() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM labeling_categories")
        return cursor.fetchone()[0]

def is_labeling_category(category: str) -> bool:
    if not category: return False
    cat_clean = category.strip().lower()
    with db_session() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT 1 FROM labeling_categories WHERE name = ? COLLATE NOCASE", (cat_clean,))
        return cursor.fetchone() is not None

# --- История запросов ---

def add_to_history(user_id: int, pn: str):
    if not pn: return
    pn_clean = pn.strip().upper()
    with db_session() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT pn FROM user_history WHERE user_id = ? ORDER BY timestamp DESC LIMIT 1", (user_id,))
        last_row = cursor.fetchone()
        if last_row and last_row[0] == pn_clean:
            return
        conn.execute("INSERT INTO user_history (user_id, pn) VALUES (?, ?)", (user_id, pn_clean))

def get_user_history(user_id: int, limit: int = 50) -> List[str]:
    with db_session() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT pn FROM user_history WHERE user_id = ? ORDER BY timestamp DESC LIMIT ?", (user_id, limit))
        return [r[0] for r in cursor.fetchall()]

def clear_user_history(user_id: int):
    with db_session() as conn:
        conn.execute("DELETE FROM user_history WHERE user_id = ?", (user_id,))

# --- Кэш деталей ---

def get_cached_part(pn: str) -> Optional[Dict[str, Any]]:
    if not pn: return None
    pn_clean = pn.strip().lower().replace(" ", "").replace("+", "")
    with db_session() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT data, datetime(timestamp, 'localtime') FROM part_cache WHERE pn = ?", (pn_clean,))
        row = cursor.fetchone()
        if row:
            try:
                data = json.loads(row[0])
                data["Дата кэширования"] = row[1]
                return data
            except Exception:
                return None
    return None

def save_cached_part(pn: str, data: Dict[str, Any]):
    if not pn or not data: return
    pn_clean = pn.strip().lower().replace(" ", "").replace("+", "")
    with db_session() as conn:
        conn.execute("INSERT OR REPLACE INTO part_cache (pn, data) VALUES (?, ?)", (pn_clean, json.dumps(data, ensure_ascii=False)))

def get_cache_stats() -> int:
    with db_session() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM part_cache")
        return cursor.fetchone()[0]

def delete_cached_part(pn: str):
    if not pn: return
    pn_clean = pn.strip().lower().replace(" ", "").replace("+", "")
    with db_session() as conn:
        conn.execute("DELETE FROM part_cache WHERE pn = ?", (pn_clean,))

def get_all_cached_pns() -> List[str]:
    with db_session() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT pn FROM part_cache")
        return [r[0] for r in cursor.fetchall()]

# --- Кэш переводов ---

def get_translation(text_en: str) -> Optional[str]:
    """Возвращает кэшированный перевод."""
    if not text_en: return None
    with db_session() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT text_ru FROM translations WHERE text_en = ?", (text_en.strip(),))
        row = cursor.fetchone()
        return row[0] if row else None

def save_translation(text_en: str, text_ru: str):
    """Сохраняет перевод в базу данных."""
    if not text_en or not text_ru: return
    with db_session() as conn:
        conn.execute("INSERT OR REPLACE INTO translations (text_en, text_ru) VALUES (?, ?)", (text_en.strip(), text_ru.strip()))

