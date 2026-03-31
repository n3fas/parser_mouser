import sqlite3
import json
import threading
from pathlib import Path
from typing import Optional, Dict, Any

DB_PATH = Path(__file__).parent / "mouser_cache.db"
_db_lock = threading.Lock()

def _get_connection():
    # check_same_thread=False позволяет использовать одно подключение (или пересоздавать его) в разных потоках
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute(
        "CREATE TABLE IF NOT EXISTS part_cache ("
        "pn TEXT PRIMARY KEY, "
        "data TEXT, "
        "timestamp DATETIME DEFAULT CURRENT_TIMESTAMP"
        ")"
    )
    conn.execute(
        "CREATE TABLE IF NOT EXISTS user_history ("
        "id INTEGER PRIMARY KEY AUTOINCREMENT, "
        "user_id INTEGER, "
        "pn TEXT, "
        "timestamp DATETIME DEFAULT CURRENT_TIMESTAMP"
        ")"
    )
    conn.execute(
        "CREATE TABLE IF NOT EXISTS docs_categories ("
        "name TEXT PRIMARY KEY"
        ")"
    )
    conn.execute(
        "CREATE TABLE IF NOT EXISTS labeling_categories ("
        "name TEXT PRIMARY KEY"
        ")"
    )
    conn.execute(
        "CREATE TABLE IF NOT EXISTS ip_brands ("
        "name TEXT PRIMARY KEY"
        ")"
    )
    return conn

def clear_ip_brands():
    with _db_lock:
        with _get_connection() as conn:
            conn.execute("DELETE FROM ip_brands")

def add_ip_brand(name: str) -> bool:
    if not name: return False
    name_clean = name.strip().upper()
    with _db_lock:
        with _get_connection() as conn:
            try:
                conn.execute("INSERT INTO ip_brands (name) VALUES (?)", (name_clean,))
                return True
            except sqlite3.IntegrityError:
                return False

def is_ip_brand(name: str) -> bool:
    if not name: return False
    name_clean = name.strip().upper()
    with _db_lock:
        with _get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM ip_brands")
            for row in cursor.fetchall():
                db_name = row[0]
                # Проверяем точное совпадение
                if db_name.upper() == name_clean.upper():
                    return True
    return False

def get_ip_brands_count() -> int:
    with _db_lock:
        with _get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM ip_brands")
            return cursor.fetchone()[0]

def get_all_ip_brands() -> list[str]:
    """Возвращает список всех брендов из таблицы ТРОИС."""
    with _db_lock:
        with _get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM ip_brands ORDER BY name")
            return [r[0] for r in cursor.fetchall()]

def add_docs_category(name: str) -> bool:
    if not name: return False
    name_clean = name.strip()
    with _db_lock:
        with _get_connection() as conn:
            try:
                conn.execute("INSERT INTO docs_categories (name) VALUES (?)", (name_clean,))
                return True
            except sqlite3.IntegrityError:
                return False

def remove_docs_category(name: str) -> bool:
    if not name: return False
    name_clean = name.strip()
    with _db_lock:
        with _get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM docs_categories WHERE name = ? COLLATE NOCASE", (name_clean,))
            return cursor.rowcount > 0

def get_docs_categories(limit: int = 0, offset: int = 0) -> list[str]:
    with _db_lock:
        with _get_connection() as conn:
            cursor = conn.cursor()
            query = "SELECT name FROM docs_categories ORDER BY name"
            if limit > 0:
                query += f" LIMIT {limit} OFFSET {offset}"
            cursor.execute(query)
            return [r[0] for r in cursor.fetchall()]

def get_docs_categories_count() -> int:
    with _db_lock:
        with _get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM docs_categories")
            return cursor.fetchone()[0]

def is_docs_category(category: str) -> bool:
    if not category: return False
    cat_clean = category.strip().lower()
    with _db_lock:
        with _get_connection() as conn:
            cursor = conn.cursor()
            # Для ускорения проверки можно было бы сделать SELECT с WHERE, но категорий мало, линейный поиск достаточно быстр
            cursor.execute("SELECT name FROM docs_categories")
            for row in cursor.fetchall():
                if row[0].strip().lower() == cat_clean:
                    return True
    return False

def add_labeling_category(name: str) -> bool:
    if not name: return False
    name_clean = name.strip()
    with _db_lock:
        with _get_connection() as conn:
            try:
                conn.execute("INSERT INTO labeling_categories (name) VALUES (?)", (name_clean,))
                return True
            except sqlite3.IntegrityError:
                return False

def remove_labeling_category(name: str) -> bool:
    if not name: return False
    name_clean = name.strip()
    with _db_lock:
        with _get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM labeling_categories WHERE name = ? COLLATE NOCASE", (name_clean,))
            return cursor.rowcount > 0

def get_labeling_categories(limit: int = 0, offset: int = 0) -> list[str]:
    with _db_lock:
        with _get_connection() as conn:
            cursor = conn.cursor()
            query = "SELECT name FROM labeling_categories ORDER BY name"
            if limit > 0:
                query += f" LIMIT {limit} OFFSET {offset}"
            cursor.execute(query)
            return [r[0] for r in cursor.fetchall()]

def get_labeling_categories_count() -> int:
    with _db_lock:
        with _get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM labeling_categories")
            return cursor.fetchone()[0]

def is_labeling_category(category: str) -> bool:
    if not category: return False
    cat_clean = category.strip().lower()
    with _db_lock:
        with _get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM labeling_categories")
            for row in cursor.fetchall():
                if row[0].strip().lower() == cat_clean:
                    return True
    return False

def add_to_history(user_id: int, pn: str):
    if not pn:
        return
    pn_clean = pn.strip().upper()
    with _db_lock:
        with _get_connection() as conn:
            # Не добавляем дубликат, если это был последний запрос этого пользователя
            cursor = conn.cursor()
            cursor.execute("SELECT pn FROM user_history WHERE user_id = ? ORDER BY timestamp DESC LIMIT 1", (user_id,))
            last_row = cursor.fetchone()
            if last_row and last_row[0] == pn_clean:
                return
                
            conn.execute(
                "INSERT INTO user_history (user_id, pn) VALUES (?, ?)",
                (user_id, pn_clean)
            )

def get_user_history(user_id: int, limit: int = 50) -> list[str]:
    with _db_lock:
        with _get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT DISTINCT pn FROM user_history WHERE user_id = ? ORDER BY timestamp DESC LIMIT ?", 
                (user_id, limit)
            )
            rows = cursor.fetchall()
            return [r[0] for r in rows]

def clear_user_history(user_id: int):
    with _db_lock:
        with _get_connection() as conn:
            conn.execute("DELETE FROM user_history WHERE user_id = ?", (user_id,))

def get_cached_part(pn: str) -> Optional[Dict[str, Any]]:
    """Возвращает кэшированные данные детали по парт-номеру, если они есть."""
    if not pn:
        return None
    # Нормализуем для БД: убираем пробелы и плюсы, как и перед API
    pn_clean = pn.strip().lower().replace(" ", "").replace("+", "")
    
    with _db_lock:
        with _get_connection() as conn:
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
    """Сохраняет данные детали в кэш."""
    if not pn or not data:
        return
    # Нормализуем для БД: убираем пробелы и плюсы
    pn_clean = pn.strip().lower().replace(" ", "").replace("+", "")
    
    with _db_lock:
        with _get_connection() as conn:
            conn.execute(
                "INSERT OR REPLACE INTO part_cache (pn, data) VALUES (?, ?)",
                (pn_clean, json.dumps(data, ensure_ascii=False))
            )

def get_cache_stats() -> int:
    """Возвращает количество записей в кэше деталей."""
    with _db_lock:
        with _get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM part_cache")
            return cursor.fetchone()[0]

def delete_cached_part(pn: str):
    """Удаляет деталь из кэша (для принудительного обновления)."""
    if not pn:
        return
    # Нормализуем для БД: убираем пробелы и плюсы
    pn_clean = pn.strip().lower().replace(" ", "").replace("+", "")
    with _db_lock:
        with _get_connection() as conn:
            conn.execute("DELETE FROM part_cache WHERE pn = ?", (pn_clean,))

def get_all_cached_pns() -> list[str]:
    """Возвращает список всех уникальных партномеров из кэша."""
    with _db_lock:
        with _get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT pn FROM part_cache")
            return [r[0] for r in cursor.fetchall()]
