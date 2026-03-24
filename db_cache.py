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
    return conn

def get_cached_part(pn: str) -> Optional[Dict[str, Any]]:
    """Возвращает кэшированные данные детали по парт-номеру, если они есть."""
    if not pn:
        return None
    pn_clean = pn.strip().lower()
    
    with _db_lock:
        with _get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT data FROM part_cache WHERE pn = ?", (pn_clean,))
            row = cursor.fetchone()
            if row:
                try:
                    return json.loads(row[0])
                except Exception:
                    return None
    return None

def save_cached_part(pn: str, data: Dict[str, Any]):
    """Сохраняет данные детали в кэш."""
    if not pn or not data:
        return
    pn_clean = pn.strip().lower()
    
    with _db_lock:
        with _get_connection() as conn:
            conn.execute(
                "INSERT OR REPLACE INTO part_cache (pn, data) VALUES (?, ?)",
                (pn_clean, json.dumps(data, ensure_ascii=False))
            )
