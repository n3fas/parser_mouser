import sys
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime

class EmojiFormatter(logging.Formatter):
    FORMATS = {
        logging.DEBUG: "🐛 %(asctime)s [%(levelname)s] %(message)s",
        logging.INFO: "ℹ️ %(asctime)s [%(levelname)s] %(message)s",
        logging.WARNING: "⚠️ %(asctime)s [%(levelname)s] %(message)s",
        logging.ERROR: "❌ %(asctime)s [%(levelname)s] %(message)s",
        logging.CRITICAL: "🚨 %(asctime)s [%(levelname)s] %(message)s"
    }

    def format(self, record):
        log_fmt = self.FORMATS.get(record.levelno, "%(asctime)s [%(levelname)s] %(message)s")
        formatter = logging.Formatter(log_fmt, datefmt="%Y-%m-%d %H:%M:%S")
        return formatter.format(record)

def get_base_logger(name="mouser_bot", log_file="bot.log"):
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger
        
    logger.setLevel(logging.INFO)

    file_handler = RotatingFileHandler(log_file, maxBytes=5*1024*1024, backupCount=3, encoding='utf-8')
    file_handler.setFormatter(EmojiFormatter())
    logger.addHandler(file_handler)

    # In .pyw files, sys.stdout and sys.stderr are None
    if sys.stdout is not None:
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(EmojiFormatter())
        logger.addHandler(console_handler)

    return logger

class StreamToLogger:
    def __init__(self, logger, log_level=logging.INFO):
        self.logger = logger
        self.log_level = log_level

    def write(self, buf):
        for line in buf.rstrip().splitlines():
            line = line.strip()
            if line:
                lvl = self.log_level
                if "error" in line.lower() or "ошибка" in line.lower() or "exception" in line.lower() or "traceback" in line.lower():
                    lvl = logging.ERROR
                elif "warning" in line.lower() or "внимание" in line.lower():
                    lvl = logging.WARNING
                self.logger.log(lvl, line)

    def flush(self):
        pass

def setup_bot_logger(log_file="bot.log"):
    logger = get_base_logger("mouser_bot", log_file)
    
    # Configure root logger so that standard exceptions and other libs (aiogram) use our formatter
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)
        
    for handler in logger.handlers:
        root_logger.addHandler(handler)

    # Redirect sys.stdout and sys.stderr to catch prints from other files like mouser_cli.py
    sys.stdout = StreamToLogger(logger, logging.INFO)
    # Always redirect stderr even if running in console so we get errors with nice format
    sys.stderr = StreamToLogger(logger, logging.ERROR)
    
    return logger
