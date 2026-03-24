import os
import sys
import asyncio
import logging
import tempfile
from pathlib import Path

from dotenv import load_dotenv, find_dotenv
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import CommandStart
from aiogram.types import Message, FSInputFile
from aiogram.enums import ParseMode
from aiogram.client.default import DefaultBotProperties
from openpyxl import load_workbook

# Импортируем готовые функции из нашего проекта
from mouser_cli import _get_api_key, write_xlsx
from mouser_menu import _lookup_one, _split_pns

load_dotenv(find_dotenv(), override=False)

TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
if not TOKEN:
    print("ВНИМАНИЕ: Переменная TELEGRAM_BOT_TOKEN не задана в .env файле.")
    # Бот упадет при запуске без токена, но оставим это для логов.

dp = Dispatcher()

def process_pns_sync(pns, api_key):
    """Синхронная функция обработки списка партномеров для запуска в отдельном потоке"""
    results = []
    for pn in pns:
        row = _lookup_one(pn, api_key, retries=3, timeout=20, save_raw_dir=None)
        if row:
            results.append(row)
    return results

def process_excel_sync(filepath, api_key):
    """Синхронная функция обработки Excel-файла для запуска в отдельном потоке"""
    try:
        wb = load_workbook(filepath, data_only=True)
        ws = wb.active
    except Exception as e:
        return None, f"Ошибка чтения Excel: {e}"

    part_col_idx = None
    start_row = None

    for row_idx, row in enumerate(ws.iter_rows(max_row=50), start=1):
        for col_idx, cell in enumerate(row, start=1):
            if cell.value and isinstance(cell.value, str):
                val_clean = cell.value.strip().lower().replace(" ", "").replace(".", "")
                if val_clean in ("partno", "partnumber"):
                    part_col_idx = col_idx
                    start_row = row_idx + 1
                    break
        if part_col_idx:
            break

    if not part_col_idx:
        return None, "Столбец с партномером ('Part no.', 'PART NUMBER' и т.д.) не найден (проверены первые 50 строк)."

    pns = []
    for row_idx in range(start_row, ws.max_row + 1):
        val = ws.cell(row=row_idx, column=part_col_idx).value
        if val is None or str(val).strip() == "":
            break
        pns.append(str(val).strip())

    if not pns:
        return None, "Список парт-номеров пуст."

    # Разбиваем на пачки по 50
    chunks = [pns[i:i + 50] for i in range(0, len(pns), 50)]
    
    all_results = []
    for chunk in chunks:
        for pn in chunk:
            row = _lookup_one(pn, api_key, retries=3, timeout=20, save_raw_dir=None)
            if row:
                all_results.append(row)

    if not all_results:
        return None, "Нет успешных результатов."

    return all_results, None


@dp.message(CommandStart())
async def command_start_handler(message: Message) -> None:
    text = (
        "Привет! Я бот для парсинга каталога Mouser.\n\n"
        "Отправь мне:\n"
        "1. <b>Текст</b> с парт-номерами (через пробел или запятую) - я верну информацию прямо сюда.\n"
        "2. <b>Excel-файл</b> (.xlsx) - я найду в нем столбец 'Part no.' (или 'PART NUMBER'), "
        "соберу все номера и отправлю обратно заполненный Excel-файл со всеми данными, включая перевод."
    )
    await message.answer(text)


def format_text_result(results):
    lines = []
    for r in results:
        req_pn = r.get("Запрошенный партномер", "-")
        mouser_pn = r.get("Партномер Mouser", "-")
        mfr = r.get("Производитель", "-")
        cat = r.get("Категория", "-")
        desc = r.get("Описание", "-")
        price = r.get("Цена за единицу", "-")
        stock = r.get("Доступно", "-")
        
        lines.append(f"🔍 <b>{req_pn}</b>")
        lines.append(f"Mouser PN: {mouser_pn}")
        lines.append(f"Производитель: {mfr}")
        lines.append(f"Категория: {cat}")
        lines.append(f"Наличие: {stock} шт. | Цена: {price}")
        lines.append(f"Описание: {desc}")
        lines.append("-" * 25)
    return "\n".join(lines)


@dp.message(F.text)
async def handle_text(message: Message) -> None:
    try:
        api_key = _get_api_key()
    except SystemExit:
        await message.answer("Ошибка: MOUSER_API_KEY не настроен на сервере.")
        return

    pns = _split_pns(message.text)
    if not pns:
        await message.answer("Не нашел парт-номеров в сообщении.")
        return

    msg = await message.answer(f"⏳ Ищу информацию по {len(pns)} деталям...")
    
    # Запускаем синхронный парсинг в отдельном потоке, чтобы не блокировать бота
    results = await asyncio.to_thread(process_pns_sync, pns, api_key)
    
    if not results:
        await msg.edit_text("⚠️ Ничего не найдено или произошла ошибка.")
        return

    out_text = format_text_result(results)
    
    # Если текст слишком длинный для одного сообщения Telegram, скинем файлом
    if len(out_text) > 4000:
        await msg.edit_text("✅ Готово! Результат слишком большой, формирую Excel-файл...")
        
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_name = tmp.name
            
        await asyncio.to_thread(write_xlsx, results, tmp_name)
        
        file = FSInputFile(tmp_name, filename="mouser_results.xlsx")
        await message.answer_document(file)
        os.remove(tmp_name)
        await msg.delete()
    else:
        await msg.edit_text(out_text)


@dp.message(F.document)
async def handle_document(message: Message, bot: Bot) -> None:
    doc = message.document
    if not doc.file_name.lower().endswith('.xlsx'):
        await message.answer("Пожалуйста, отправьте файл в формате .xlsx")
        return

    try:
        api_key = _get_api_key()
    except SystemExit:
        await message.answer("Ошибка: MOUSER_API_KEY не настроен на сервере.")
        return

    msg = await message.answer("⏳ Скачиваю файл...")
    
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        input_path = tmp.name
        
    await bot.download(doc, destination=input_path)
    
    await msg.edit_text("⏳ Обрабатываю Excel-файл (учитываем лимиты 30 зап/мин)...")
    
    results, err = await asyncio.to_thread(process_excel_sync, input_path, api_key)
    os.remove(input_path)
    
    if err:
        await msg.edit_text(f"❌ {err}")
        return
        
    if not results:
        await msg.edit_text("⚠️ Нет результатов для сохранения.")
        return

    await msg.edit_text("⏳ Формирую итоговый файл...")
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_out:
        out_path = tmp_out.name
        
    await asyncio.to_thread(write_xlsx, results, out_path)
    
    file = FSInputFile(out_path, filename=f"mouser_result_{doc.file_name}")
    await message.answer_document(file, caption="✅ Готово!")
    os.remove(out_path)
    await msg.delete()


async def main() -> None:
    if not TOKEN:
        print("ОШИБКА: Задайте TELEGRAM_BOT_TOKEN перед запуском бота.")
        sys.exit(1)
        
    bot = Bot(token=TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    print("Бот запущен. Ожидание сообщений...")
    await dp.start_polling(bot)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, stream=sys.stdout)
    asyncio.run(main())