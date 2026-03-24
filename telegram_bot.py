import os
import sys
import asyncio
import logging
import tempfile
import math
from pathlib import Path

from dotenv import load_dotenv, find_dotenv
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import CommandStart, Command
from aiogram.types import Message, FSInputFile, InlineKeyboardMarkup, InlineKeyboardButton, CallbackQuery, LinkPreviewOptions, ReplyKeyboardRemove, BotCommand
from aiogram.enums import ParseMode
from aiogram.client.default import DefaultBotProperties
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from openpyxl import load_workbook, Workbook

class UpdateState(StatesGroup):
    waiting_for_data = State()

# Импортируем готовые функции из нашего проекта
from mouser_cli import _get_api_key, write_xlsx
from mouser_menu import _lookup_one, _split_pns
from db_cache import add_to_history, get_user_history, clear_user_history, get_cache_stats, delete_cached_part

load_dotenv(find_dotenv(), override=False)

TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
if not TOKEN:
    print("ВНИМАНИЕ: Переменная TELEGRAM_BOT_TOKEN не задана в .env файле.")

# Опциональная настройка для ограничения работы бота конкретным топиком в группе
ALLOWED_THREAD_ID = os.getenv("ALLOWED_THREAD_ID")
if ALLOWED_THREAD_ID:
    try:
        ALLOWED_THREAD_ID = int(ALLOWED_THREAD_ID)
    except ValueError:
        print("ВНИМАНИЕ: ALLOWED_THREAD_ID должен быть числом.")
        ALLOWED_THREAD_ID = None

dp = Dispatcher()

def get_menu_text():
    return (
        "🤖 <b>Как работает парсер Mouser:</b>\n\n"
        "Отправьте мне:\n"
        "1️⃣ <b>Текст</b> с парт-номерами (через пробел или запятую) - я верну информацию прямо сюда.\n"
        "2️⃣ <b>Excel-файл</b> (.xlsx) - я найду в нем столбец 'Part no.' (или 'PART NUMBER'), "
        "соберу все номера и отправлю обратно заполненный Excel-файл со всеми данными, включая перевод."
    )

def get_inline_menu_keyboard():
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="🔄 Актуализация", callback_data="menu_update")],
            [InlineKeyboardButton(text="🕒 История запросов", callback_data="menu_history")],
            [InlineKeyboardButton(text="📊 Статистика БД", callback_data="menu_stats")],
            [InlineKeyboardButton(text="🗑 Очистить историю", callback_data="menu_clear_history")]
        ]
    )

# Храним последние запрошенные партномера для генерации Excel по кнопке
_user_last_pns = {}

def process_pns_sync(pns, api_key, force_update=False):
    """Синхронная функция обработки списка партномеров для запуска в отдельном потоке"""
    results = []
    for pn in pns:
        row = _lookup_one(pn, api_key, retries=3, timeout=20, save_raw_dir=None, force_update=force_update)
        if row:
            results.append(row)
    return results

def extract_pns_from_excel_sync(filepath):
    """Извлекает партномера из Excel файла"""
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
        return None, "NO_COLUMN"

    pns = []
    for row_idx in range(start_row, ws.max_row + 1):
        val = ws.cell(row=row_idx, column=part_col_idx).value
        if val is None or str(val).strip() == "":
            break
        pns.append(str(val).strip())

    if not pns:
        return None, "Список парт-номеров пуст."

    return pns, None

def create_template_excel():
    """Создает пустой шаблон Excel"""
    wb = Workbook()
    ws = wb.active
    ws.append(["Part no."])
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    wb.save(tmp.name)
    return tmp.name

def get_progress_bar(current, total, length=10):
    percent = current / total
    filled = int(length * percent)
    bar = "⬛" * filled + "⬜" * (length - filled)
    return f"{bar} ({int(percent * 100)}%)"

@dp.message(CommandStart())
async def command_start_handler(message: Message) -> None:
    await message.answer(
        "Привет! Я готов к работе. Нажмите кнопку «Меню» (слева от поля ввода текста), чтобы узнать, что я умею.",
        reply_markup=ReplyKeyboardRemove()
    )

@dp.message(Command("menu"))
async def handle_menu_command(message: Message) -> None:
    await message.answer(
        get_menu_text(),
        reply_markup=get_inline_menu_keyboard()
    )

@dp.callback_query(F.data == "menu_update")
async def callback_menu_update(callback: CallbackQuery, state: FSMContext):
    await state.set_state(UpdateState.waiting_for_data)
    await callback.message.answer(
        "🔄 <b>Режим актуализации</b>\n\n"
        "Отправьте мне парт-номера текстом или Excel-файл. "
        "Я проигнорирую локальную базу и скачаю самые свежие данные с сайта Mouser."
    )
    await callback.answer()

@dp.callback_query(F.data == "menu_history")
async def callback_menu_history(callback: CallbackQuery):
    await callback.answer()
    history = get_user_history(callback.from_user.id)
    if not history:
        await callback.message.answer("Ваша история запросов пуста.")
        return
        
    text = "🕒 <b>Ваша история последних запросов (до 50 шт):</b>\n\n"
    text += "\n".join([f"• <code>{pn}</code>" for pn in history])
    
    # Разделяем длинный текст на части, если он больше 4000 символов
    if len(text) > 4000:
        for i in range(0, len(text), 4000):
            await callback.message.answer(text[i:i+4000])
    else:
        await callback.message.answer(text)

@dp.callback_query(F.data == "menu_stats")
async def callback_menu_stats(callback: CallbackQuery):
    count = get_cache_stats()
    await callback.message.answer(f"📊 <b>Статистика базы данных:</b>\n\nВ локальном кэше сохранено деталей: <b>{count}</b>")
    await callback.answer()

@dp.callback_query(F.data == "menu_clear_history")
async def callback_menu_clear_history(callback: CallbackQuery):
    clear_user_history(callback.from_user.id)
    await callback.answer("✅ Ваша история запросов очищена.", show_alert=True)

def format_text_result(results):
    lines = []
    for r in results:
        req_pn = r.get("Запрошенный партномер", "-")
        mouser_pn = r.get("Партномер Mouser", "-")
        mfr = r.get("Производитель", "-")
        cat = r.get("Категория", "-")
        desc = r.get("Описание", "-")
        img = r.get("Ссылка на фото")
        ds = r.get("Даташит")
        cache_date = r.get("Дата кэширования")

        lines.append(f"🔍 <b>{req_pn}</b>")
        lines.append(f"Mouser PN: <code>{mouser_pn}</code>")
        lines.append(f"Производитель: {mfr}")
        lines.append(f"Категория: {cat}")
        lines.append(f"Описание: {desc}")
        if cache_date:
            lines.append(f"<i>Взято из БД: {cache_date}</i>")
        
        links = []
        if ds and ds != "-":
            links.append(f'<a href="{ds}">📄 Даташит</a>')
        if img and img != "-":
            links.append(f'<a href="{img}">🖼 Фото</a>')
            
        if links:
            lines.append(" | ".join(links))
            
        lines.append("-" * 25)
    return "\n".join(lines)


@dp.message(F.text & ~F.text.startswith('/'))
async def handle_text(message: Message, bot: Bot, state: FSMContext) -> None:
    # Проверяем состояние
    current_state = await state.get_state()
    force_update = current_state == UpdateState.waiting_for_data.state
    if force_update:
        await state.clear()

    # Для групп проверяем, упомянут ли бот или это ответ боту
    if message.chat.type != "private":
        # Проверяем топик, если он задан
        if ALLOWED_THREAD_ID is not None:
            # message_thread_id может быть None в главном чате (General)
            if message.message_thread_id != ALLOWED_THREAD_ID:
                return

        bot_user = await bot.get_me()
        is_reply_to_bot = message.reply_to_message and message.reply_to_message.from_user.id == bot_user.id
        is_mentioned = f"@{bot_user.username}" in message.text
        if not (is_reply_to_bot or is_mentioned):
            return

    try:
        api_key = _get_api_key()
    except SystemExit:
        await message.answer("Ошибка: MOUSER_API_KEY не настроен на сервере.")
        return

    # Очищаем текст от упоминания бота, если оно есть
    bot_user = await bot.get_me()
    text = message.text.replace(f"@{bot_user.username}", "").strip()
    pns = _split_pns(text)
    if not pns:
        return

    update_text = " (принудительное обновление)..." if force_update else "..."
    msg = await message.answer(f"⏳ Ищу информацию по {len(pns)} деталям{update_text}")
    
    # Сохраняем в историю
    for pn in pns:
        add_to_history(message.from_user.id, pn)
    
    # Сохраняем последние PNs для выгрузки
    _user_last_pns[message.from_user.id] = pns

    # Запускаем синхронный парсинг
    results = await asyncio.to_thread(process_pns_sync, pns, api_key, force_update)
    
    if not results:
        await msg.edit_text("⚠️ Ничего не найдено или произошла ошибка.")
        return

    out_text = format_text_result(results)
    
    # Клавиатура
    keyboard = []
    if len(pns) == 1:
        # Если одна деталь, даем возможность очистить кэш
        keyboard.append([InlineKeyboardButton(text="🔄 Очистить кэш этой детали", callback_data=f"clear_{pns[0][:40]}")])
    
    keyboard.append([InlineKeyboardButton(text="📄 Выгрузить это в Excel", callback_data="export_excel")])
    markup = InlineKeyboardMarkup(inline_keyboard=keyboard)

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
        await msg.edit_text(
            out_text, 
            reply_markup=markup, 
            link_preview_options=LinkPreviewOptions(is_disabled=True)
        )

@dp.callback_query(F.data == "export_excel")
async def callback_export_excel(callback: CallbackQuery):
    pns = _user_last_pns.get(callback.from_user.id)
    if not pns:
        await callback.answer("Данные устарели, отправьте парт-номера заново.", show_alert=True)
        return

    await callback.message.edit_reply_markup(reply_markup=None)
    msg = await callback.message.answer("⏳ Формирую Excel-файл...")
    
    api_key = _get_api_key()
    results = await asyncio.to_thread(process_pns_sync, pns, api_key)
    
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_name = tmp.name
    await asyncio.to_thread(write_xlsx, results, tmp_name)
    
    file = FSInputFile(tmp_name, filename="mouser_results.xlsx")
    await callback.message.answer_document(file)
    os.remove(tmp_name)
    await msg.delete()
    await callback.answer()

@dp.callback_query(F.data.startswith("clear_"))
async def callback_clear_cache(callback: CallbackQuery):
    pn = callback.data.split("_", 1)[1]
    delete_cached_part(pn)
    await callback.answer(f"Кэш для {pn} очищен! Можете искать заново.", show_alert=True)


@dp.message(F.document)
async def handle_document(message: Message, bot: Bot, state: FSMContext) -> None:
    # Проверяем состояние
    current_state = await state.get_state()
    force_update = current_state == UpdateState.waiting_for_data.state
    if force_update:
        await state.clear()

    # Для групп проверяем, упомянут ли бот в подписи к файлу или файл отправлен в ответ боту
    if message.chat.type != "private":
        bot_user = await bot.get_me()
        is_reply_to_bot = message.reply_to_message and message.reply_to_message.from_user.id == bot_user.id
        is_mentioned = message.caption and f"@{bot_user.username}" in message.caption
        if not (is_reply_to_bot or is_mentioned):
            return

    doc = message.document
    if not doc.file_name.lower().endswith(('.xlsx', '.xls')):
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
    
    await msg.edit_text("⏳ Читаю Excel-файл...")
    pns, err = await asyncio.to_thread(extract_pns_from_excel_sync, input_path)
    os.remove(input_path)
    
    if err == "NO_COLUMN":
        template_path = await asyncio.to_thread(create_template_excel)
        file = FSInputFile(template_path, filename="Template.xlsx")
        await msg.delete()
        await message.answer_document(
            file, 
            caption="❌ Я не нашел нужную колонку ('Part no.').\n\nПожалуйста, используйте этот шаблон, заполните первую колонку и отправьте мне обратно."
        )
        os.remove(template_path)
        return
    elif err:
        await msg.edit_text(f"❌ {err}")
        return

    if not pns:
        await msg.edit_text("⚠️ Нет результатов для сохранения.")
        return

    # Сохраняем в историю
    for pn in pns:
        add_to_history(message.from_user.id, pn)

    total_pns = len(pns)
    update_text = " (принудительное обновление)" if force_update else ""
    await msg.edit_text(f"⏳ Начинаю обработку {total_pns} деталей{update_text}...\n{get_progress_bar(0, total_pns)}")

    all_results = []
    processed = 0
    last_update_time = asyncio.get_event_loop().time()
    
    # Обрабатываем по одному, чтобы прогресс бар был плавным
    for pn in pns:
        # Вызываем _lookup_one в отдельном потоке для каждого партномера
        row = await asyncio.to_thread(_lookup_one, pn, api_key, 3, 20, None, force_update)
        if row:
            all_results.append(row)
            
        processed += 1
        
        # Обновляем сообщение не чаще 1 раза в секунду, чтобы не ловить лимиты Telegram (Flood Control)
        current_time = asyncio.get_event_loop().time()
        if current_time - last_update_time > 1.0 or processed == total_pns:
            progress_text = (
                f"⏳ Обработка файла (учитываем лимиты API)...\n\n"
                f"Обработано: {processed} / {total_pns}\n"
                f"{get_progress_bar(processed, total_pns)}\n\n"
            )
            try:
                await msg.edit_text(progress_text)
                last_update_time = current_time
            except Exception as e:
                # Игнорируем ошибки MessageNotModified
                pass

    await msg.edit_text("⏳ Формирую итоговый файл...")
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_out:
        out_path = tmp_out.name
        
    await asyncio.to_thread(write_xlsx, all_results, out_path)
    
    file = FSInputFile(out_path, filename=f"mouser_result_{doc.file_name}")
    await message.answer_document(file, caption="✅ Готово! Все детали обработаны.")
    os.remove(out_path)
    await msg.delete()

async def main() -> None:
    if not TOKEN:
        print("ОШИБКА: Задайте TELEGRAM_BOT_TOKEN перед запуском бота.")
        sys.exit(1)
        
    bot = Bot(token=TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    
    # Настраиваем кнопку меню слева от поля ввода
    await bot.set_my_commands([
        BotCommand(command="start", description="Запустить бота"),
        BotCommand(command="menu", description="Главное меню"),
        BotCommand(command="history", description="История запросов"),
        BotCommand(command="stats", description="Статистика БД")
    ])
    
    print("Бот запущен. Ожидание сообщений...")
    await dp.start_polling(bot)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, stream=sys.stdout)
    asyncio.run(main())