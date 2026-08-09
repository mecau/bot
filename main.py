import asyncio
import logging
import base64
import io
import json
import os
import re
import urllib.parse
import aiohttp
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import CommandStart, Command
from aiogram.types import (
    InlineKeyboardMarkup, InlineKeyboardButton,
    ReplyKeyboardMarkup, KeyboardButton, Message, BotCommand, BufferedInputFile
)
from openai import AsyncOpenAI
from docx import Document
from docx.shared import Pt, Inches, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from config import (
    BOT_TOKEN, GROQ_API_KEY, 
    CHANNEL_1_USERNAME, CHANNEL_1_URL,
    CHANNEL_2_USERNAME, CHANNEL_2_URL,
    TEXT_MODEL, VISION_MODEL, PROMPTS, AD_FOOTER
)

MY_ADMIN_ID = 1184589026

logging.basicConfig(level=logging.INFO)
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

groq_client = AsyncOpenAI(
    api_key=GROQ_API_KEY,
    base_url="https://api.groq.com/openai/v1"
)

USERS_FILE = "users.json"
FAV_FILE = "favorites.json"

all_users_cache = set()

def load_user_ids() -> set:
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return set(data)
        except Exception:
            pass
    return set()

def save_user_id(user_id: int):
    global all_users_cache
    all_users_cache.add(user_id)
    try:
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump(list(all_users_cache), f, ensure_ascii=False)
    except Exception as e:
        logging.error(f"Не удалось сохранить пользователя: {e}")

all_users_cache = load_user_ids()

def load_favorites() -> dict:
    if os.path.exists(FAV_FILE):
        try:
            with open(FAV_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_favorites(favs: dict):
    with open(FAV_FILE, "w", encoding="utf-8") as f:
        json.dump(favs, f, ensure_ascii=False, indent=2)

def add_to_favorites(user_id: int, text: str):
    favs = load_favorites()
    uid_str = str(user_id)
    if uid_str not in favs:
        favs[uid_str] = []
    if text not in favs[uid_str]:
        favs[uid_str].append(text)
        save_favorites(favs)

users_db = {}
broadcast_states = set()
ppt_states = set()
excel_states = set()
user_ppt_images = {}

def get_user_data(user_id: int):
    if user_id not in users_db:
        users_db[user_id] = {
            "mode": "ai",
            "history": [],
            "last_output": "Здесь пока нет ответов."
        }
    return users_db[user_id]

async def check_subscription(user_id: int) -> bool:
    if user_id == MY_ADMIN_ID:
        return True
    channels = [CHANNEL_1_USERNAME, CHANNEL_2_USERNAME]
    for ch in channels:
        try:
            member = await bot.get_chat_member(chat_id=f"@{ch}", user_id=user_id)
            if member.status not in ["creator", "administrator", "member"]:
                return False
        except Exception as e:
            logging.error(f"Ошибка проверки подписки на канал @{ch}: {e}")
            return False
    return True

def get_sub_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📢 Подписаться на канал 1", url=CHANNEL_1_URL)],
        [InlineKeyboardButton(text="📢 Подписаться на канал 2", url=CHANNEL_2_URL)],
        [InlineKeyboardButton(text="✅ Проверить подписку", callback_data="check_sub")]
    ])

def get_main_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📄 Скачать ответ в Word"), KeyboardButton(text="📑 Создать титульник ГОСТ")],
            [KeyboardButton(text="📈 Создать презентацию"), KeyboardButton(text="📊 Создать таблицу Excel")],
            [KeyboardButton(text="⭐ Избранное"), KeyboardButton(text="🔄 Сменить режим")],
            [KeyboardButton(text="ℹ️ О MecauAI")]
        ],
        resize_keyboard=True
    )

def get_admin_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📄 Скачать ответ в Word"), KeyboardButton(text="📑 Создать титульник ГОСТ")],
            [KeyboardButton(text="📈 Создать презентацию"), KeyboardButton(text="📊 Создать таблицу Excel")],
            [KeyboardButton(text="⭐ Избранное"), KeyboardButton(text="🔄 Сменить режим")],
            [KeyboardButton(text="ℹ️ О MecauAI"), KeyboardButton(text="📊 Статистика бота")],
            [KeyboardButton(text="📢 Сделать рассылку")]
        ],
        resize_keyboard=True
    )

def keyboard_for(user_id: int):
    return get_admin_keyboard() if user_id == MY_ADMIN_ID else get_main_keyboard()

def get_answer_inline_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="💡 Объяснить проще", callback_data="btn_simplify"),
            InlineKeyboardButton(text="⭐ В избранное", callback_data="btn_save_fav")
        ]
    ])

def clean_text_for_html(text: str) -> str:
    text = re.sub(r'<think>.*?</think>', '', text, flags=re.DOTALL)
    return text.strip()

@dp.message(CommandStart())
async def cmd_start(message: Message):
    user_id = message.from_user.id
    save_user_id(user_id)

    if not await check_subscription(user_id):
        await message.answer(
            f"🔒 Доступ заблокирован!\n\n"
            f"Чтобы пользоваться MecauAI, подпишись на оба наших канала:\n"
            f" 👉 {CHANNEL_1_URL}\n"
            f" 👉 {CHANNEL_2_URL}\n\n"
            f"После подписки нажми кнопку ниже 👇",
            reply_markup=get_sub_keyboard()
        )
        return

    start_text = (
        f"Привет, {message.from_user.first_name}! Ты активировал MecauAI 🚀\n\n"
        "Я твой карманный помощник для учебы и разработки. Отправляй любые вопросы, задачи, картинки, код или создавай презентации с таблицами!"
    )
    await message.answer(start_text, reply_markup=keyboard_for(user_id))

@dp.callback_query(F.data == "check_sub")
async def cb_check_sub(callback: types.CallbackQuery):
    user_id = callback.from_user.id
    if await check_subscription(user_id):
        save_user_id(user_id)
        try:
            await callback.message.delete()
        except Exception:
            pass
        await callback.message.answer(
            "🎉 Подписка на оба канала подтверждена! Добро пожаловать в MecauAI 🚀",
            reply_markup=keyboard_for(user_id)
        )
    else:
        await callback.answer("❌ Ты подписался еще не на все каналы!", show_alert=True)

@dp.message(F.text == "ℹ️ О MecauAI")
@dp.message(Command("about"))
async def cmd_about(message: Message):
    if not await check_subscription(message.from_user.id):
        await message.answer(f"🔒 Сначала подпишись на каналы:\n1️⃣ {CHANNEL_1_URL}\n2️⃣ {CHANNEL_2_URL}", reply_markup=get_sub_keyboard())
        return
    about_text = (
        "🧠 Возможности MecauAI:\n\n"
        "• 💻 Режим «ИИ-Программист»: Написание чистого кода с копированием в 1 клик.\n"
        "• 🧠 Академический ассистент и Лучший друг.\n"
        "• 📄 Экспорт в .docx и Титульники по ГОСТу.\n"
        "• 📈 Генерация презентаций (.pptx) с твоими картинками.\n"
        "• 📊 Автоматические таблицы Excel (.xlsx) и Избранное."
    )
    await message.answer(about_text, reply_markup=keyboard_for(message.from_user.id))

@dp.message(F.text == "🔄 Сменить режим")
@dp.message(Command("mode"))
async def cmd_mode(message: Message):
    if not await check_subscription(message.from_user.id):
        await message.answer(f"🔒 Сначала подпишись на каналы:\n1️⃣ {CHANNEL_1_URL}\n2️⃣ {CHANNEL_2_URL}", reply_markup=get_sub_keyboard())
        return
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="💻 ИИ-Программист",       callback_data="set_mode_coder")],
        [InlineKeyboardButton(text="🧠 Академический ассистент", callback_data="set_mode_ai")],
        [InlineKeyboardButton(text="🫂 Лучший друг",             callback_data="set_mode_friend")]
    ])
    await message.answer("Выбери режим работы бота:", reply_markup=kb)

@dp.callback_query(F.data.startswith("set_mode_"))
async def cb_set_mode(callback: types.CallbackQuery):
    mode = callback.data.split("_")[-1]
    user_data = get_user_data(callback.from_user.id)
    user_data["mode"] = mode
    
    names = {
        "coder": "ИИ-Программист 💻",
        "ai": "Академический ассистент 🧠",
        "friend": "Лучший друг 🫂"
    }
    await callback.message.edit_text(f"Режим переключен на: {names.get(mode, 'Стандартный')}")
    await callback.answer()

@dp.message(F.text == "📊 Статистика бота")
async def cmd_stats(message: Message):
    if message.from_user.id != MY_ADMIN_ID:
        return
    users = load_user_ids()
    await message.answer(f"📊 Статистика бота:\n\n👥 Всего пользователей: {len(users)}")

@dp.message(F.text == "📢 Сделать рассылку")
async def cmd_broadcast_prompt(message: Message):
    if message.from_user.id != MY_ADMIN_ID:
        return
    broadcast_states.add(message.from_user.id)
    await message.answer("📢 Отправь текст рассылки следующим сообщением (все пользователи бота получат его).")

@dp.message(F.text == "📈 Создать презентацию")
async def cmd_ppt_prompt(message: Message):
    user_id = message.from_user.id
    if not await check_subscription(user_id):
        await message.answer(f"🔒 Сначала подпишись на каналы:\n1️⃣ {CHANNEL_1_URL}\n2️⃣ {CHANNEL_2_URL}", reply_markup=get_sub_keyboard())
        return
    ppt_states.add(user_id)
    excel_states.discard(user_id)
    user_ppt_images.pop(user_id, None)
    await message.answer(
        "📈 Отправь тему презентации следующим сообщением.\n\n"
        "💡 *Хочешь добавить свои картинки?* Сначала отправь фото, а затем отправь тему презентации!",
        parse_mode="Markdown"
    )

@dp.message(F.text == "📊 Создать таблицу Excel")
async def cmd_excel_prompt(message: Message):
    user_id = message.from_user.id
    if not await check_subscription(user_id):
        await message.answer(f"🔒 Сначала подпишись на каналы:\n1️⃣ {CHANNEL_1_URL}\n2️⃣ {CHANNEL_2_URL}", reply_markup=get_sub_keyboard())
        return
    excel_states.add(user_id)
    ppt_states.discard(user_id)
    await message.answer("📊 Опиши задачу или тему для таблицы следующим сообщением (например: *«Таблица успеваемости студентов с оценками»*), и я сформирую готовый Excel-файл (.xlsx)!", parse_mode="Markdown")

@dp.message(F.text == "⭐ Избранное")
async def cmd_favorites(message: Message):
    user_id = message.from_user.id
    if not await check_subscription(user_id):
        await message.answer(f"🔒 Сначала подпишись на каналы:\n1️⃣ {CHANNEL_1_URL}\n2️⃣ {CHANNEL_2_URL}", reply_markup=get_sub_keyboard())
        return

    favs = load_favorites()
    uid_str = str(user_id)
    user_favs = favs.get(uid_str, [])

    if not user_favs:
        await message.answer("⭐ У тебя пока нет сохраненных ответов в избранном.")
        return

    await message.answer(f"⭐ Твои сохраненные ответы ({len(user_favs)}):")
    for idx, fav_text in enumerate(user_favs, 1):
        display_text = f"Сохранение #{idx}:\n\n{fav_text}"
        if len(display_text) > 4000:
            display_text = display_text[:3997] + "..."
        await message.answer(display_text, disable_web_page_preview=True)

@dp.callback_query(F.data.in_({"btn_simplify", "btn_save_fav"}))
async def cb_answer_actions(callback: types.CallbackQuery):
    user_id = callback.from_user.id
    msg_text = callback.message.text or callback.message.caption or ""
    
    clean_msg = msg_text.split("—\n⚡")[0].strip() if "—\n⚡" in msg_text else msg_text.strip()
    if not clean_msg:
        await callback.answer("⚠️ Нечего обрабатывать!", show_alert=True)
        return

    if callback.data == "btn_save_fav":
        add_to_favorites(user_id, clean_msg)
        await callback.answer("⭐ Успешно сохранено в избранное!", show_alert=True)
    elif callback.data == "btn_simplify":
        await callback.answer("💡 Сжимаю до сути...")
        try:
            response = await groq_client.chat.completions.create(
                model=TEXT_MODEL,
                messages=[
                    {"role": "system", "content": "Объясни предельно коротко и просто в 3-5 предложениях."},
                    {"role": "user", "content": clean_msg}
                ],
                temperature=0.7
            )
            simplified_reply = clean_text_for_html(response.choices[0].message.content)
            full_reply = f"💡 <b>Коротко на пальцах:</b>\n\n{simplified_reply}{AD_FOOTER}"
            await callback.message.answer(full_reply, parse_mode="HTML", reply_markup=get_answer_inline_keyboard(), disable_web_page_preview=True)
        except Exception:
            await callback.message.answer("⚠️ Ошибка при обработке. Попробуй ещё раз. Если повторится, пиши - @mecau")

@dp.message(F.text == "📄 Скачать ответ в Word")
async def cmd_download_word(message: Message):
    user_id = message.from_user.id
    if not await check_subscription(user_id):
        await message.answer(f"🔒 Сначала подпишись на каналы:\n1️⃣ {CHANNEL_1_URL}\n2️⃣ {CHANNEL_2_URL}", reply_markup=get_sub_keyboard())
        return

    user_data = get_user_data(user_id)
    text_to_save = user_data.get("last_output", "").strip()
    if not text_to_save or text_to_save == "Здесь пока нет ответов.":
        await message.answer("⚠️ Нет данных для сохранения. Попробуй ещё раз. Если повторится, пиши - @mecau")
        return

    doc = Document()
    for section in doc.sections:
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)
        section.left_margin = Mm(30)
        section.right_margin = Mm(15)

    p = doc.add_paragraph()
    run = p.add_run(text_to_save)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)

    file_doc = BufferedInputFile(bio.read(), filename="MecauAI_Answer.docx")
    await message.answer_document(file_doc, caption="📄 Вот твой ответ в формате Word!")

@dp.message(F.text == "📑 Создать титульник ГОСТ")
async def cmd_gost_title(message: Message):
    if not await check_subscription(message.from_user.id):
        await message.answer(f"🔒 Сначала подпишись на каналы:\n1️⃣ {CHANNEL_1_URL}\n2️⃣ {CHANNEL_2_URL}", reply_markup=get_sub_keyboard())
        return
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📘 Индивидуальный проект", callback_data="gost_project")],
        [InlineKeyboardButton(text="📗 Курсовая работа",        callback_data="gost_coursework")],
        [InlineKeyboardButton(text="📕 Дипломная работа (ВКР)", callback_data="gost_diploma")],
        [InlineKeyboardButton(text="📙 Отчёт по практике",      callback_data="gost_practice")],
    ])
    await message.answer("📑 Выбери тип работы для титульника:", reply_markup=kb)

@dp.callback_query(F.data.startswith("gost_"))
async def cb_gost_generate(callback: types.CallbackQuery):
    work_type_map = {
        "gost_project":    ("ИНДИВИДУАЛЬНЫЙ ПРОЕКТ",                   "Индивидуальный_проект"),
        "gost_coursework": ("КУРСОВАЯ РАБОТА",                         "Курсовая_работа"),
        "gost_diploma":    ("ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА (ВКР)", "ВКР"),
        "gost_practice":   ("ОТЧЁТ ПО ПРАКТИКЕ",                      "Отчет_по_практике"),
    }
    work_label, filename_base = work_type_map[callback.data]

    doc = Document()
    for section in doc.sections:
        section.top_margin    = Mm(20)
        section.bottom_margin = Mm(20)
        section.left_margin   = Mm(30)
        section.right_margin  = Mm(15)

    def add_centered(text, size=14, bold=False):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(size)
        run.bold = bold
        return p

    def add_right(text, size=14):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(size)
        return p

    def add_empty(count=1):
        for _ in range(count):
            p = doc.add_paragraph()
            p.add_run("").font.name = 'Times New Roman'

    add_centered("ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ПРОФЕССИОНАЛЬНОЕ\nОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ\n«НАЗВАНИЕ КОЛЛЕДЖА»", size=14)
    add_empty(4)
    add_centered(work_label, size=14, bold=True)
    add_empty(1)
    add_centered("на тему:\n«Введи тему работы здесь»", size=14, bold=True)
    add_empty(6)
    add_right("Выполнил(а): студент(ка) группы ГРУППА\nФамилия Имя Отчество\n\nРуководитель:\nДолжность, Фамилия И.О.")
    add_empty(5)
    add_centered("Город — 2026", size=14)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)

    file_doc = BufferedInputFile(bio.read(), filename=f"Titulnik_{filename_base}.docx")
    await callback.message.answer_document(file_doc, caption=f"📑 Титульник «{work_label}» готов!")
    await callback.answer()

MENU_BUTTONS = {
    "📄 Скачать ответ в Word", "📑 Создать титульник ГОСТ",
    "📈 Создать презентацию", "📊 Создать таблицу Excel",
    "⭐ Избранное", "🔄 Сменить режим", "ℹ️ О MecauAI", 
    "📢 Сделать рассылку", "📊 Статистика бота"
}

@dp.message(F.photo)
async def handle_photo(message: Message):
    user_id = message.from_user.id
    if not await check_subscription(user_id):
        await message.answer(f"🔒 Сначала подпишись на каналы:\n1️⃣ {CHANNEL_1_URL}\n2️⃣ {CHANNEL_2_URL}", reply_markup=get_sub_keyboard())
        return

    save_user_id(user_id)
    photo = message.photo[-1]
    file_info = await bot.get_file(photo.file_id)
    downloaded_file = await bot.download_file(file_info.file_path)
    img_bytes = downloaded_file.read()

    if user_id in ppt_states:
        if user_id not in user_ppt_images:
            user_ppt_images[user_id] = []
        user_ppt_images[user_id].append(img_bytes)
        await message.answer(f"🖼 Картинка сохранена для презентации ({len(user_ppt_images[user_id])} шт.)! Теперь отправь тему презентации.")
        return

    await bot.send_chat_action(chat_id=message.chat.id, action="typing")
    try:
        base64_image = base64.b64encode(img_bytes).decode('utf-8')
        response = await groq_client.chat.completions.create(
            model=VISION_MODEL,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "text", "text": message.caption or "Подробно проанализируй это изображение, распознай текст если он есть и ответь по сути."},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}
                ]
            }]
        )
        reply = clean_text_for_html(response.choices[0].message.content)
        user_data = get_user_data(user_id)
        user_data["last_output"] = reply
        await message.answer(f"{reply}{AD_FOOTER}", parse_mode="HTML", reply_markup=get_answer_inline_keyboard(), disable_web_page_preview=True)
    except Exception as e:
        logging.error(f"Ошибка обработки фото: {e}")
        await message.answer("⚠️ Ошибка при обработке изображения. Попробуй ещё раз. Если повторится, пиши - @mecau")

@dp.message(F.text)
async def handle_text(message: Message):
    user_id = message.from_user.id

    if user_id == MY_ADMIN_ID and user_id in broadcast_states:
        broadcast_states.remove(user_id)
        users = load_user_ids()
        success, failed = 0, 0
        status_msg = await message.answer("📢 Начинаю рассылку...")
        for uid in users:
            try:
                await bot.send_message(uid, message.text, disable_web_page_preview=True)
                success += 1
                await asyncio.sleep(0.05)
            except Exception:
                failed += 1
        await status_msg.edit_text(f"✅ Рассылка завершена!\n\n👥 Доставлено: {success}\n❌ Ошибок: {failed}")
        return

        if user_id in ppt_states:
        ppt_states.remove(user_id)
        status_msg = await message.answer("📈 Генерирую развернутую презентацию и оформление, подожди немного...")
        try:
            await bot.send_chat_action(chat_id=message.chat.id, action="typing")
            user_images = user_ppt_images.pop(user_id, [])
            num_slides = max(len(user_images), 5) if user_images else 6

            response = await groq_client.chat.completions.create(
                model=TEXT_MODEL,
                messages=[
                    {
                        "role": "system",
                        "content": (
                            f"Составь подробную, качественную учебную презентацию ровно из {num_slides} слайдов на заданную тему. "
                            "Для каждого слайда пиши развернутые, информативные тезисы (по 3-4 полноценных предложения или емких пункта, несущих реальный смысл, а не просто заголовки). "
                            "Ответ выдай СТРОГО в формате валидного JSON без markdown-оформления (без ```json), в виде списка объектов: "
                            "[{{\"title\": \"Заголовок слайда\", \"points\": [\"Развернутый тезис с объяснением сути 1\", \"Развернутый тезис 2\"], \"image_prompt\": \"Detailed professional visual illustration of the slide topic in English\"}}]. "
                            "В поле image_prompt ВСЕГДА пиши качественный промпт на английском языке для генерации картинки."
                        )
                    },
                    {"role": "user", "content": f"Тема: {message.text}"}
                ],
                temperature=0.7
            )
            raw_content = clean_text_for_html(response.choices[0].message.content)
            
            if "```json" in raw_content:
                raw_content = raw_content.split("```json")[1].split("```")[0].strip()
            elif "```" in raw_content:
                raw_content = raw_content.split("```")[1].split("```")[0].strip()
            
            slides_data = json.loads(raw_content)
            
            prs = Presentation()
            prs.slide_width = PptxInches(13.333)
            prs.slide_height = PptxInches(7.5)

            # Импортируем цвета и фигуры для оформления внутри блока
            from pptx.dml.color import RGBColor
            from pptx.enum.shapes import MSO_SHAPE

            # Титульный слайд с оформлением
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            
            # Фоновая плашка на титульник
            bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
            bg_shape.fill.solid()
            bg_shape.fill.fore_color.rgb = RGBColor(24, 43, 73)  # Глубокий темно-синий
            bg_shape.line.color.rgb = RGBColor(24, 43, 73)

            # Текст титульника
            title_box = slide.shapes.add_textbox(PptxInches(1.0), PptxInches(2.2), PptxInches(11.333), PptxInches(3.0))
            tf_title = title_box.text_frame
            tf_title.word_wrap = True
            
            p_main = tf_title.paragraphs[0]
            p_main.text = message.text
            p_main.font.name = 'Times New Roman'
            p_main.font.size = PptxPt(36)
            p_main.font.bold = True
            p_main.font.color.rgb = RGBColor(255, 255, 255)
            p_main.alignment = WD_ALIGN_PARAGRAPH.CENTER

            p_sub = tf_title.add_paragraph()
            p_sub.text = "Презентация подготовлена с помощью MecauAI"
            p_sub.font.name = 'Times New Roman'
            p_sub.font.size = PptxPt(18)
            p_sub.font.color.rgb = RGBColor(200, 215, 235)
            p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

            async with aiohttp.ClientSession() as session:
                for idx, item in enumerate(slides_data):
                    s = prs.slides.add_slide(prs.slide_layouts[6])
                    
                    # Заголовок слайда
                    tb_title = s.shapes.add_textbox(PptxInches(0.8), PptxInches(0.5), PptxInches(11.7), PptxInches(0.9))
                    tf_t = tb_title.text_frame
                    tf_t.word_wrap = True
                    p_t = tf_t.paragraphs[0]
                    p_t.text = item.get("title", f"Слайд {idx+1}")
                    p_t.font.name = 'Times New Roman'
                    p_t.font.size = PptxPt(28)
                    p_t.font.bold = True
                    p_t.font.color.rgb = RGBColor(24, 43, 73)

                    # Акцентная линия под заголовком
                    line = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, PptxInches(0.8), PptxInches(1.4), PptxInches(2.5), PptxInches(0.05))
                    line.fill.solid()
                    line.fill.fore_color.rgb = RGBColor(79, 129, 189)
                    line.line.color.rgb = RGBColor(79, 129, 189)

                    img_stream = None
                    if user_images and idx < len(user_images):
                        img_stream = io.BytesIO(user_images[idx])
                    elif "image_prompt" in item and item["image_prompt"]:
                        try:
                            img_prompt = urllib.parse.quote(item["image_prompt"])
                            img_url = f"[https://gen.pollinations.ai/image/](https://gen.pollinations.ai/image/){img_prompt}?width=800&height=600&nologo=true&model=flux"
                            async with session.get(img_url, timeout=15) as resp:
                                if resp.status == 200:
                                    img_bytes = await resp.read()
                                    if len(img_bytes) > 1000:
                                        img_stream = io.BytesIO(img_bytes)
                        except Exception as gen_err:
                            logging.error(f"Не удалось сгенерировать картинку для слайда {idx}: {gen_err}")

                    has_img = img_stream is not None
                    
                    # Текстовый блок для тезисов (шире, если нет картинки)
                    content_width = PptxInches(7.2) if has_img else PptxInches(11.7)
                    tb_content = s.shapes.add_textbox(PptxInches(0.8), PptxInches(1.7), content_width, PptxInches(5.2))
                    tf = tb_content.text_frame
                    tf.word_wrap = True
                    
                    points = item.get("points", [])
                    for p_idx, pt in enumerate(points):
                        p = tf.add_paragraph() if p_idx > 0 else tf.paragraphs[0]
                        p.text = "▪  " + pt
                        p.font.name = 'Times New Roman'
                        p.font.size = PptxPt(16)
                        p.font.color.rgb = RGBColor(50, 50, 50)
                        p.space_after = PptxPt(14)  # Отступ между абзацами для читаемости

                    # Вставка картинки, если она есть
                    if has_img and img_stream:
                        try:
                            s.shapes.add_picture(img_stream, left=PptxInches(8.3), top=PptxInches(1.7), width=PptxInches(4.2))
                        except Exception as img_err:
@dp.message(F.text)
async def handle_text(message: Message):
    user_id = message.from_user.id

    if user_id == MY_ADMIN_ID and user_id in broadcast_states:
        broadcast_states.remove(user_id)
        users = load_user_ids()
        success, failed = 0, 0
        status_msg = await message.answer("📢 Начинаю рассылку...")
        for uid in users:
            try:
                await bot.send_message(uid, message.text, disable_web_page_preview=True)
                success += 1
                await asyncio.sleep(0.05)
            except Exception:
                failed += 1
        await status_msg.edit_text(f"✅ Рассылка завершена!\n\n👥 Доставлено: {success}\n❌ Ошибок: {failed}")
        return

    if user_id in ppt_states:
        ppt_states.remove(user_id)
        status_msg = await message.answer("📈 Генерирую развернутую презентацию и оформление, подожди немного...")
        try:
            await bot.send_chat_action(chat_id=message.chat.id, action="typing")
            user_images = user_ppt_images.pop(user_id, [])
            num_slides = max(len(user_images), 5) if user_images else 6

            response = await groq_client.chat.completions.create(
                model=TEXT_MODEL,
                messages=[
                    {
                        "role": "system",
                        "content": (
                            f"Составь подробную, качественную учебную презентацию ровно из {num_slides} слайдов на заданную тему. "
                            "Для каждого слайда пиши развернутые, информативные тезисы (по 3-4 полноценных предложения или емких пункта, несущих реальный смысл, а не просто заголовки). "
                            "Ответ выдай СТРОГО в формате валидного JSON без markdown-оформления (без ```json), в виде списка объектов: "
                            "[{\"title\": \"Заголовок слайда\", \"points\": [\"Развернутый тезис с объяснением сути 1\", \"Развернутый тезис 2\"], \"image_prompt\": \"Detailed professional visual illustration of the slide topic in English\"}]. "
                            "В поле image_prompt ВСЕГДА пиши качественный промпт на английском языке для генерации картинки."
                        )
                    },
                    {"role": "user", "content": f"Тема: {message.text}"}
                ],
                temperature=0.7
            )
            raw_content = clean_text_for_html(response.choices[0].message.content)
            
            if "```json" in raw_content:
                raw_content = raw_content.split("```json")[1].split("```")[0].strip()
            elif "```" in raw_content:
                raw_content = raw_content.split("```")[1].split("```")[0].strip()
            
            slides_data = json.loads(raw_content)
            
            prs = Presentation()
            prs.slide_width = PptxInches(13.333)
            prs.slide_height = PptxInches(7.5)

            from pptx.dml.color import RGBColor
            from pptx.enum.shapes import MSO_SHAPE

            slide = prs.slides.add_slide(prs.slide_layouts[6])
            
            bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
            bg_shape.fill.solid()
            bg_shape.fill.fore_color.rgb = RGBColor(24, 43, 73)
            bg_shape.line.color.rgb = RGBColor(24, 43, 73)

            title_box = slide.shapes.add_textbox(PptxInches(1.0), PptxInches(2.2), PptxInches(11.333), PptxInches(3.0))
            tf_title = title_box.text_frame
            tf_title.word_wrap = True
            
            p_main = tf_title.paragraphs[0]
            p_main.text = message.text
            p_main.font.name = 'Times New Roman'
            p_main.font.size = PptxPt(36)
            p_main.font.bold = True
            p_main.font.color.rgb = RGBColor(255, 255, 255)
            p_main.alignment = WD_ALIGN_PARAGRAPH.CENTER

            p_sub = tf_title.add_paragraph()
            p_sub.text = "Презентация подготовлена с помощью MecauAI"
            p_sub.font.name = 'Times New Roman'
            p_sub.font.size = PptxPt(18)
            p_sub.font.color.rgb = RGBColor(200, 215, 235)
            p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

            async with aiohttp.ClientSession() as session:
                for idx, item in enumerate(slides_data):
                    s = prs.slides.add_slide(prs.slide_layouts[6])
                    
                    tb_title = s.shapes.add_textbox(PptxInches(0.8), PptxInches(0.5), PptxInches(11.7), PptxInches(0.9))
                    tf_t = tb_title.text_frame
                    tf_t.word_wrap = True
                    p_t = tf_t.paragraphs[0]
                    p_t.text = item.get("title", f"Слайд {idx+1}")
                    p_t.font.name = 'Times New Roman'
                    p_t.font.size = PptxPt(28)
                    p_t.font.bold = True
                    p_t.font.color.rgb = RGBColor(24, 43, 73)

                    line = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, PptxInches(0.8), PptxInches(1.4), PptxInches(2.5), PptxInches(0.05))
                    line.fill.solid()
                    line.fill.fore_color.rgb = RGBColor(79, 129, 189)
                    line.line.color.rgb = RGBColor(79, 129, 189)

                    img_stream = None
                    if user_images and idx < len(user_images):
                        img_stream = io.BytesIO(user_images[idx])
                    elif "image_prompt" in item and item["image_prompt"]:
                        try:
                            img_prompt = urllib.parse.quote(item["image_prompt"])
                            img_url = f"[https://gen.pollinations.ai/image/](https://gen.pollinations.ai/image/){img_prompt}?width=800&height=600&nologo=true&model=flux"
                            async with session.get(img_url, timeout=15) as resp:
                                if resp.status == 200:
                                    img_bytes = await resp.read()
                                    if len(img_bytes) > 1000:
                                        img_stream = io.BytesIO(img_bytes)
                        except Exception as gen_err:
                            logging.error(f"Не удалось сгенерировать картинку для слайда {idx}: {gen_err}")

                    has_img = img_stream is not None
                    
                    content_width = PptxInches(7.2) if has_img else PptxInches(11.7)
                    tb_content = s.shapes.add_textbox(PptxInches(0.8), PptxInches(1.7), content_width, PptxInches(5.2))
                    tf = tb_content.text_frame
                    tf.word_wrap = True
                    
                    points = item.get("points", [])
                    for p_idx, pt in enumerate(points):
                        p = tf.add_paragraph() if p_idx > 0 else tf.paragraphs[0]
                        p.text = "▪  " + pt
                        p.font.name = 'Times New Roman'
                        p.font.size = PptxPt(16)
                        p.font.color.rgb = RGBColor(50, 50, 50)
                        p.space_after = PptxPt(14)

                    if has_img and img_stream:
                        try:
                            s.shapes.add_picture(img_stream, left=PptxInches(8.3), top=PptxInches(1.7), width=PptxInches(4.2))
                        except Exception as img_err:
                            logging.error(f"Не удалось вставить картинку на слайд {idx}: {img_err}")

            bio = io.BytesIO()
            prs.save(bio)
            bio.seek(0)
            file_doc = BufferedInputFile(bio.read(), filename="Presentation_Styled.pptx")
            
            await status_msg.delete()
            await message.answer_document(file_doc, caption="📈 Развернутая и оформленная презентация готова!")
            
        except Exception as e:
            logging.error(f"Критическая ошибка при генерации презентации: {e}", exc_info=True)
            try:
                await status_msg.delete()
            except Exception:
                pass
            await message.answer("⚠️ Произошла ошибка при создании презентации. Попробуй еще раз.")
        return

    if user_id in excel_states:
        excel_states.remove(user_id)
        status_msg = await message.answer("📊 Генерирую таблицу Excel, подожди секунду...")
        try:
            await bot.send_chat_action(chat_id=message.chat.id, action="typing")
            response = await groq_client.chat.completions.create(
                model=TEXT_MODEL,
                messages=[
                    {
                        "role": "system",
                        "content": "Создай структуру таблицы на основе запроса пользователя. Ответ выдай СТРОГО в формате валидного JSON без markdown-оформления (без ```json), в виде объекта с ключом 'headers' (список строк-заголовков) и ключом 'rows' (список списков с данными для строк таблицы). Пример: {\"headers\": [\"№\", \"Название\", \"Значение\"], \"rows\": [[1, \"Пример 1\", 100], [2, \"Пример 2\", 200]]}"
                    },
                    {"role": "user", "content": f"Задача для таблицы: {message.text}"}
                ],
                temperature=0.7
            )
            raw_content = clean_text_for_html(response.choices[0].message.content)
            if "```json" in raw_content:
                raw_content = raw_content.split("```json")[1].split("```")[0].strip()
            elif "```" in raw_content:
                raw_content = raw_content.split("```")[1].split("```")[0].strip()
            
            table_data = json.loads(raw_content)
            headers = table_data.get("headers", ["Колонка 1", "Колонка 2"])
            rows = table_data.get("rows", [["Данные 1", "Данные 2"]])

            wb = Workbook()
            ws = wb.active
            ws.title = "Таблица"

            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
            align_center = Alignment(horizontal="center", vertical="center")
            align_left = Alignment(horizontal="left", vertical="center")
            border_thin = Border(
                left=Side(style='thin', color='D9D9D9'),
                right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'),
                bottom=Side(style='thin', color='D9D9D9')
            )

            ws.append(headers)
            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=1, column=col_num)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = align_center
                cell.border = border_thin

            for row_idx, row_data in enumerate(rows, start=2):
                ws.append(row_data)
                for col_num in range(1, len(row_data) + 1):
                    cell = ws.cell(row=row_idx, column=col_num)
                    cell.font = Font(name="Calibri", size=11)
                    cell.alignment = align_left
                    cell.border = border_thin

            for col in ws.columns:
                max_length = 0
                column_letter = col[0].column_letter
                for cell in col:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except Exception:
                        pass
                ws.column_dimensions[column_letter].width = max(max_length + 4, 12)

            bio = io.BytesIO()
            wb.save(bio)
            bio.seek(0)
            file_doc = BufferedInputFile(bio.read(), filename="MecauAI_Table.xlsx")
            
            await status_msg.delete()
            await message.answer_document(file_doc, caption="📊 Твоя таблица Excel (.xlsx) готова!")
        except Exception as e:
            logging.error(f"Ошибка при создании Excel: {e}", exc_info=True)
            try:
                await status_msg.delete()
            except Exception:
                pass
            await message.answer("⚠️ Произошла ошибка при создании таблицы Excel. Попробуй еще раз.")
        return

    if message.text in MENU_BUTTONS:
        return

    if not await check_subscription(user_id):
        await message.answer(f"🔒 Сначала подпишись на каналы:\n1️⃣ {CHANNEL_1_URL}\n2️⃣ {CHANNEL_2_URL}", reply_markup=get_sub_keyboard())
        return

    save_user_id(user_id)
    user_data = get_user_data(user_id)
    current_mode = user_data["mode"]

    system_prompt = PROMPTS.get(current_mode, PROMPTS["ai"])

    user_data["history"].append({"role": "user", "content": message.text})
    if len(user_data["history"]) > 6:
        user_data["history"] = user_data["history"][-6:]

    messages_payload = [{"role": "system", "content": system_prompt}] + user_data["history"]

    await bot.send_chat_action(chat_id=message.chat.id, action="typing")
    try:
        response = await groq_client.chat.completions.create(
            model=TEXT_MODEL,
            messages=messages_payload,
            temperature=0.7
        )
        ai_reply = clean_text_for_html(response.choices[0].message.content)
        if not ai_reply:
            ai_reply = "Готово."

        user_data["history"].append({"role": "assistant", "content": ai_reply})
        user_data["last_output"] = ai_reply

        if current_mode == "coder":
            full_message = f"{ai_reply}\n\n{AD_FOOTER}"
            await message.answer(full_message, parse_mode="Markdown", reply_markup=get_answer_inline_keyboard(), disable_web_page_preview=True)
        else:
            full_message = f"{ai_reply}{AD_FOOTER}"
            await message.answer(full_message, parse_mode="HTML", reply_markup=get_answer_inline_keyboard(), disable_web_page_preview=True)
    except Exception:
        await message.answer("⚠️ Ошибка при обработке запроса. Попробуй ещё раз. Если повторится, пиши - @mecau")



    if user_id in excel_states:
        excel_states.remove(user_id)
        status_msg = await message.answer("📊 Генерирую таблицу Excel, подожди секунду...")
        try:
            await bot.send_chat_action(chat_id=message.chat.id, action="typing")
            response = await groq_client.chat.completions.create(
                model=TEXT_MODEL,
                messages=[
                    {
                        "role": "system",
                        "content": "Создай структуру таблицы на основе запроса пользователя. Ответ выдай СТРОГО в формате валидного JSON без markdown-оформления (без ```json), в виде объекта с ключом 'headers' (список строк-заголовков) и ключом 'rows' (список списков с данными для строк таблицы). Пример: {\"headers\": [\"№\", \"Название\", \"Значение\"], \"rows\": [[1, \"Пример 1\", 100], [2, \"Пример 2\", 200]]}"
                    },
                    {"role": "user", "content": f"Задача для таблицы: {message.text}"}
                ],
                temperature=0.7
            )
            raw_content = clean_text_for_html(response.choices[0].message.content)
            if "```json" in raw_content:
                raw_content = raw_content.split("```json")[1].split("```")[0].strip()
            elif "```" in raw_content:
                raw_content = raw_content.split("```")[1].split("```")[0].strip()
            
            table_data = json.loads(raw_content)
            headers = table_data.get("headers", ["Колонка 1", "Колонка 2"])
            rows = table_data.get("rows", [["Данные 1", "Данные 2"]])

            wb = Workbook()
            ws = wb.active
            ws.title = "Таблица"

            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
            align_center = Alignment(horizontal="center", vertical="center")
            align_left = Alignment(horizontal="left", vertical="center")
            border_thin = Border(
                left=Side(style='thin', color='D9D9D9'),
                right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'),
                bottom=Side(style='thin', color='D9D9D9')
            )

            ws.append(headers)
            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=1, column=col_num)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = align_center
                cell.border = border_thin

            for row_idx, row_data in enumerate(rows, start=2):
                ws.append(row_data)
                for col_num in range(1, len(row_data) + 1):
                    cell = ws.cell(row=row_idx, column=col_num)
                    cell.font = Font(name="Calibri", size=11)
                    cell.alignment = align_left
                    cell.border = border_thin

            for col in ws.columns:
                max_length = 0
                column_letter = col[0].column_letter
                for cell in col:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except Exception:
                        pass
                ws.column_dimensions[column_letter].width = max(max_length + 4, 12)

            bio = io.BytesIO()
            wb.save(bio)
            bio.seek(0)
            file_doc = BufferedInputFile(bio.read(), filename="MecauAI_Table.xlsx")
            
            await status_msg.delete()
            await message.answer_document(file_doc, caption="📊 Твоя таблица Excel (.xlsx) готова!")
        except Exception as e:
            logging.error(f"Ошибка при создании Excel: {e}", exc_info=True)
            try:
                await status_msg.delete()
            except Exception:
                pass
            await message.answer("⚠️ Произошла ошибка при создании таблицы Excel. Попробуй еще раз.")
        return

    if message.text in MENU_BUTTONS:
        return

    if not await check_subscription(user_id):
        await message.answer(f"🔒 Сначала подпишись на каналы:\n1️⃣ {CHANNEL_1_URL}\n2️⃣ {CHANNEL_2_URL}", reply_markup=get_sub_keyboard())
        return

    save_user_id(user_id)
    user_data = get_user_data(user_id)
    current_mode = user_data["mode"]

    system_prompt = PROMPTS.get(current_mode, PROMPTS["ai"])

    user_data["history"].append({"role": "user", "content": message.text})
    if len(user_data["history"]) > 6:
        user_data["history"] = user_data["history"][-6:]

    messages_payload = [{"role": "system", "content": system_prompt}] + user_data["history"]

    await bot.send_chat_action(chat_id=message.chat.id, action="typing")
    try:
        response = await groq_client.chat.completions.create(
            model=TEXT_MODEL,
            messages=messages_payload,
            temperature=0.7
        )
        ai_reply = clean_text_for_html(response.choices[0].message.content)
        if not ai_reply:
            ai_reply = "Готово."

        user_data["history"].append({"role": "assistant", "content": ai_reply})
        user_data["last_output"] = ai_reply

        if current_mode == "coder":
            full_message = f"{ai_reply}\n\n{AD_FOOTER}"
            await message.answer(full_message, parse_mode="Markdown", reply_markup=get_answer_inline_keyboard(), disable_web_page_preview=True)
        else:
            full_message = f"{ai_reply}{AD_FOOTER}"
            await message.answer(full_message, parse_mode="HTML", reply_markup=get_answer_inline_keyboard(), disable_web_page_preview=True)
    except Exception:
        await message.answer("⚠️ Ошибка при обработке запроса. Попробуй ещё раз. Если повторится, пиши - @mecau")

async def main():
    await bot.set_my_commands([
        BotCommand(command="start", description="Перезапустить бота"),
        BotCommand(command="mode",  description="Сменить режим ассистента"),
        BotCommand(command="about", description="О возможностях"),
    ])
    await bot.delete_webhook()
    print("🚀 Бот MecauAI запущен и слушает обновления...")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
