import asyncio
import logging
import io
import urllib.parse
import json
import os
from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, BufferedInputFile, InlineKeyboardMarkup, InlineKeyboardButton, CallbackQuery
from aiogram.filters import Command
from aiogram.fsm.storage.memory import MemoryStorage
from groq import AsyncGroq
import aiohttp

# Библиотеки для работы с документами
from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt
from pptx.enum.text import WD_ALIGN_PARAGRAPH
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ================= НАСТРОЙКИ И ТОКЕНЫ =================
TOKEN = os.getenv("BOT_TOKEN", "ТВОЙ_ТОКЕН_БОТА")
GROQ_API_KEY = os.getenv("GROQ_API_KEY", "ТВОЙ_API_KEY_GROQ")

TEXT_MODEL = "llama-3.3-70b-versatile"
VISION_MODEL = "qwen/qwen3.6-27b"

MY_ADMIN_ID = 123456789  # Замени на свой Telegram ID при необходимости

CHANNEL_1_URL = "https://t.me/ твои_каналы_1"
CHANNEL_2_URL = "https://t.me/ твои_каналы_2"
CHANNEL_1_ID = "@твой_канал_1"
CHANNEL_2_ID = "@твой_канал_2"

logging.basicConfig(level=logging.INFO)
bot = Bot(token=TOKEN)
storage = MemoryStorage()
dp = Dispatcher(storage=storage)
groq_client = AsyncGroq(api_key=GROQ_API_KEY)

# Базы состояний и данных в памяти
user_ppt_images = {}
ppt_states = set()
excel_states = set()
broadcast_states = set()
USERS_FILE = "users.json"

AD_FOOTER = "\n\n───────────────\n🤖 *MecauAI* — твой умный помощник для учебы и работы."

MENU_BUTTONS = [
    "🤖 ИИ-Помощник", "💻 ИИ-Программист", "📝 Реферат / Эссе", 
    "📈 Презентация", "📊 Таблица Excel", "ℹ️ О боте", "💎 Подписка"
]

PROMPTS = {
    "ai": "Ты — MecauAI, универсальный и дружелюбный помощник для студентов. Отвечай понятно, емко и структурировано.",
    "coder": "Ты — продвинутый ИИ-программист. Пиши чистый, рабочий код (преимущественно на Python), объясняй логику и оформляйте блоки кода правильно.",
    "essay": "Ты — академический писатель. Помогай структурированно писать рефераты, эссе, курсовые и доклады, соблюдая строгий научный стиль."
}

def load_user_ids():
    if not os.path.exists(USERS_FILE):
        return []
    try:
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def save_user_id(user_id):
    users = load_user_ids()
    if user_id not in users:
        users.append(user_id)
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump(users, f)

def get_user_data(user_id):
    if not hasattr(dp, "user_sessions"):
        dp.user_sessions = {}
    if user_id not in dp.user_sessions:
        dp.user_sessions[user_id] = {
            "mode": "ai",
            "history": [],
            "last_output": ""
        }
    return dp.user_sessions[user_id]

async def check_subscription(user_id: int) -> bool:
    if user_id == MY_ADMIN_ID:
        return True
    try:
        member1 = await bot.get_chat_member(chat_id=CHANNEL_1_ID, user_id=user_id)
        member2 = await bot.get_chat_member(chat_id=CHANNEL_2_ID, user_id=user_id)
        
        not_allowed = ['left', 'kicked', 'restricted']
        if member1.status in not_allowed or member2.status in not_allowed:
            return False
        return True
    except Exception:
        return True

def clean_text_for_html(text: str) -> str:
    return text

def get_main_keyboard():
    from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="🤖 ИИ-Помощник"), KeyboardButton(text="💻 ИИ-Программист")],
            [KeyboardButton(text="📝 Реферат / Эссе"), KeyboardButton(text="📈 Презентация")],
            [KeyboardButton(text="📊 Таблица Excel"), KeyboardButton(text="ℹ️ О боте")],
            [KeyboardButton(text="💎 Подписка")]
        ],
        resize_keyboard=True
    )

def get_sub_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📢 Канал 1", url=CHANNEL_1_URL), InlineKeyboardButton(text="📢 Канал 2", url=CHANNEL_2_URL)],
        [InlineKeyboardButton(text="✅ Проверить подписку", callback_data="check_sub")]
    ])

def get_answer_inline_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="📄 Скачать в Word", callback_data="export_word"),
            InlineKeyboardButton(text="📊 В таблицу Excel", callback_data="export_excel")
        ]
    ])

# ================= ГЕНЕРАЦИЯ ПРЕЗЕНТАЦИИ =================
async def generate_presentation(message: Message, user_images: list):
    status_msg = await message.answer("📈 Генерирую развернутую презентацию и оформление, подожди немного...")
    try:
        await message.bot.send_chat_action(chat_id=message.chat.id, action="typing")
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

# ================= ХЕНДЛЕРЫ =================
@dp.message(Command("start"))
async def cmd_start(message: Message):
    user_id = message.from_user.id
    save_user_id(user_id)
    if not await check_subscription(user_id):
        await message.answer(f"🔒 Сначала подпишись на каналы:\n1️⃣ {CHANNEL_1_URL}\n2️⃣ {CHANNEL_2_URL}", reply_markup=get_sub_keyboard())
        return
    
    welcome_text = (
        "👋 Привет! Я **MecauAI** — твой персональный ИИ-помощник для учебы, написания кода и работы.\n\n"
        "Выбери нужный режим в меню снизу или просто отправь мне свой вопрос!"
    )
    await message.answer(welcome_text, reply_markup=get_main_keyboard(), parse_mode="Markdown")

@dp.message(Command("broadcast"))
async def cmd_broadcast(message: Message):
    if message.from_user.id != MY_ADMIN_ID:
        return
    broadcast_states.add(message.from_user.id)
    await message.answer("📢 Введи текст для рассылки всем пользователям бота:")

@dp.message(F.text == "🤖 ИИ-Помощник")
async def mode_ai(message: Message):
    user_data = get_user_data(message.from_user.id)
    user_data["mode"] = "ai"
    await message.answer("🤖 Режим: **ИИ-Помощник** активирован. Задавай любые вопросы!", parse_mode="Markdown")

@dp.message(F.text == "💻 ИИ-Программист")
async def mode_coder(message: Message):
    user_data = get_user_data(message.from_user.id)
    user_data["mode"] = "coder"
    await message.answer("💻 Режим: **ИИ-Программист** активирован. Скидывай задачи по коду и багам!", parse_mode="Markdown")

@dp.message(F.text == "📝 Реферат / Эссе")
async def mode_essay(message: Message):
    user_data = get_user_data(message.from_user.id)
    user_data["mode"] = "essay"
    await message.answer("📝 Режим: **Реферат / Эссе** активирован. Напиши тему, и я помогу составить текст.", parse_mode="Markdown")

@dp.message(F.text == "📈 Презентация")
async def mode_ppt(message: Message):
    user_id = message.from_user.id
    ppt_states.add(user_id)
    await message.answer("📈 Отправь тему презентации (и можешь приложить файлы/картинки, если нужно), и я соберу оформленный `.pptx` файл!")

@dp.message(F.text == "📊 Таблица Excel")
async def mode_excel(message: Message):
    user_id = message.from_user.id
    excel_states.add(user_id)
    await message.answer("📊 Опиши задачу или данные для таблицы, и я сформирую `.xlsx` файл со стилями!")

@dp.message(F.text == "ℹ️ О боте")
async def info_bot(message: Message):
    await message.answer("🤖 **MecauAI** — твой надежный помощник в учебе, генерации контента, презентаций и кода.", parse_mode="Markdown")

@dp.message(F.text == "💎 Подписка")
async def sub_info(message: Message):
    await message.answer(f"💎 Проверка подписки на каналы:\n1️⃣ {CHANNEL_1_URL}\n2️⃣ {CHANNEL_2_URL}", reply_markup=get_sub_keyboard())

@dp.callback_query(F.data == "check_sub")
async def callback_check_sub(callback: CallbackQuery):
    if await check_subscription(callback.from_user.id):
        await callback.message.answer("✅ Подписка подтверждена! Приятного использования.", reply_markup=get_main_keyboard())
        await callback.message.delete()
    else:
        await callback.answer("❌ Ты подписался еще не на все каналы!", show_alert=True)

@dp.callback_query(F.data == "export_word")
async def callback_export_word(callback: CallbackQuery):
    user_data = get_user_data(callback.from_user.id)
    last_text = user_data.get("last_output", "")
    if not last_text:
        await callback.answer("❌ Нет текста для экспорта!", show_alert=True)
        return
    
    from docx import Document
    doc = Document()
    doc.add_heading("MecauAI - Документ", level=1)
    doc.add_paragraph(last_text)
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    file_doc = BufferedInputFile(bio.read(), filename="MecauAI_Document.docx")
    
    await callback.message.answer_document(file_doc, caption="📄 Твой документ Word (.docx) готов!")
    await callback.answer()

@dp.callback_query(F.data == "export_excel")
async def callback_export_excel(callback: CallbackQuery):
    user_data = get_user_data(callback.from_user.id)
    last_text = user_data.get("last_output", "")
    if not last_text:
        await callback.answer("❌ Нет данных для экспорта!", show_alert=True)
        return
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Данные"
    ws.append(["Результат работы MecauAI"])
    for line in last_text.split("\n"):
        if line.strip():
            ws.append([line.strip()])
            
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    file_doc = BufferedInputFile(bio.read(), filename="MecauAI_Data.xlsx")
    
    await callback.message.answer_document(file_doc, caption="📊 Твоя таблица Excel (.xlsx) готова!")
    await callback.answer()

@dp.message(F.photo)
async def handle_photo(message: Message):
    user_id = message.from_user.id
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

    # Vision обработка через Groq / Qwen
    status_msg = await message.answer("👁 Анализирую изображение...")
    try:
        import base64
        base64_image = base64.b64encode(img_bytes).decode('utf-8')
        
        response = await groq_client.chat.completions.create(
            model=VISION_MODEL,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": message.caption or "Опиши подробно, что изображено на картинке, или реши задачу, если она тут есть."},
                        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}
                    ]
                }
            ],
            temperature=0.7
        )
        ai_reply = clean_text_for_html(response.choices[0].message.content)
        await status_msg.delete()
        await message.answer(f"{ai_reply}{AD_FOOTER}", parse_mode="HTML", reply_markup=get_answer_inline_keyboard(), disable_web_page_preview=True)
    except Exception as e:
        logging.error(f"Ошибка Vision: {e}")
        try:
            await status_msg.delete()
        except Exception:
            pass
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
        user_images = user_ppt_images.pop(user_id, [])
        await generate_presentation(message, user_images)
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

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
