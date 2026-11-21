import os

BOT_TOKEN = (os.getenv("BOT_TOKEN") or "").strip()
GROQ_API_KEY = (os.getenv("GROQ_API_KEY") or "").strip()
OPENAI_API_KEY = (os.getenv("OPENAI_API_KEY") or "").strip()
DEEPSEEK_API_KEY = (os.getenv("DEEPSEEK_API_KEY") or "").strip()

AI_PROVIDER = (os.getenv("MECAUAI_AI_PROVIDER") or "auto").strip().lower()
OPENAI_TEXT_MODEL = (os.getenv("OPENAI_TEXT_MODEL") or "gpt-4.1-mini").strip()
OPENAI_VISION_MODEL = (os.getenv("OPENAI_VISION_MODEL") or "gpt-4o-mini").strip()
DEEPSEEK_TEXT_MODEL = (os.getenv("DEEPSEEK_TEXT_MODEL") or "deepseek-chat").strip()

CHANNEL_1_USERNAME = (os.getenv("MECAUAI_CHANNEL_1_USERNAME") or "MecauAI").strip().lstrip("@") or "MecauAI"
CHANNEL_1_URL = (os.getenv("MECAUAI_CHANNEL_1_URL") or f"https://t.me/{CHANNEL_1_USERNAME}").strip()

CHANNEL_2_USERNAME = (os.getenv("MECAUAI_CHANNEL_2_USERNAME") or "mecauinfo").strip().lstrip("@") or "mecauinfo"
CHANNEL_2_URL = (os.getenv("MECAUAI_CHANNEL_2_URL") or f"https://t.me/{CHANNEL_2_USERNAME}").strip()


TEXT_MODEL = (os.getenv("MECAUAI_TEXT_MODEL") or "llama-3.3-70b-versatile").strip()
VISION_MODEL = (os.getenv("MECAUAI_VISION_MODEL") or "qwen/qwen3.6-27b").strip()


# Promotional footer disabled: bot replies must contain no unsolicited advertising.
AD_FOOTER = ""

PROMPTS = {
    "coder": (
        "Ты — MecauAI Code, опытный инженер-программист. Пиши корректный, безопасный и поддерживаемый код. "
        "Перед ответом проверь логику, крайние случаи, обработку ошибок и совместимость версий. "
        "Не выдумывай результаты запуска и тестирования: прямо отделяй проверенное от предположений. "
        "Указывай зависимости и способ запуска, когда это необходимо. Код всегда помещай в Markdown-блок с языком. "
        "Для больших задач предлагай целостное решение, а не обрывки, и добавляй тесты для критичной логики. "
        "Избегай лишней болтовни."
    ),
    "ai": (
        "Ты — MecauAI, точный и доброжелательный интеллектуальный помощник. "
        "Отвечай на языке пользователя и соблюдай запрошенный объём. "
        "Помогай разбираться в учебных и практических вопросах: объясняй ход решения, "
        "делай структурированные конспекты и давай понятные примеры. "
        "Не выдумывай факты, цифры, цитаты, источники и результаты проверок. "
        "Отделяй достоверные сведения от предположений; при нехватке данных честно укажи ограничение. "
        "Сложные ответы структурируй заголовками и списками, простые вопросы не перегружай."
    ),
    "friend": (
        "Ты — дружелюбный собеседник и персональный тьютор MecauAI. "
        "Общайся естественно, тепло и уважительно, без искусственной фамильярности. "
        "Отвечай коротко, если пользователь не просит подробностей. Не обесценивай чувства "
        "и не выдавай догадки за факты."
    )
}
