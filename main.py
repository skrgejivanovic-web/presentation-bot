import os
import logging
import asyncio
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from pptx import Presentation
from pptx.util import Inches
from openai import OpenAI
import tempfile
from aiohttp import web

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

client = OpenAI(
    api_key=GROQ_API_KEY,
    base_url="https://api.groq.com/openai/v1"
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! Я бот, который создаёт презентации.\n\n"
        "Просто напиши тему, например:\n"
        "«Искусственный интеллект в образовании»\n\n"
        "Я сделаю презентацию и пришлю файл .pptx"
    )

async def generate_presentation(update: Update, context: ContextTypes.DEFAULT_TYPE):
    topic = update.message.text.strip()
    
    if not topic or len(topic) < 3:
        await update.message.reply_text("Напиши тему чуть подробнее.")
        return

    await update.message.reply_text("Генерирую презентацию... Подожди 20–40 секунд.")

    try:
        prompt = f"""
Ты эксперт по созданию презентаций. 
Создай структуру презентации на русском языке на тему: "{topic}".

Верни ТОЛЬКО в таком формате (без лишнего текста):

СЛАЙД 1: Заголовок
- Пункт 1
- Пункт 2
- Пункт 3

СЛАЙД 2: Заголовок
- Пункт 1
- Пункт 2
- Пункт 3

Сделай 6-8 слайдов. Первый слайд — титульный.
"""

        response = client.chat.completions.create(
            model="openai/gpt-oss-20b",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.7,
            max_tokens=2000
        )

        content = response.choices[0].message.content

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        slides_data = []
        current_slide = None

        for line in content.split("\n"):
            line = line.strip()
            if line.upper().startswith("СЛАЙД"):
                if current_slide:
                    slides_data.append(current_slide)
                title = line.split(":", 1)[-1].strip()
                current_slide = {"title": title, "points": []}
            elif line.startswith("-") and current_slide is not None:
                current_slide["points"].append(line[1:].strip())

        if current_slide:
            slides_data.append(current_slide)

        if slides_data:
            slide_layout = prs.slide_layouts[0]
            slide = prs.slides.add_slide(slide_layout)
            slide.shapes.title.text = slides_data[0]["title"]
            if len(slide.placeholders) > 1:
                slide.placeholders[1].text = "Сгенерировано ИИ-ботом"

        for slide_info in slides_data[1:]:
            slide_layout = prs.slide_layouts[1]
            slide = prs.slides.add_slide(slide_layout)
            slide.shapes.title.text = slide_info["title"]

            body = slide.placeholders[1]
            tf = body.text_frame
            tf.clear()

            for i, point in enumerate(slide_info["points"]):
                if i == 0:
                    tf.text = point
                else:
                    p = tf.add_paragraph()
                    p.text = point
                    p.level = 0

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            prs.save(tmp.name)
            tmp_path = tmp.name

        with open(tmp_path, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=f"{topic[:50]}.pptx",
                caption=f"Презентация: {topic}"
            )

        os.remove(tmp_path)

    except Exception as e:
        logger.error(f"Ошибка: {e}")
        await update.message.reply_text("Произошла ошибка. Попробуй ещё раз.")

async def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, generate_presentation))

    # Запускаем бота
    await app.initialize()
    await app.start()
    await app.updater.start_polling()

    # Простой веб-сервер, чтобы Render не убивал процесс
    async def handle(request):
        return web.Response(text="Bot is running")

    web_app = web.Application()
    web_app.router.add_get("/", handle)
    runner = web.AppRunner(web_app)
    await runner.setup()
    port = int(os.environ.get("PORT", 10000))
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()

    # Держим бота живым
    await asyncio.Event().wait()

if __name__ == "__main__":
    asyncio.run(main())