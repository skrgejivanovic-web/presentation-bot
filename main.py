import os
import logging
import asyncio
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RgbColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
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

# Цвета
DARK_BG = RgbColor(15, 23, 42)       # Тёмно-синий фон
ACCENT = RgbColor(56, 189, 248)      # Голубой акцент
WHITE = RgbColor(255, 255, 255)
LIGHT_GRAY = RgbColor(203, 213, 225)

def set_slide_background(slide, color):
    background = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0,
        slide.slide_layout.width if hasattr(slide.slide_layout, 'width') else Inches(13.333),
        Inches(7.5)
    )
    background.fill.solid()
    background.fill.fore_color.rgb = color
    background.line.fill.background()
    # Отправляем фигуру на задний план
    spTree = slide.shapes._spTree
    sp = background._element
    spTree.remove(sp)
    spTree.insert(2, sp)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! Я бот, который создаёт красивые презентации.\n\n"
        "Просто напиши тему, например:\n"
        "«Искусственный интеллект в образовании»\n\n"
        "Я сделаю презентацию и пришлю файл .pptx"
    )

async def generate_presentation(update: Update, context: ContextTypes.DEFAULT_TYPE):
    topic = update.message.text.strip()
    
    if not topic or len(topic) < 3:
        await update.message.reply_text("Напиши тему чуть подробнее.")
        return

    await update.message.reply_text("Генерирую красивую презентацию... Подожди 20–40 секунд.")

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

Сделай 7-8 слайдов. Первый слайд — титульный (только заголовок).
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

        # === Титульный слайд ===
        if slides_data:
            blank_layout = prs.slide_layouts[6]  # Пустой макет
            slide = prs.slides.add_slide(blank_layout)
            set_slide_background(slide, DARK_BG)

            # Заголовок
            title_box = slide.shapes.add_textbox(Inches(0.8), Inches(2.5), Inches(11.7), Inches(1.5))
            tf = title_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = slides_data[0]["title"]
            p.font.size = Pt(40)
            p.font.bold = True
            p.font.color.rgb = WHITE
            p.alignment = PP_ALIGN.CENTER

            # Подзаголовок
            sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(4.3), Inches(11.7), Inches(0.6))
            tf = sub_box.text_frame
            p = tf.paragraphs[0]
            p.text = "Сгенерировано ИИ-ботом"
            p.font.size = Pt(18)
            p.font.color.rgb = ACCENT
            p.alignment = PP_ALIGN.CENTER

        # === Остальные слайды ===
        for slide_info in slides_data[1:]:
            blank_layout = prs.slide_layouts[6]
            slide = prs.slides.add_slide(blank_layout)
            set_slide_background(slide, DARK_BG)

            # Заголовок слайда
            title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.4), Inches(12), Inches(1))
            tf = title_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = slide_info["title"]
            p.font.size = Pt(28)
            p.font.bold = True
            p.font.color.rgb = ACCENT

            # Пункты
            content_box = slide.shapes.add_textbox(Inches(0.9), Inches(1.6), Inches(11.5), Inches(5.2))
            tf = content_box.text_frame
            tf.word_wrap = True

            for i, point in enumerate(slide_info["points"]):
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                p.text = "•  " + point
                p.font.size = Pt(20)
                p.font.color.rgb = WHITE
                p.space_after = Pt(14)
                p.level = 0

        # Сохраняем
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
        await update.message.reply_text(f"Ошибка: {str(e)}")

async def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, generate_presentation))

    await app.initialize()
    await app.start()
    await app.updater.start_polling()

    async def handle(request):
        return web.Response(text="Bot is running")

    web_app = web.Application()
    web_app.router.add_get("/", handle)
    runner = web.AppRunner(web_app)
    await runner.setup()
    port = int(os.environ.get("PORT", 10000))
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()

    await asyncio.Event().wait()

if __name__ == "__main__":
    asyncio.run(main())