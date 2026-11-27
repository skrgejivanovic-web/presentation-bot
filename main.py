import os
import logging
import asyncio
import re
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RgbColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from openai import OpenAI
import tempfile
from aiohttp import web

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

client = OpenAI(api_key=GROQ_API_KEY, base_url="https://api.groq.com/openai/v1")
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Хранилище последних презентаций пользователей (в памяти)
user_data = {}

# Стили
STYLES = {
    "тёмный": {
        "bg": RgbColor(15, 23, 42),
        "accent": RgbColor(56, 189, 248),
        "text": RgbColor(248, 250, 252),
        "sub": RgbColor(148, 163, 184)
    },
    "светлый": {
        "bg": RgbColor(248, 250, 252),
        "accent": RgbColor(2, 132, 199),
        "text": RgbColor(15, 23, 42),
        "sub": RgbColor(100, 116, 139)
    },
    "бизнес": {
        "bg": RgbColor(15, 23, 42),
        "accent": RgbColor(37, 99, 235),
        "text": RgbColor(241, 245, 249),
        "sub": RgbColor(148, 163, 184)
    }
}

def set_background(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_accent_bar(slide, color):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.18), Inches(7.5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()

def create_presentation(slides_data, style_name="тёмный"):
    style = STYLES.get(style_name, STYLES["тёмный"])
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Титульный
    if slides_data:
        slide = prs.slides.add_slide(blank)
        set_background(slide, style["bg"])
        add_accent_bar(slide, style["accent"])

        box = slide.shapes.add_textbox(Inches(1.0), Inches(2.5), Inches(11.3), Inches(2.0))
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = slides_data[0]["title"]
        p.font.size = Pt(40)
        p.font.bold = True
        p.font.color.rgb = style["text"]

        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(4.6), Inches(2.5), Inches(0.08))
        line.fill.solid()
        line.fill.fore_color.rgb = style["accent"]
        line.line.fill.background()

        sub = slide.shapes.add_textbox(Inches(1.0), Inches(5.0), Inches(11), Inches(0.5))
        p = sub.text_frame.paragraphs[0]
        p.text = f"Стиль: {style_name.capitalize()} • Сгенерировано ИИ"
        p.font.size = Pt(14)
        p.font.color.rgb = style["sub"]

    # Контент
    for idx, info in enumerate(slides_data[1:]):
        slide = prs.slides.add_slide(blank)
        set_background(slide, style["bg"])
        add_accent_bar(slide, style["accent"])

        title_box = slide.shapes.add_textbox(Inches(0.9), Inches(0.3), Inches(11.5), Inches(0.8))
        p = title_box.text_frame.paragraphs[0]
        p.text = info["title"]
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = style["accent"]

        body = slide.shapes.add_textbox(Inches(0.9), Inches(1.3), Inches(11.5), Inches(5.5))
        tf = body.text_frame
        tf.word_wrap = True

        for i, point in enumerate(info["points"]):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.text = "●   " + point
            p.font.size = Pt(17)
            p.font.color.rgb = style["text"]
            p.space_after = Pt(12)

    return prs

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! Я создаю презентации с разными стилями и возможностью редактирования.\n\n"
        "Как пользоваться:\n"
        "1. Напиши тему (можно указать стиль):\n"
        "   стиль светлый\n"
        "   Искусственный интеллект\n\n"
        "Доступные стили: тёмный, светлый, бизнес\n\n"
        "2. Чтобы изменить слайд после создания:\n"
        "   измени слайд 3\n"
        "   Новый пункт 1\n"
        "   Новый пункт 2"
    )

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()

    # === РЕДАКТИРОВАНИЕ СЛАЙДА ===
    if text.lower().startswith("измени слайд"):
        if user_id not in user_data:
            await update.message.reply_text("Сначала создай презентацию.")
            return

        match = re.search(r"измени слайд\s+(\d+)", text.lower())
        if not match:
            await update.message.reply_text("Напиши в формате:\nизмени слайд 3\nНовый пункт 1\nНовый пункт 2")
            return

        slide_num = int(match.group(1))
        lines = text.split("\n")[1:]
        new_points = [line.strip() for line in lines if line.strip()]

        data = user_data[user_id]
        if slide_num < 1 or slide_num > len(data["slides"]):
            await update.message.reply_text(f"Слайд {slide_num} не существует. Всего слайдов: {len(data['slides'])}")
            return

        # Обновляем
        if slide_num == 1:
            data["slides"][0]["title"] = new_points[0] if new_points else data["slides"][0]["title"]
        else:
            data["slides"][slide_num - 1]["points"] = new_points if new_points else data["slides"][slide_num - 1]["points"]

        await update.message.reply_text(f"Обновляю слайд {slide_num}...")

        prs = create_presentation(data["slides"], data["style"])
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            prs.save(tmp.name)
            path = tmp.name

        with open(path, "rb") as f:
            await update.message.reply_document(document=f, filename="updated.pptx", caption=f"Слайд {slide_num} обновлён")
        os.remove(path)
        return

    # === СОЗДАНИЕ НОВОЙ ПРЕЗЕНТАЦИИ ===
    style = "тёмный"
    topic = text

    # Проверяем стиль
    for s in STYLES.keys():
        if text.lower().startswith(f"стиль {s}"):
            style = s
            topic = text[len(f"стиль {s}"):].strip()
            break

    if not topic or len(topic) < 3:
        await update.message.reply_text("Напиши тему презентации.")
        return

    await update.message.reply_text(f"Создаю презентацию в стиле «{style}»... 30–50 секунд.")

    try:
        prompt = f"""
Создай подробную структуру презентации на русском на тему: "{topic}".
8–9 слайдов. На каждом 4–6 информативных пунктов.
Первый слайд — только заголовок.

Формат:
СЛАЙД 1: Заголовок
- Пункт 1
- Пункт 2
"""

        response = client.chat.completions.create(
            model="openai/gpt-oss-20b",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.7,
            max_tokens=3000
        )
        content = response.choices[0].message.content

        slides_data = []
        current = None
        for line in content.split("\n"):
            line = line.strip()
            if line.upper().startswith("СЛАЙД"):
                if current:
                    slides_data.append(current)
                title = line.split(":", 1)[-1].strip()
                current = {"title": title, "points": []}
            elif line.startswith("-") and current is not None:
                point = line[1:].strip()
                if point:
                    current["points"].append(point)
        if current:
            slides_data.append(current)

        # Сохраняем для редактирования
        user_data[user_id] = {"slides": slides_data, "style": style, "topic": topic}

        prs = create_presentation(slides_data, style)

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            prs.save(tmp.name)
            path = tmp.name

        with open(path, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=f"{topic[:40]}.pptx",
                caption=f"Стиль: {style}\nЧтобы изменить слайд, напиши:\nизмени слайд 3\nНовый текст"
            )
        os.remove(path)

    except Exception as e:
        logger.error(f"Ошибка: {e}")
        await update.message.reply_text(f"Ошибка: {str(e)}")

async def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

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