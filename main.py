import os
import logging
import asyncio
import re
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from openai import OpenAI
import tempfile
from aiohttp import web

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

client = OpenAI(api_key=GROQ_API_KEY, base_url="https://api.groq.com/openai/v1")
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

user_data = {}

STYLES = {
    "тёмный": {
        "bg": RGBColor(15, 23, 42),
        "accent": RGBColor(56, 189, 248),
        "title": RGBColor(255, 255, 255),
        "text": RGBColor(226, 232, 240),
        "sub": RGBColor(148, 163, 184)
    },
    "светлый": {
        "bg": RGBColor(248, 250, 252),
        "accent": RGBColor(2, 132, 199),
        "title": RGBColor(15, 23, 42),
        "text": RGBColor(30, 41, 59),
        "sub": RGBColor(100, 116, 139)
    },
    "бизнес": {
        "bg": RGBColor(15, 23, 42),
        "accent": RGBColor(59, 130, 246),
        "title": RGBColor(255, 255, 255),
        "text": RGBColor(226, 232, 240),
        "sub": RGBColor(148, 163, 184)
    }
}

def set_background(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_accent_bar(slide, color):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.15), Inches(7.5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()

def create_presentation(slides_data, style_name="тёмный"):
    style = STYLES.get(style_name, STYLES["тёмный"])
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # ===== ТИТУЛЬНЫЙ СЛАЙД =====
    if slides_data:
        slide = prs.slides.add_slide(blank)
        set_background(slide, style["bg"])
        add_accent_bar(slide, style["accent"])

        # Большой заголовок
        title_box = slide.shapes.add_textbox(Inches(0.9), Inches(2.2), Inches(11.5), Inches(2.4))
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = slides_data[0]["title"]
        p.font.size = Pt(44)
        p.font.bold = True
        p.font.color.rgb = style["title"]
        p.alignment = PP_ALIGN.LEFT

        # Акцентная линия
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.9), Inches(4.7), Inches(3.0), Inches(0.09))
        line.fill.solid()
        line.fill.fore_color.rgb = style["accent"]
        line.line.fill.background()

        # Подпись
        sub = slide.shapes.add_textbox(Inches(0.9), Inches(5.1), Inches(11), Inches(0.5))
        p = sub.text_frame.paragraphs[0]
        p.text = f"Стиль: {style_name.capitalize()}  •  Сгенерировано ИИ"
        p.font.size = Pt(15)
        p.font.color.rgb = style["sub"]

    # ===== ОБЫЧНЫЕ СЛАЙДЫ =====
    for info in slides_data[1:]:
        slide = prs.slides.add_slide(blank)
        set_background(slide, style["bg"])
        add_accent_bar(slide, style["accent"])

        # Заголовок слайда
        title_box = slide.shapes.add_textbox(Inches(0.9), Inches(0.35), Inches(11.5), Inches(0.9))
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = info["title"]
        p.font.size = Pt(26)
        p.font.bold = True
        p.font.color.rgb = style["accent"]

        # Текст
        body = slide.shapes.add_textbox(Inches(0.9), Inches(1.4), Inches(11.5), Inches(5.4))
        tf = body.text_frame
        tf.word_wrap = True

        for i, point in enumerate(info["points"]):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.text = "●   " + point
            p.font.size = Pt(18)
            p.font.color.rgb = style["text"]
            p.space_after = Pt(14)

    return prs

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! Просто напиши тему презентации.\n\n"
        "Можно указать стиль:\n"
        "стиль светлый\n"
        "Искусственный интеллект\n\n"
        "Стили: тёмный, светлый, бизнес\n\n"
        "После создания можно изменить слайд:\n"
        "измени слайд 3\n"
        "Новый пункт 1\n"
        "Новый пункт 2"
    )

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()

    # Редактирование
    if text.lower().startswith("измени слайд"):
        if user_id not in user_data:
            await update.message.reply_text("Сначала создай презентацию.")
            return

        match = re.search(r"измени слайд\s+(\d+)", text.lower())
        if not match:
            await update.message.reply_text("Формат:\nизмени слайд 3\nНовый пункт 1\nНовый пункт 2")
            return

        slide_num = int(match.group(1))
        new_points = [line.strip() for line in text.split("\n")[1:] if line.strip()]

        data = user_data[user_id]
        if slide_num < 1 or slide_num > len(data["slides"]):
            await update.message.reply_text(f"Слайд {slide_num} не существует.")
            return

        if slide_num == 1:
            data["slides"][0]["title"] = new_points[0] if new_points else data["slides"][0]["title"]
        else:
            data["slides"][slide_num-1]["points"] = new_points or data["slides"][slide_num-1]["points"]

        await update.message.reply_text(f"Обновляю слайд {slide_num}...")
        prs = create_presentation(data["slides"], data["style"])

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            prs.save(tmp.name)
            path = tmp.name

        with open(path, "rb") as f:
            await update.message.reply_document(document=f, filename="updated.pptx", caption=f"Слайд {slide_num} обновлён")
        os.remove(path)
        return

    # Создание новой презентации
    style = "тёмный"
    topic = text

    for s in STYLES:
        if text.lower().startswith(f"стиль {s}"):
            style = s
            topic = text[len(f"стиль {s}"):].strip()
            break

    if len(topic) < 3:
        await update.message.reply_text("Напиши тему презентации.")
        return

    await update.message.reply_text(f"Создаю презентацию («{style}» стиль)...")

    try:
        prompt = f"""Создай подробную структуру презентации на русском на тему: "{topic}".
Сделай 8 слайдов. На каждом слайде 4-6 информативных пунктов.
Первый слайд — только заголовок темы.

Формат строго:
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

        user_data[user_id] = {"slides": slides_data, "style": style}

        prs = create_presentation(slides_data, style)

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            prs.save(tmp.name)
            path = tmp.name

        with open(path, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=f"{topic[:40]}.pptx",
                caption=f"Готово! Стиль: {style}\n\nЧтобы изменить слайд:\nизмени слайд 3\nНовый текст"
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
    site = web.TCPSite(runner, "0.0.0.0", int(os.environ.get("PORT", 10000)))
    await site.start()

    await asyncio.Event().wait()

if __name__ == "__main__":
    asyncio.run(main())