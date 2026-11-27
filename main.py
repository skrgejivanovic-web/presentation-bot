import asyncio
import io
import json
import logging
import os
import re

from aiohttp import web
from openai import AsyncOpenAI
from pptx import Presentation
from pptx.util import Pt
from telegram import (
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    InputFile,
    Update,
)
from telegram.constants import ChatAction
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

logging.basicConfig(
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    level=logging.INFO,
)
log = logging.getLogger("presentation-bot")

BOT_TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_TOKEN", "")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")

if OPENAI_API_KEY.startswith("gsk_"):
    DEFAULT_BASE_URL = "https://api.groq.com/openai/v1"
    DEFAULT_MODEL = "openai/gpt-oss-120b"
else:
    DEFAULT_BASE_URL = None
    DEFAULT_MODEL = "gpt-4o-mini"

OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL") or DEFAULT_BASE_URL
OPENAI_MODEL = os.getenv("OPENAI_MODEL") or DEFAULT_MODEL
PORT = int(os.getenv("PORT", "10000"))

COUNTS = (3, 5, 7, 10, 12, 15)

ai = AsyncOpenAI(api_key=OPENAI_API_KEY, base_url=OPENAI_BASE_URL)


class Session:
    def __init__(self):
        self.topic = ""
        self.count = 10
        self.lang = "ru"
        self.busy = False


sessions = {}


def get_session(user_id):
    if user_id not in sessions:
        sessions[user_id] = Session()
    return sessions[user_id]


def mark(selected, text):
    return f"✅ {text}" if selected else text


def settings_keyboard(s):
    rows = []
    for row in (COUNTS[:3], COUNTS[3:]):
        rows.append(
            [
                InlineKeyboardButton(
                    text=mark(s.count == n, str(n)),
                    callback_data=f"n:{n}",
                )
                for n in row
            ]
        )
    rows.append(
        [
            InlineKeyboardButton(
                text=mark(s.lang == "ru", "🇷🇺 Русский"),
                callback_data="lang:ru",
            ),
            InlineKeyboardButton(
                text=mark(s.lang == "en", "🇬🇧 English"),
                callback_data="lang:en",
            ),
        ]
    )
    rows.append(
        [InlineKeyboardButton("🚀 Создать презентацию", callback_data="go")]
    )
    return InlineKeyboardMarkup(rows)


def settings_text(s):
    return (
        f"📌 Тема: {s.topic}\n\n"
        f"Слайдов: {s.count}\n"
        "Выбери количество слайдов и язык, затем жми «Создать».\n"
        "Или просто напиши новую тему сообщением."
    )


async def refresh(q, s):
    try:
        await q.edit_message_text(
            settings_text(s), reply_markup=settings_keyboard(s)
        )
    except Exception:
        pass


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    sessions[update.effective_user.id] = Session()
    await update.message.reply_text(
        "Привет! Я делаю презентации .pptx по любой теме.\n\n"
        "Напиши тему одним сообщением.\n"
        "Например: «Искусственный интеллект в медицине»."
    )


async def on_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    s = get_session(update.effective_user.id)
    text = (update.message.text or "").strip()
    if not text:
        return
    if s.busy:
        await update.message.reply_text(
            "⏳ Подожди, я ещё делаю предыдущую презентацию."
        )
        return
    s.topic = text[:300]
    await update.message.reply_text(
        settings_text(s), reply_markup=settings_keyboard(s)
    )


async def on_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    s = get_session(q.from_user.id)
    data = q.data or ""

    if data.startswith("n:"):
        try:
            s.count = int(data.split(":", 1)[1])
        except ValueError:
            pass
        await q.answer()
        await refresh(q, s)

    elif data.startswith("lang:"):
        s.lang = data.split(":", 1)[1]
        await q.answer()
        await refresh(q, s)

    elif data == "go":
        if s.busy:
            await q.answer("Уже генерирую, подожди 🙂")
            return
        if not s.topic:
            await q.answer("Сначала напиши тему сообщением.")
            return
        await q.answer()
        await q.edit_message_text(
            f"⏳ Генерирую «{s.topic}» ({s.count} слайдов)…\n"
            "Это займёт до минуты."
        )
        await generate_and_send(update, context, s)


PROMPT = (
    "You are a presentation writer. Return ONLY valid JSON, no markdown, "
    'with keys: "title" (string), "subtitle" (string), "slides" '
    '(array of objects with "title" (string) and "bullets" '
    "(array of 3-5 short strings)). "
    "Write in {lang}. Create exactly {count} content slides about: {topic}"
)


async def ask_ai(s):
    lang = "Russian" if s.lang == "ru" else "English"
    prompt = PROMPT.format(lang=lang, count=s.count, topic=s.topic)
    resp = await ai.chat.completions.create(
        model=OPENAI_MODEL,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.7,
    )
    raw = resp.choices[0].message.content or ""
    start, end = raw.find("{"), raw.rfind("}")
    if start == -1 or end == -1:
        raise ValueError("no JSON in AI response")
    content = json.loads(raw[start : end + 1])
    if not isinstance(content.get("slides"), list) or not content["slides"]:
        raise ValueError("bad AI response")
    return content


def build_pptx(content, s):
    prs = Presentation()
    title = str(content.get("title") or s.topic)
    subtitle = str(content.get("subtitle") or "")

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    if subtitle:
        slide.placeholders[1].text = subtitle

    for item in content.get("slides", [])[: s.count]:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = str(item.get("title") or "")
        body = slide.placeholders[1].text_frame
        body.clear()
        bullets = item.get("bullets") or []
        if isinstance(bullets, str):
            bullets = [bullets]
        for i, b in enumerate(bullets[:6]):
            p = body.paragraphs[0] if i == 0 else body.add_paragraph()
            p.text = str(b)
            p.level = 0
            p.font.size = Pt(20)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


async def generate_and_send(update, context, s):
    chat_id = update.effective_chat.id
    status = update.effective_message
    s.busy = True
    try:
        await context.bot.send_chat_action(chat_id, ChatAction.TYPING)
        content = await ask_ai(s)
        data = build_pptx(content, s)
        name = re.sub(r"[^\w\- ]+", "", s.topic).strip()[:40]
        name = name or "presentation"
        await context.bot.send_document(
            chat_id=chat_id,
            document=InputFile(io.BytesIO(data), filename=f"{name}.pptx"),
            caption="✅ Готово! Открывай в PowerPoint или Google Slides.",
        )
        try:
            await status.edit_text("✅ Презентация готова, файл выше 👆")
        except Exception:
            pass
    except Exception:
        log.exception("generation failed")
        try:
            await status.edit_text(
                "😔 Не получилось создать презентацию. "
                "Попробуй ещё раз: /start"
            )
        except Exception:
            pass
    finally:
        s.busy = False


async def health(_request):
    return web.Response(text="OK")


async def main():
    if not BOT_TOKEN or not OPENAI_API_KEY:
        log.error("Set BOT_TOKEN and OPENAI_API_KEY environment variables!")
        return

    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("cancel", cmd_start))
    app.add_handler(CallbackQueryHandler(on_callback))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))

    await app.initialize()
    await app.start()
    await app.updater.start_polling(drop_pending_updates=True)
    log.info("Bot started (polling)")

    webapp = web.Application()
    webapp.router.add_get("/", health)
    webapp.router.add_get("/health", health)
    runner = web.AppRunner(webapp)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", PORT)
    await site.start()
    log.info("Web server listening on port %s", PORT)

    try:
        await asyncio.Event().wait()
    finally:
        await runner.cleanup()
        await app.updater.stop()
        await app.stop()
        await app.shutdown()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        pass