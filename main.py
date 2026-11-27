import asyncio
import io
import json
import logging
import os
import re
import time
from datetime import datetime

import aiohttp
from aiohttp import web
from openai import AsyncOpenAI
from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt
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

VER = "10"

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

MODEL_CHAIN = []
for m in (
    OPENAI_MODEL,
    "openai/gpt-oss-120b",
    "openai/gpt-oss-20b",
    "qwen/qwen3.8-27b",
):
    if m and m not in MODEL_CHAIN:
        MODEL_CHAIN.append(m)

COUNTS = (3, 5, 7, 10, 12, 15)

COMMONS_API = "https://commons.wikimedia.org/w/api.php"
WIKI_API = "https://ru.wikipedia.org/w/api.php"
OPENVERSE_API = "https://api.openverse.org/v1/images/"
HTTP_HEADERS = {"User-Agent": "presentation-bot/1.0 (telegram bot)"}

PALETTE_NAMES = {
    "dark": "🌙 тёмная", "light": "☀️ светлая", "color": "🎨 яркая",
    "med": "💊 медицинская", "space": "🚀 космическая",
    "nature": "🌿 природная", "finance": "💰 деловая",
    "tech": "🤖 техно", "edu": "🎓 учебная",
    "sport": "🏆 спортивная", "art": "🎭 творческая",
    "food": "🍳 кулинарная", "travel": "✈️ путешествия",
    "history": "📜 история", "science": "🔬 наука",
    "music": "🎵 музыка", "gaming": "🎮 игры",
    "fashion": "👗 мода", "auto": "🚗 авто",
    "real_estate": "🏠 недвижимость", "psychology": "🧠 психология",
}

PALETTES = {
    "dark": {"bg": RGBColor(0x17, 0x17, 0x20), "panel": RGBColor(0x26, 0x26, 0x33), "accent": RGBColor(0x00, 0xB8, 0xD4), "accent2": RGBColor(0x7C, 0x4D, 0xFF), "text": RGBColor(0xF2, 0xF2, 0xF5), "muted": RGBColor(0x9A, 0x9A, 0xB0)},
    "light": {"bg": RGBColor(0xFA, 0xFA, 0xFC), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0x1F, 0x6F, 0xEB), "accent2": RGBColor(0x0F, 0x9D, 0x58), "text": RGBColor(0x1E, 0x1E, 0x28), "muted": RGBColor(0x6B, 0x6B, 0x7A)},
    "color": {"bg": RGBColor(0xFD, 0xF6, 0xEC), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0xFF, 0x6B, 0x35), "accent2": RGBColor(0x2E, 0x86, 0xAB), "text": RGBColor(0x2B, 0x2B, 0x2B), "muted": RGBColor(0x8A, 0x7F, 0x70)},
    "med": {"bg": RGBColor(0xF2, 0xFA, 0xF8), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0x00, 0x96, 0x88), "accent2": RGBColor(0x00, 0x69, 0x5C), "text": RGBColor(0x16, 0x38, 0x32), "muted": RGBColor(0x5F, 0x7A, 0x74)},
    "space": {"bg": RGBColor(0x0B, 0x10, 0x26), "panel": RGBColor(0x16, 0x1D, 0x3A), "accent": RGBColor(0x8A, 0xB4, 0xFF), "accent2": RGBColor(0x7C, 0x4D, 0xFF), "text": RGBColor(0xEE, 0xF0, 0xFF), "muted": RGBColor(0x8E, 0x96, 0xC0)},
    "nature": {"bg": RGBColor(0xF3, 0xF8, 0xEE), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0x55, 0x8B, 0x2F), "accent2": RGBColor(0x33, 0x69, 0x1E), "text": RGBColor(0x26, 0x30, 0x1C), "muted": RGBColor(0x6E, 0x7B, 0x62)},
    "finance": {"bg": RGBColor(0x14, 0x18, 0x1D), "panel": RGBColor(0x1E, 0x24, 0x2B), "accent": RGBColor(0xD4, 0xAF, 0x37), "accent2": RGBColor(0x2E, 0x7D, 0x32), "text": RGBColor(0xF0, 0xED, 0xE6), "muted": RGBColor(0x9A, 0xA3, 0xAD)},
    "tech": {"bg": RGBColor(0x10, 0x14, 0x18), "panel": RGBColor(0x1A, 0x20, 0x26), "accent": RGBColor(0x00, 0xE5, 0xFF), "accent2": RGBColor(0x7C, 0x4D, 0xFF), "text": RGBColor(0xE8, 0xF6, 0xF8), "muted": RGBColor(0x8F, 0xA3, 0xAD)},
    "edu": {"bg": RGBColor(0xFD, 0xFD, 0xF7), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0x1E, 0x88, 0xE5), "accent2": RGBColor(0xFB, 0x8C, 0x00), "text": RGBColor(0x21, 0x21, 0x21), "muted": RGBColor(0x75, 0x75, 0x75)},
    "sport": {"bg": RGBColor(0x1A, 0x1A, 0x1A), "panel": RGBColor(0x26, 0x26, 0x26), "accent": RGBColor(0xFF, 0x6D, 0x00), "accent2": RGBColor(0x00, 0xC8, 0x53), "text": RGBColor(0xFF, 0xFF, 0xFF), "muted": RGBColor(0xA6, 0xA6, 0xA6)},
    "art": {"bg": RGBColor(0xFF, 0xF7, 0xF0), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0xE9, 0x1E, 0x63), "accent2": RGBColor(0x6A, 0x1B, 0x9A), "text": RGBColor(0x3A, 0x24, 0x30), "muted": RGBColor(0x9C, 0x85, 0x89)},
    "food": {"bg": RGBColor(0xFF, 0xF8, 0xF0), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0xE8, 0x5D, 0x04), "accent2": RGBColor(0xBC, 0x47, 0x49), "text": RGBColor(0x3D, 0x40, 0x5B), "muted": RGBColor(0x8D, 0x99, 0xAE)},
    "travel": {"bg": RGBColor(0xE0, 0xF7, 0xFA), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0x00, 0x97, 0xA7), "accent2": RGBColor(0x00, 0x60, 0x64), "text": RGBColor(0x26, 0x32, 0x38), "muted": RGBColor(0x78, 0x90, 0x9C)},
    "history": {"bg": RGBColor(0x3E, 0x27, 0x23), "panel": RGBColor(0x4E, 0x34, 0x2E), "accent": RGBColor(0xFF, 0xD5, 0x4F), "accent2": RGBColor(0xB7, 0x1C, 0x1C), "text": RGBColor(0xFF, 0xEC, 0xB3), "muted": RGBColor(0xBC, 0xAA, 0xA4)},
    "science": {"bg": RGBColor(0x1A, 0x23, 0x7E), "panel": RGBColor(0x28, 0x35, 0x93), "accent": RGBColor(0x7C, 0x4D, 0xFF), "accent2": RGBColor(0x53, 0x6D, 0xFE), "text": RGBColor(0xE8, 0xEA, 0xF6), "muted": RGBColor(0x9F, 0xA8, 0xDA)},
    "music": {"bg": RGBColor(0x4A, 0x14, 0x8C), "panel": RGBColor(0x6A, 0x1B, 0x9A), "accent": RGBColor(0xE0, 0x40, 0xFB), "accent2": RGBColor(0xAA, 0x00, 0xFF), "text": RGBColor(0xF3, 0xE5, 0xF5), "muted": RGBColor(0xCE, 0x93, 0xD8)},
    "gaming": {"bg": RGBColor(0x0D, 0x11, 0x17), "panel": RGBColor(0x16, 0x1B, 0x22), "accent": RGBColor(0x39, 0xD3, 0x53), "accent2": RGBColor(0x00, 0xFF, 0x88), "text": RGBColor(0xC9, 0xD1, 0xD9), "muted": RGBColor(0x8B, 0x94, 0x9E)},
    "fashion": {"bg": RGBColor(0xFA, 0xFA, 0xFA), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0x00, 0x00, 0x00), "accent2": RGBColor(0x33, 0x33, 0x33), "text": RGBColor(0x1A, 0x1A, 0x1A), "muted": RGBColor(0x66, 0x66, 0x66)},
    "auto": {"bg": RGBColor(0x21, 0x21, 0x21), "panel": RGBColor(0x30, 0x30, 0x30), "accent": RGBColor(0xFF, 0x17, 0x44), "accent2": RGBColor(0xC5, 0x11, 0x62), "text": RGBColor(0xF5, 0xF5, 0xF5), "muted": RGBColor(0x9E, 0x9E, 0x9E)},
    "real_estate": {"bg": RGBColor(0xF1, 0xF8, 0xE9), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0x55, 0x8B, 0x2F), "accent2": RGBColor(0x33, 0x69, 0x1E), "text": RGBColor(0x1B, 0x5E, 0x20), "muted": RGBColor(0x68, 0x9F, 0x38)},
    "psychology": {"bg": RGBColor(0xE3, 0xF2, 0xFD), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0x19, 0x76, 0xD2), "accent2": RGBColor(0x0D, 0x47, 0xA1), "text": RGBColor(0x0D, 0x47, 0xA1), "muted": RGBColor(0x64, 0xB5, 0xF6)},
}

THEME_FONTS = {
    "dark": ("Trebuchet MS", "Calibri"),
    "light": ("Georgia", "Calibri"),
    "color": ("Verdana", "Trebuchet MS"),
    "med": ("Georgia", "Calibri"),
    "space": ("Consolas", "Trebuchet MS"),
    "nature": ("Palatino Linotype", "Georgia"),
    "finance": ("Palatino Linotype", "Calibri"),
    "tech": ("Consolas", "Calibri"),
    "edu": ("Verdana", "Calibri"),
    "sport": ("Impact", "Arial"),
    "art": ("Palatino Linotype", "Georgia"),
    "food": ("Georgia", "Calibri"),
    "travel": ("Trebuchet MS", "Calibri"),
    "history": ("Palatino Linotype", "Georgia"),
    "science": ("Consolas", "Calibri"),
    "music": ("Georgia", "Calibri"),
    "gaming": ("Consolas", "Arial"),
    "fashion": ("Georgia", "Calibri"),
    "auto": ("Impact", "Arial"),
    "real_estate": ("Georgia", "Calibri"),
    "psychology": ("Georgia", "Calibri"),
}

THEME_KEYWORDS = (
    (("медиц", "здоров", "врач", "больн", "лекар", "health", "medic", "doctor", "hospital"), "med"),
    (("косм", "звезд", "планет", "вселен", "галакт", "space", "planet", "galaxy", "universe", "astronom"), "space"),
    (("природ", "экол", "лес", "растен", "животн", "климат", "nature", "ecolog", "forest", "environment"), "nature"),
    (("финанс", "деньг", "банк", "инвест", "эконом", "бизнес", "прибыл", "finance", "money", "bank", "invest", "business"), "finance"),
    (("техно", "робот", "программ", "нейросет", "интеллект", "компьютер", "интернет", "цифров", "tech", "robot", "software", "computer", "digital"), "tech"),
    (("школ", "учеб", "образован", "курс", "студент", "преподав", "education", "school", "learn", "study"), "edu"),
    (("спорт", "футбол", "фитнес", "олимп", "хокке", "баскетбол", "sport", "fitness", "soccer", "football"), "sport"),
    (("искусств", "культур", "музык", "театр", "дизайн", "живопис", "artist", "culture", "music", "theater", "design", "paint"), "art"),
    (("еда", "кулинар", "рецепт", "готов", "блюдо", "ужин", "обед", "завтрак", "food", "cook", "recipe", "dinner", "lunch", "breakfast"), "food"),
    (("путешеств", "туризм", "отпуск", "стран", "travel", "tourism", "vacation", "country"), "travel"),
    (("истор", "древн", "войн", "импер", "цар", "history", "ancient", "war", "empire"), "history"),
    (("наук", "физик", "хим", "биолог", "математ", "science", "physics", "chemistry", "biology", "math"), "science"),
    (("игр", "геймер", "видеоигр", "gaming", "game", "gamer", "video"), "gaming"),
    (("мод", "стиль", "одежд", "fashion", "clothing"), "fashion"),
    (("авто", "машин", "автомобил", "транспорт", "auto", "car", "vehicle"), "auto"),
    (("недвижим", "дом", "квартир", "аренд", "real estate", "house", "apartment", "rent"), "real_estate"),
    (("психолог", "мозг", "сознан", "эмоц", "psychology", "brain", "mind", "emotion"), "psychology"),
)

WHITE = RGBColor(0xFF, 0xFF, 0xFF)

ai = AsyncOpenAI(api_key=OPENAI_API_KEY, base_url=OPENAI_BASE_URL)


class Session:
    def __init__(self):
        self.topic = ""
        self.count = 10
        self.lang = "ru"
        self.style = "auto"
        self.busy = False
        self.last_content = None
        self.last_imgs = None
        self.last_title_img = None
        self.editing_slide = None


sessions = {}


def get_session(user_id):
    if user_id not in sessions:
        sessions[user_id] = Session()
    return sessions[user_id]


def auto_style(topic):
    low = (topic or "").lower()
    for words, key in THEME_KEYWORDS:
        for w in words:
            if w in low:
                return key
    pool = ("dark", "light", "color", "art", "space", "tech")
    return pool[sum(ord(ch) for ch in low) % len(pool)]


def resolve_style(s):
    if s.style == "auto":
        return auto_style(s.topic)
    return s.style


def style_label(s):
    if s.style == "auto":
        key = auto_style(s.topic)
        return f"🪄 авто → {PALETTE_NAMES.get(key, key)}"
    return PALETTE_NAMES.get(s.style, s.style)


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
        [
            InlineKeyboardButton(
                text=mark(s.style == "auto", "🪄 Авто"),
                callback_data="style:auto",
            ),
            InlineKeyboardButton(
                text=mark(s.style == "dark", "🌙 Тёмная"),
                callback_data="style:dark",
            ),
            InlineKeyboardButton(
                text=mark(s.style == "light", "☀️ Светлая"),
                callback_data="style:light",
            ),
            InlineKeyboardButton(
                text=mark(s.style == "color", "🎨 Яркая"),
                callback_data="style:color",
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
        f"Слайдов: {s.count}   Оформление: {style_label(s)}\n"
        "Настрой и жми «Создать». Или напиши новую тему сообщением."
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
        "Привет! Я делаю красивые презентации .pptx: дизайн под тему, "
        "картинки, диаграммы, факты и цитаты.\n\n"
        "Напиши тему одним сообщением.\n"
        "Например: «Искусственный интеллект в медицине».\n\n"
        f"(версия {VER})"
    )


async def cmd_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    s = get_session(update.effective_user.id)
    s.editing_slide = None
    await update.message.reply_text("❌ Отменено")


async def on_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    s = get_session(update.effective_user.id)
    text = (update.message.text or "").strip()
    if not text:
        return
    if s.editing_slide:
        if s.last_content and 1 <= s.editing_slide <= len(
            s.last_content.get("slides", [])
        ):
            s.last_content["slides"][s.editing_slide - 1]["title"] = text[:120]
            s.editing_slide = None
            await rebuild_and_send(update, context, s)
        else:
            s.editing_slide = None
            await update.message.reply_text("❌ Слайд не найден, отменено")
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

    elif data.startswith("style:"):
        s.style = data.split(":", 1)[1]
        await q.answer()
        await refresh(q, s)

    elif data.startswith("edit:"):
        if not s.last_content:
            await q.answer("Сначала создай презентацию")
            return
        try:
            s.editing_slide = int(data.split(":", 1)[1])
        except ValueError:
            await q.answer("Ошибка")
            return
        await q.answer()
        await q.message.reply_text(
            f"✏️ Напиши новый заголовок для слайда {s.editing_slide} "
            "(или /cancel)"
        )

    elif data == "go":
        if s.busy:
            await q.answer("Уже генерирую, подожди 🙂")
            return
        if not s.topic:
            await q.answer("Сначала напиши тему сообщением.")
            return
        await q.answer()
        await q.edit_message_text(
            f"⏳ Рисую «{s.topic}» ({s.count} слайдов, "
            f"{style_label(s)})…"
        )
        await generate_and_send(update, context, s)


PROMPT = (
    "You are a top-tier presentation designer and writer. "
    "Return ONLY valid JSON, no markdown, no explanations, "
    'with keys: "title" (string), "subtitle" (string), '
    '"img" (3-5 ENGLISH words describing a photo for the title slide), '
    '"intro" (1-2 catchy sentences), '
    '"slides" (array of exactly {count} objects, each with: "title" '
    '(string), "lead" (one engaging sentence), '
    '"img" (3-5 ENGLISH words with the main subject first), '
    '"bullets" (array of 5-6 NON-EMPTY informative strings, each 8-18 '
    'words), '
    '"fact" (one surprising fact or witty remark), '
    '"stats" (array of 3 objects with "value" (short number or percent) '
    'and "label" (short description)), '
    '"chart" (object with "title" (string), "type" ("pie" or "bar"), '
    '"data" (array of 4-6 objects with "label" (short string) and '
    '"value" (number; percents 0-100 when type is pie))), '
    '"conclusion" (2-3 inspiring sentences). '
    "Write EVERY string strictly in {lang}, never mix languages. "
    "Topic: {topic}"
)


def extract_json(text):
    text = text.strip()
    code_match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", text)
    if code_match:
        text = code_match.group(1)
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1 or start >= end:
        return None
    return text[start : end + 1]


async def ask_ai(s):
    lang = "Russian" if s.lang == "ru" else "English"
    prompt = PROMPT.format(lang=lang, count=s.count, topic=s.topic)
    messages = [
        {
            "role": "system",
            "content": "You are a presentation writer. You reply only "
                       "with valid JSON, without markdown and comments.",
        },
        {"role": "user", "content": prompt},
    ]
    last_error = None
    for model in MODEL_CHAIN:
        for use_format in (True, False):
            try:
                kwargs = {"model": model, "messages": messages}
                if use_format:
                    kwargs["response_format"] = {"type": "json_object"}
                resp = await ai.chat.completions.create(**kwargs)
                raw = resp.choices[0].message.content or ""
                json_str = extract_json(raw)
                if not json_str:
                    last_error = ValueError(f"{model}: no JSON")
                    log.warning(str(last_error))
                    continue
                content = json.loads(json_str)
                if not isinstance(content.get("slides"), list) or \
                        not content["slides"]:
                    last_error = ValueError(f"{model}: bad slides")
                    log.warning(str(last_error))
                    continue
                log.info("AI model used: %s (json_mode=%s)",
                         model, use_format)
                return content
            except Exception as e:
                last_error = e
                log.warning("Model %s json_mode=%s failed: %s",
                            model, use_format, e)
    raise last_error


async def fallback_content(s):
    params = {
        "action": "query",
        "format": "json",
        "generator": "search",
        "gsrsearch": s.topic,
        "gsrlimit": "8",
        "gsrnamespace": "0",
        "prop": "extracts",
        "exintro": "1",
        "explaintext": "1",
    }
    timeout = aiohttp.ClientTimeout(total=15)
    async with aiohttp.ClientSession(timeout=timeout,
                                     headers=HTTP_HEADERS) as sess:
        async with sess.get(WIKI_API, params=params) as r:
            data = await r.json(content_type=None)
    pages = (data.get("query") or {}).get("pages") or {}
    ordered = sorted(pages.values(), key=lambda p: p.get("index", 99))
    slides = []
    for page in ordered:
        extract = (page.get("extract") or "").strip()
        if not extract:
            continue
        sentences = [x.strip() for x in re.split(r"(?<=[.!?])\s+", extract)
                     if x.strip()]
        bullets = [x for x in sentences if len(x) > 20][:5]
        if not bullets:
            bullets = [extract[:300]]
        slides.append({
            "title": page.get("title", s.topic),
            "lead": sentences[0] if sentences else "",
            "img": page.get("title", s.topic),
            "bullets": bullets,
            "fact": "",
        })
        if len(slides) >= s.count:
            break
    if not slides:
        raise ValueError("no Wikipedia data for topic")
    return {
        "title": s.topic,
        "subtitle": "Материалы собраны из Википедии",
        "img": s.topic,
        "intro": "",
        "slides": slides,
        "stats": [],
        "chart": None,
        "conclusion": "Изучайте новое каждый день!",
    }


def prepare_image(raw, w=900, h=600):
    img = Image.open(io.BytesIO(raw))
    img = img.convert("RGB")
    ar = w / h
    iw, ih = img.size
    if iw / ih > ar:
        new_w = int(ih * ar)
        left = (iw - new_w) // 2
        img = img.crop((left, 0, left + new_w, ih))
    else:
        new_h = int(iw / ar)
        top = (ih - new_h) // 2
        img = img.crop((0, top, iw, top + new_h))
    img = img.resize((w, h))
    out = io.BytesIO()
    img.save(out, format="JPEG", quality=85)
    return out.getvalue()


def _tokens(q):
    return set(re.findall(r"[a-z0-9а-яё]{4,}", (q or "").lower()))


def _relevant(query, title):
    qw = _tokens(query)
    tw = _tokens(title)
    if not qw or not tw:
        return True
    return bool(qw & tw)


async def _download_image(sess, url):
    try:
        async with sess.get(url) as r:
            if r.status != 200:
                return None
            raw = await r.read()
        if not raw:
            return None
        return prepare_image(raw)
    except Exception as e:
        log.warning("image download failed %s: %s", url, e)
        return None


async def fetch_topic_image(query):
    timeout = aiohttp.ClientTimeout(total=5, connect=3)
    async with aiohttp.ClientSession(timeout=timeout,
                                     headers=HTTP_HEADERS) as sess:
        try:
            params = {
                "action": "query",
                "format": "json",
                "generator": "search",
                "gsrnamespace": "6",
                "gsrsearch": f"filetype:bitmap {query}",
                "gsrlimit": "5",
                "prop": "imageinfo",
                "iiprop": "url|mime",
                "iiurlwidth": "900",
            }
            async with sess.get(COMMONS_API, params=params) as r:
                data = await r.json(content_type=None)
            pages = (data.get("query") or {}).get("pages") or {}
            for page in pages.values():
                if not _relevant(query, page.get("title")):
                    continue
                info = (page.get("imageinfo") or [{}])[0]
                url = info.get("thumburl") or info.get("url")
                mime = info.get("mime") or ""
                if url and mime in ("image/jpeg", "image/png"):
                    img = await _download_image(sess, url)
                    if img:
                        log.info("image from commons: %s", query)
                        return img
        except Exception as e:
            log.warning("commons failed for %s: %s", query, e)
        try:
            params = {
                "action": "query",
                "format": "json",
                "generator": "search",
                "gsrsearch": query,
                "gsrlimit": "3",
                "gsrnamespace": "0",
                "prop": "pageimages",
                "piprop": "thumbnail",
                "pithumbsize": "900",
            }
            async with sess.get(WIKI_API, params=params) as r:
                data = await r.json(content_type=None)
            pages = (data.get("query") or {}).get("pages") or {}
            for page in pages.values():
                url = (page.get("thumbnail") or {}).get("source") or ""
                if url:
                    img = await _download_image(sess, url)
                    if img:
                        log.info("image from wikipedia: %s", query)
                        return img
        except Exception as e:
            log.warning("wiki image failed for %s: %s", query, e)
        try:
            async with sess.get(OPENVERSE_API,
                                params={"q": query, "page_size": "5"}) as r:
                if r.status == 200:
                    data = await r.json(content_type=None)
                    for res in (data.get("results") or []):
                        if not _relevant(query, res.get("title")):
                            continue
                        url = res.get("url") or ""
                        if url:
                            img = await _download_image(sess, url)
                            if img:
                                log.info("image from openverse: %s", query)
                                return img
        except Exception as e:
            log.warning("openverse failed for %s: %s", query, e)
    return None


async def fetch_images_parallel(queries, limit=4):
    tasks = []
    for i, q in enumerate(queries):
        if i >= limit or not q:
            continue
        tasks.append(fetch_topic_image(q))
    
    if not tasks:
        return [None] * len(queries)
    
    results = await asyncio.gather(*tasks, return_exceptions=True)
    
    out = []
    for r in results:
        if isinstance(r, BaseException):
            out.append(None)
        else:
            out.append(r)
    
    while len(out) < len(queries):
        out.append(None)
    
    return out


def set_bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def add_rect(slide, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE):
    sp = slide.shapes.add_shape(shape, x, y, w, h)
    sp.fill.solid()
    sp.fill.fore_color.rgb = color
    sp.line.fill.background()
    sp.shadow.inherit = False
    return sp


def add_text(slide, x, y, w, h, text, size, color, bold=False,
             italic=False, align=PP_ALIGN.LEFT, font="Calibri"):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = align
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.italic = italic
    p.font.color.rgb = color
    p.font.name = font
    return tb


def add_numbered_title(slide, x, y, w, h, number, text, size, pal):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r1 = p.add_run()
    r1.text = f"{number}  "
    r1.font.size = Pt(size)
    r1.font.bold = True
    r1.font.color.rgb = pal["accent"]
    r1.font.name = pal["hfont"]
    r2 = p.add_run()
    r2.text = text
    r2.font.size = Pt(size)
    r2.font.bold = True
    r2.font.color.rgb = pal["text"]
    r2.font.name = pal["hfont"]
    return tb


def add_picture(slide, img, x, y, w, h, pal):
    add_rect(slide, x - Inches(0.07), y - Inches(0.07),
             w + Inches(0.14), h + Inches(0.14), pal["accent"])
    slide.shapes.add_picture(io.BytesIO(img), x, y, width=w, height=h)


def clean_bullets(raw, limit=5):
    out = []
    for b in raw or []:
        t = str(b).strip()
        t = t.strip("\"'""«»`)")
        if len(t) < 3 or t in (",", "."):
            continue
        out.append(t[:240])
        if len(out) >= limit:
            break
    return out


def add_bullets(slide, x, y, w, h, bullets, pal, size=20):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(14)
        r1 = p.add_run()
        r1.text = "▸ "
        r1.font.size = Pt(size)
        r1.font.bold = True
        r1.font.color.rgb = pal["accent"]
        r1.font.name = pal["bfont"]
        r2 = p.add_run()
        r2.text = str(b)
        r2.font.size = Pt(size)
        r2.font.color.rgb = pal["text"]
        r2.font.name = pal["bfont"]
    return tb


def chart_rows(chart):
    rows = []
    for d in (chart or {}).get("data") or []:
        if not isinstance(d, dict):
            continue
        label = str(d.get("label", "")).strip()[:40]
        try:
            value = float(d.get("value", 0))
        except (TypeError, ValueError):
            continue
        if label:
            rows.append((label, value))
        if len(rows) >= 6:
            break
    return rows if len(rows) >= 2 else []


def add_title_slide(prs, content, pal, s, img):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["bg"])
    add_rect(slide, Inches(-1.6), Inches(4.6), Inches(5.2), Inches(5.2),
             pal["accent2"], MSO_SHAPE.OVAL)
    add_rect(slide, Inches(10.4), Inches(-2.2), Inches(5.0), Inches(5.0),
             pal["accent"], MSO_SHAPE.OVAL)
    tw = Inches(7.4) if img else Inches(11.3)
    add_text(slide, Inches(1.0), Inches(1.9), tw, Inches(1.9),
             str(content.get("title") or s.topic), 52, pal["text"],
             bold=True, font=pal["hfont"])
    add_rect(slide, Inches(1.05), Inches(3.85), Inches(3.2), Inches(0.09),
             pal["accent"])
    add_text(slide, Inches(1.05), Inches(4.1), tw, Inches(0.8),
             str(content.get("subtitle") or ""), 24, pal["muted"],
             font=pal["bfont"])
    intro = str(content.get("intro") or "")
    if intro:
        add_text(slide, Inches(1.05), Inches(5.0), tw, Inches(1.1),
                 intro, 20, pal["accent"], italic=True, font=pal["bfont"])
    if img:
        add_picture(slide, img, Inches(8.8), Inches(2.2),
                    Inches(3.8), Inches(2.9), pal)


def add_agenda_slide(prs, titles, pal):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["bg"])
    add_rect(slide, 0, 0, Inches(0.18), prs.slide_height, pal["accent2"])
    add_rect(slide, 0, 0, prs.slide_width, Inches(0.12), pal["accent"])
    add_text(slide, Inches(0.6), Inches(0.4), Inches(12.1), Inches(0.9),
             "Что внутри", 40, pal["text"], bold=True, font=pal["hfont"])
    add_rect(slide, Inches(0.62), Inches(1.45), Inches(2.6), Inches(0.08),
             pal["accent"])
    tb = slide.shapes.add_textbox(Inches(0.65), Inches(1.9),
                                  Inches(12.0), Inches(5.1))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, t in enumerate(titles[:8]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(16)
        r1 = p.add_run()
        r1.text = f"{i + 1:02d}  "
        r1.font.size = Pt(22)
        r1.font.bold = True
        r1.font.color.rgb = pal["accent"]
        r1.font.name = pal["hfont"]
        r2 = p.add_run()
        r2.text = str(t)
        r2.font.size = Pt(22)
        r2.font.color.rgb = pal["text"]
        r2.font.name = pal["bfont"]


def add_content_slide(prs, item, pal, num, total, img):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["bg"])
    add_rect(slide, 0, 0, Inches(0.18), prs.slide_height, pal["accent2"])
    add_rect(slide, 0, 0, prs.slide_width, Inches(0.12), pal["accent"])
    add_numbered_title(slide, Inches(0.6), Inches(0.35), Inches(12.1),
                       Inches(1.0), f"{num - 2:02d}",
                       str(item.get("title") or ""), 40, pal)
    add_rect(slide, Inches(0.62), Inches(1.42), Inches(2.6), Inches(0.08),
             pal["accent"])
    lead = str(item.get("lead") or "")
    if lead:
        add_text(slide, Inches(0.62), Inches(1.62), Inches(12.1),
                 Inches(0.6), lead, 19, pal["muted"], italic=True,
                 font=pal["bfont"])
    bullets = clean_bullets(item.get("bullets"), 5)
    bsize = 20 if (img or len(bullets) >= 4) else 24
    wide = Inches(12.0)
    if img:
        wide = Inches(8.2)
        bullets = bullets[:4]
        add_picture(slide, img, Inches(9.3), Inches(2.25),
                    Inches(3.45), Inches(2.6), pal)
    add_bullets(slide, Inches(0.65), Inches(2.3), wide,
                Inches(3.6), bullets, pal, size=bsize)
    fact = str(item.get("fact") or "")
    if fact:
        add_rect(slide, Inches(0.6), Inches(6.1), Inches(12.1), Inches(1.0),
                 pal["panel"], MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(slide, Inches(0.6), Inches(6.1), Inches(0.1), Inches(1.0),
                 pal["accent"])
        add_text(slide, Inches(0.95), Inches(6.3), Inches(11.5),
                 Inches(0.65), "💡 " + fact, 17, pal["text"], italic=True,
                 font=pal["bfont"])
    add_text(slide, Inches(11.9), Inches(7.12), Inches(1.3), Inches(0.35),
             f"{num} / {total}", 12, pal["muted"], align=PP_ALIGN.RIGHT,
             font=pal["bfont"])


def add_stats_slide(prs, stats, pal, num, total):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["bg"])
    add_rect(slide, 0, 0, Inches(0.18), prs.slide_height, pal["accent2"])
    add_rect(slide, 0, 0, prs.slide_width, Inches(0.12), pal["accent"])
    add_text(slide, Inches(0.6), Inches(0.4), Inches(12.1), Inches(0.9),
             "Цифры и факты", 40, pal["text"], bold=True, font=pal["hfont"])
    add_rect(slide, Inches(0.62), Inches(1.45), Inches(2.6), Inches(0.08),
             pal["accent"])
    n = min(len(stats), 3)
    card_w = Inches(3.7)
    gap = Inches(0.45)
    total_w = n * card_w + (n - 1) * gap
    start = int((prs.slide_width - total_w) / 2)
    for i, st in enumerate(stats[:3]):
        x = start + i * (card_w + gap)
        add_rect(slide, x, Inches(2.1), card_w, Inches(3.0), pal["panel"],
                 MSO_SHAPE.ROUNDED_RECTANGLE)
        if isinstance(st, dict):
            value = str(st.get("value", ""))
            label = str(st.get("label", ""))
        else:
            value = str(st)
            label = ""
        add_text(slide, x + Inches(0.2), Inches(2.6), card_w - Inches(0.4),
                 Inches(1.2), value, 54, pal["accent"], bold=True,
                 align=PP_ALIGN.CENTER, font=pal["hfont"])
        add_text(slide, x + Inches(0.2), Inches(3.9), card_w - Inches(0.4),
                 Inches(1.0), label, 18, pal["muted"],
                 align=PP_ALIGN.CENTER, font=pal["bfont"])
    add_text(slide, Inches(11.9), Inches(7.12), Inches(1.3), Inches(0.35),
             f"{num} / {total}", 12, pal["muted"], align=PP_ALIGN.RIGHT,
             font=pal["bfont"])


def add_chart_slide(prs, chart, rows, pal, num, total):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["bg"])
    add_rect(slide, 0, 0, Inches(0.18), prs.slide_height, pal["accent2"])
    add_rect(slide, 0, 0, prs.slide_width, Inches(0.12), pal["accent"])
    add_text(slide, Inches(0.6), Inches(0.4), Inches(12.1), Inches(0.9),
             str(chart.get("title") or "Диаграмма"), 40, pal["text"],
             bold=True, font=pal["hfont"])
    add_rect(slide, Inches(0.62), Inches(1.45), Inches(2.6), Inches(0.08),
             pal["accent"])
    cd = CategoryChartData()
    cd.categories = [r[0] for r in rows]
    cd.add_series("Значения", tuple(r[1] for r in rows))
    ctype = XL_CHART_TYPE.PIE
    if str(chart.get("type", "")).lower() == "bar":
        ctype = XL_CHART_TYPE.BAR_CLUSTERED
    gf = slide.shapes.add_chart(
        ctype, Inches(1.0), Inches(1.8), Inches(11.3), Inches(5.1), cd
    )
    ch = gf.chart
    try:
        ch.has_legend = True
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
        ch.font.size = Pt(14)
        ch.font.color.rgb = pal["text"]
        plot = ch.plots[0]
        plot.has_data_labels = True
        plot.data_labels.font.size = Pt(13)
        plot.data_labels.font.color.rgb = pal["text"]
        if ctype == XL_CHART_TYPE.PIE:
            colors = (pal["accent"], pal["accent2"], pal["muted"],
                      pal["accent"], pal["accent2"], pal["muted"])
            for i, point in enumerate(plot.series[0].points):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = colors[i % len(colors)]
    except Exception as e:
        log.warning("chart styling skipped: %s", e)
    add_text(slide, Inches(11.9), Inches(7.12), Inches(1.3), Inches(0.35),
             f"{num} / {total}", 12, pal["muted"], align=PP_ALIGN.RIGHT,
             font=pal["bfont"])


def add_final_slide(prs, content, pal, num, total):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["accent2"])
    add_rect(slide, Inches(-1.4), Inches(-1.8), Inches(4.5), Inches(4.5),
             pal["accent"], MSO_SHAPE.OVAL)
    add_rect(slide, Inches(10.6), Inches(5.0), Inches(4.5), Inches(4.5),
             pal["bg"], MSO_SHAPE.OVAL)
    add_text(slide, Inches(1.2), Inches(1.9), Inches(10.9), Inches(1.7),
             str(content.get("conclusion") or ""), 28, WHITE,
             align=PP_ALIGN.CENTER, font=pal["bfont"])
    add_text(slide, Inches(1.2), Inches(3.8), Inches(10.9), Inches(1.6),
             "СПАСИБО ЗА ВНИМАНИЕ!", 56, WHITE, bold=True,
             align=PP_ALIGN.CENTER, font=pal["hfont"])
    add_text(slide, Inches(11.9), Inches(7.12), Inches(1.3), Inches(0.35),
             f"{num} / {total}", 12, WHITE, align=PP_ALIGN.RIGHT,
             font=pal["bfont"])


def build_pptx(content, s, slide_imgs, title_img):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = str(content.get("title") or s.topic)
    prs.core_properties.subject = s.topic
    prs.core_properties.author = "Presentation Bot"
    key = resolve_style(s)
    pal = dict(PALETTES.get(key, PALETTES["dark"]))
    hfont, bfont = THEME_FONTS.get(key, ("Trebuchet MS", "Calibri"))
    pal["hfont"] = hfont
    pal["bfont"] = bfont
    slides = content.get("slides", [])[: s.count]
    stats = content.get("stats") or []
    has_stats = isinstance(stats, list) and bool(stats)
    rows = chart_rows(content.get("chart"))
    has_chart = bool(rows)
    total = 2 + len(slides) + (1 if has_stats else 0)
    total += (1 if has_chart else 0) + 1
    add_title_slide(prs, content, pal, s, title_img)
    add_agenda_slide(prs, [it.get("title", "") for it in slides], pal)
    num = 3
    for i, item in enumerate(slides):
        img = slide_imgs[i] if i < len(slide_imgs) else None
        add_content_slide(prs, item, pal, num, total, img)
        num += 1
    if has_stats:
        add_stats_slide(prs, stats, pal, num, total)
        num += 1
    if has_chart:
        add_chart_slide(prs, content.get("chart"), rows, pal, num, total)
        num += 1
    add_final_slide(prs, content, pal, num, total)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def edit_keyboard(count):
    rows = []
    row = []
    for i in range(1, min(count, 8) + 1):
        row.append(
            InlineKeyboardButton(text=str(i), callback_data=f"edit:{i}")
        )
        if len(row) == 4:
            rows.append(row)
            row = []
    if row:
        rows.append(row)
    return InlineKeyboardMarkup(rows)


async def rebuild_and_send(update, context, s):
    if not s.last_content or not s.last_imgs:
        await update.message.reply_text("❌ Нет данных для пересборки")
        return
    
    chat_id = update.effective_chat.id
    try:
        data = build_pptx(s.last_content, s, s.last_imgs, s.last_title_img)
        name = re.sub(r"[^\w\- ]+", "", s.topic).strip()[:40]
        name = name or "presentation"
        await context.bot.send_document(
            chat_id=chat_id,
            document=InputFile(io.BytesIO(data), filename=f"{name}.pptx"),
            caption="✅ Слайд обновлён, новый файл выше 👆",
        )
    except Exception as e:
        log.exception("rebuild failed")
        await update.message.reply_text(f"😔 Ошибка пересборки: {e}")


async def generate_and_send(update, context, s):
    chat_id = update.effective_chat.id
    status = update.effective_message
    s.busy = True
    t0 = time.monotonic()
    try:
        await context.bot.send_chat_action(chat_id, ChatAction.TYPING)
        await status.edit_text("🧠 Пишу текст…")
        source = "AI"
        try:
            content = await ask_ai(s)
        except Exception as e:
            log.warning("AI unavailable, using Wikipedia fallback: %s", e)
            content = await fallback_content(s)
            source = "wiki"
        s.last_content = content
        slides = content.get("slides", [])[: s.count]
        await status.edit_text("🖼 Ищу фото (параллельно)…")
        queries = [str(it.get("img") or it.get("title") or s.topic)
                   for it in slides]
        slide_imgs = await fetch_images_parallel(queries, limit=4)
        title_img = None
        try:
            res = await fetch_images_parallel(
                [str(content.get("img") or s.topic)], limit=1
            )
            title_img = res[0] if res else None
        except Exception:
            title_img = None
        s.last_imgs = slide_imgs
        s.last_title_img = title_img
        await status.edit_text("🎨 Собираю слайды…")
        data = build_pptx(content, s, slide_imgs, title_img)
        name = re.sub(r"[^\w\- ]+", "", s.topic).strip()[:40]
        name = name or "presentation"
        pal_name = PALETTE_NAMES.get(resolve_style(s), "")
        pics = sum(1 for x in slide_imgs if x) + (1 if title_img else 0)
        secs = int(time.monotonic() - t0)
        caption = (f"✅ Готово за {secs}с! Оформление: {pal_name} "
                   f"| 🖼 фото: {pics} | v{VER}")
        if source == "wiki":
            caption += "\n⚠️ AI-лимит исчерпан — текст из Википедии"
        await context.bot.send_document(
            chat_id=chat_id,
            document=InputFile(io.BytesIO(data), filename=f"{name}.pptx"),
            caption=caption,
        )
        await context.bot.send_message(
            chat_id,
            "✏️ Хочешь поправить заголовок слайда? Нажми номер и напиши "
            "новый текст:",
            reply_markup=edit_keyboard(len(slides)),
        )
        try:
            await status.edit_text("✅ Презентация готова, файл выше 👆")
        except Exception:
            pass
    except Exception as e:
        log.exception("generation failed")
        reason = f"{type(e).__name__}: {e}"
        if len(reason) > 300:
            reason = reason[:300] + "…"
        try:
            await status.edit_text(
                "😔 Не получилось создать презентацию.\n"
                f"Причина: {reason}\n\n"
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
    app.add_handler(CommandHandler("cancel", cmd_cancel))
    app.add_handler(CallbackQueryHandler(on_callback))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND,
                                   on_message))

    await app.initialize()
    await app.start()
    await app.updater.start_polling(drop_pending_updates=True)
    log.info("Bot started (polling), version %s", VER)
    log.info("Model chain: %s", MODEL_CHAIN)

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