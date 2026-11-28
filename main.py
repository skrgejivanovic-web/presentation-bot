import asyncio
import io
import json
import logging
import os
import re
import time
from datetime import datetime
import base64

import aiohttp
from aiohttp import web
from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
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

VER = "23.1"

BOT_TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_TOKEN", "")
GIGACHAT_CLIENT_ID = os.getenv("GIGACHAT_CLIENT_ID", "")
GIGACHAT_AUTH_KEY = os.getenv("GIGACHAT_AUTH_KEY", "")

PORT = int(os.getenv("PORT", "10000"))

COMMONS_API = "https://commons.wikimedia.org/w/api.php"
WIKI_API = "https://ru.wikipedia.org/w/api.php"
EN_WIKI_API = "https://en.wikipedia.org/w/api.php"
OPENVERSE_API = "https://api.openverse.org/v1/images/"
HTTP_HEADERS = {"User-Agent": "presentation-bot/2.0 (telegram bot; contact: dev@example.com)"}

# ✅ ИСПРАВЛЕНО: Правильный официальный URL для OAuth и Chat
GIGACHAT_OAUTH_URL = "https://gigachat.devices.sberbank.ru/api/v1/oauth/token"
REAL_GIGACHAT_API_URL = "https://gigachat.devices.sberbank.ru/api/v1/chat/completions"

COUNTS = (3, 5, 7, 10, 12, 15)

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
    "dark": {"bg": RGBColor(0x17, 0x17, 0x20), "panel": RGBColor(0x26, 0x26, 0x33), "accent": RGBColor(0x00, 0xB8, 0xD4), "accent2": RGBColor(0x7C, 0x4D, 0xFF), "text": RGBColor(0xF2, 0xF2, 0xF5), "muted": RGBColor(0xC0, 0xC0, 0xD0)},
    "light": {"bg": RGBColor(0xFA, 0xFA, 0xFC), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0x1F, 0x6F, 0xEB), "accent2": RGBColor(0x0F, 0x9D, 0x58), "text": RGBColor(0x1E, 0x1E, 0x28), "muted": RGBColor(0x5A, 0x5A, 0x6A)},
    "color": {"bg": RGBColor(0xFD, 0xF6, 0xEC), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0xFF, 0x6B, 0x35), "accent2": RGBColor(0x2E, 0x86, 0xAB), "text": RGBColor(0x2B, 0x2B, 0x2B), "muted": RGBColor(0x6A, 0x60, 0x55)},
    "med": {"bg": RGBColor(0xF2, 0xFA, 0xF8), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0x00, 0x96, 0x88), "accent2": RGBColor(0x00, 0x69, 0x5C), "text": RGBColor(0x16, 0x38, 0x32), "muted": RGBColor(0x4A, 0x60, 0x5A)},
    "space": {"bg": RGBColor(0x0B, 0x10, 0x26), "panel": RGBColor(0x16, 0x1D, 0x3A), "accent": RGBColor(0x8A, 0xB4, 0xFF), "accent2": RGBColor(0x7C, 0x4D, 0xFF), "text": RGBColor(0xEE, 0xF0, 0xFF), "muted": RGBColor(0xAB, 0xB2, 0xD8)},
    "nature": {"bg": RGBColor(0xF3, 0xF8, 0xEE), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0x55, 0x8B, 0x2F), "accent2": RGBColor(0x33, 0x69, 0x1E), "text": RGBColor(0x26, 0x30, 0x1C), "muted": RGBColor(0x55, 0x60, 0x4A)},
    "finance": {"bg": RGBColor(0x14, 0x18, 0x1D), "panel": RGBColor(0x1E, 0x24, 0x2B), "accent": RGBColor(0xD4, 0xAF, 0x37), "accent2": RGBColor(0x2E, 0x7D, 0x32), "text": RGBColor(0xF0, 0xED, 0xE6), "muted": RGBColor(0xB8, 0xBE, 0xC8)},
    "tech": {"bg": RGBColor(0x10, 0x14, 0x18), "panel": RGBColor(0x1A, 0x20, 0x26), "accent": RGBColor(0x00, 0xE5, 0xFF), "accent2": RGBColor(0x7C, 0x4D, 0xFF), "text": RGBColor(0xE8, 0xF6, 0xF8), "muted": RGBColor(0xA8, 0xBB, 0xC5)},
    "edu": {"bg": RGBColor(0xFD, 0xFD, 0xF7), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0x1E, 0x88, 0xE5), "accent2": RGBColor(0xFB, 0x8C, 0x00), "text": RGBColor(0x21, 0x21, 0x21), "muted": RGBColor(0x55, 0x55, 0x55)},
    "sport": {"bg": RGBColor(0x1A, 0x1A, 0x1A), "panel": RGBColor(0x26, 0x26, 0x26), "accent": RGBColor(0xFF, 0x6D, 0x00), "accent2": RGBColor(0x00, 0xC8, 0x53), "text": RGBColor(0xFF, 0xFF, 0xFF), "muted": RGBColor(0xC8, 0xC8, 0xC8)},
    "art": {"bg": RGBColor(0xFF, 0xF7, 0xF0), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0xE9, 0x1E, 0x63), "accent2": RGBColor(0x6A, 0x1B, 0x9A), "text": RGBColor(0x3A, 0x24, 0x30), "muted": RGBColor(0x7A, 0x65, 0x70)},
    "food": {"bg": RGBColor(0xFF, 0xF8, 0xF0), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0xE8, 0x5D, 0x04), "accent2": RGBColor(0xBC, 0x47, 0x49), "text": RGBColor(0x3D, 0x40, 0x5B), "muted": RGBColor(0x7D, 0x85, 0x9E)},
    "travel": {"bg": RGBColor(0xE0, 0xF7, 0xFA), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0x00, 0x97, 0xA7), "accent2": RGBColor(0x00, 0x60, 0x64), "text": RGBColor(0x26, 0x32, 0x38), "muted": RGBColor(0x58, 0x70, 0x7C)},
    "history": {"bg": RGBColor(0x3E, 0x27, 0x23), "panel": RGBColor(0x4E, 0x34, 0x2E), "accent": RGBColor(0xFF, 0xD5, 0x4F), "accent2": RGBColor(0xB7, 0x1C, 0x1C), "text": RGBColor(0xFF, 0xEC, 0xB3), "muted": RGBColor(0xDD, 0xCC, 0xAA)},
    "science": {"bg": RGBColor(0x1A, 0x23, 0x7E), "panel": RGBColor(0x28, 0x35, 0x93), "accent": RGBColor(0x7C, 0x4D, 0xFF), "accent2": RGBColor(0x53, 0x6D, 0xFE), "text": RGBColor(0xE8, 0xEA, 0xF6), "muted": RGBColor(0xBF, 0xC8, 0xF6)},
    "music": {"bg": RGBColor(0x4A, 0x14, 0x8C), "panel": RGBColor(0x6A, 0x1B, 0x9A), "accent": RGBColor(0xE0, 0x40, 0xFB), "accent2": RGBColor(0xAA, 0x00, 0xFF), "text": RGBColor(0xF3, 0xE5, 0xF5), "muted": RGBColor(0xDE, 0xB3, 0xE8)},
    "gaming": {"bg": RGBColor(0x0D, 0x11, 0x17), "panel": RGBColor(0x16, 0x1B, 0x22), "accent": RGBColor(0x39, 0xD3, 0x53), "accent2": RGBColor(0x00, 0xFF, 0x88), "text": RGBColor(0xC9, 0xD1, 0xD9), "muted": RGBColor(0xAB, 0xB4, 0xBE)},
    "fashion": {"bg": RGBColor(0xFA, 0xFA, 0xFA), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0x00, 0x00, 0x00), "accent2": RGBColor(0x33, 0x33, 0x33), "text": RGBColor(0x1A, 0x1A, 0x1A), "muted": RGBColor(0x55, 0x55, 0x55)},
    "auto": {"bg": RGBColor(0x21, 0x21, 0x21), "panel": RGBColor(0x30, 0x30, 0x30), "accent": RGBColor(0xFF, 0x17, 0x44), "accent2": RGBColor(0xC5, 0x11, 0x62), "text": RGBColor(0xF5, 0xF5, 0xF5), "muted": RGBColor(0xBE, 0xBE, 0xBE)},
    "real_estate": {"bg": RGBColor(0xF1, 0xF8, 0xE9), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0x55, 0x8B, 0x2F), "accent2": RGBColor(0x33, 0x69, 0x1E), "text": RGBColor(0x1B, 0x5E, 0x20), "muted": RGBColor(0x48, 0x7F, 0x38)},
    "psychology": {"bg": RGBColor(0xE3, 0xF2, 0xFD), "panel": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0x19, 0x76, 0xD2), "accent2": RGBColor(0x0D, 0x47, 0xA1), "text": RGBColor(0x0D, 0x47, 0xA1), "muted": RGBColor(0x44, 0x75, 0xB6)},
}

THEME_FONTS = {
    "dark": ("Segoe UI", "Calibri"),
    "light": ("Georgia", "Calibri"),
    "color": ("Verdana", "Trebuchet MS"),
    "med": ("Segoe UI", "Calibri"),
    "space": ("Consolas", "Segoe UI"),
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
    (("медиц", "здоров", "врач", "больн", "лекар", "therapy", "health", "medical", "doctor", "hospital"), "med"),
    (("косм", "звезд", "планет", "вселен", "галакт", "orbit", "rocket", "moon", "mars"), "space"),
    (("природ", "экол", "лес", "растен", "животн", "климат", "ocean", "wildlife", "eco"), "nature"),
    (("финанс", "деньг", "банк", "инвест", "эконом", "бизнес", "budget", "stock", "profit"), "finance"),
    (("техно", "робот", "программ", "нейросет", "intell", "computer", "digital", "code"), "tech"),
    (("школ", "учеб", "образован", "курс", "student", "lesson", "exam"), "edu"),
    (("спорт", "футбол", "fitness", "olympic", "game", "match"), "sport"),
    (("искусств", "культур", "музык", "театр", "design", "paint"), "art"),
    (("еда", "кулинар", "рецепт", "готов", "food", "cook", "dish"), "food"),
    (("путешеств", "туризм", "отпуск", "стран", "travel", "trip", "city"), "travel"),
    (("истор", "древн", "войн", "импер", "history", "war", "king"), "history"),
    (("наук", "физик", "хим", "биолог", "math", "physics", "bio"), "science"),
    (("игр", "геймер", "video game", "play", "quest"), "gaming"),
    (("мод", "стиль", "одежд", "fashion", "style", "wear"), "fashion"),
    (("авто", "машин", "автомобил", "car", "drive", "transport"), "auto"),
    (("недвижим", "дом", "квартир", "rent", "house", "flat"), "real_estate"),
    (("психолог", "мозг", "сознан", "mind", "brain", "emotion"), "psychology"),
)

WHITE = RGBColor(0xFF, 0xFF, 0xFF)


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
        self.history = []


sessions = {}


def get_session(user_id):
    if user_id not in sessions:
        sessions[user_id] = Session()
    return sessions[user_id]


def auto_style(topic):
    low = (topic or "").lower()
    scores = {}
    for keywords, theme_key in THEME_KEYWORDS:
        score = 0
        for kw in keywords:
            if kw.lower() in low:
                weight = min(len(kw), 15) / 10
                score += weight
        if score > 0:
            scores[theme_key] = score
    if scores:
        best_theme = max(scores.items(), key=lambda x: x[1])[0]
        log.info(f"Theme matched: {best_theme}")
        return best_theme
    pool = ("dark", "light", "color", "art", "space", "tech")
    idx = sum(ord(ch) for ch in low) % len(pool)
    fallback = pool[idx]
    log.info(f"No keyword match, using deterministic fallback: {fallback}")
    return fallback


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
        rows.append([InlineKeyboardButton(mark(s.count == n, str(n)), callback_data=f"n:{n}") for n in row])
    rows.append([
        InlineKeyboardButton(mark(s.lang == "ru", "🇷🇺 Русский"), callback_data="lang:ru"),
        InlineKeyboardButton(mark(s.lang == "en", "🇬🇧 English"), callback_data="lang:en"),
    ])
    rows.append([
        InlineKeyboardButton(mark(s.style == "auto", "🪄 Авто"), callback_data="style:auto"),
        InlineKeyboardButton(mark(s.style == "dark", "🌙 Тёмная"), callback_data="style:dark"),
        InlineKeyboardButton(mark(s.style == "light", "☀️ Светлая"), callback_data="style:light"),
        InlineKeyboardButton(mark(s.style == "color", "🎨 Яркая"), callback_data="style:color"),
    ])
    if s.history:
        hist_row = []
        recent_topics = s.history[-3:]
        for i, h in enumerate(recent_topics, start=1):
            short_name = h[:12] + "..." if len(h) > 12 else h
            hist_row.append(InlineKeyboardButton(short_name, callback_data=f"hidx:{i}"))
        if hist_row:
            rows.append(hist_row)
    rows.append([InlineKeyboardButton("🚀 Создать презентацию", callback_data="go")])
    return InlineKeyboardMarkup(rows)


def settings_text(s):
    return (
        f"📌 Тема: {s.topic}\n\n"
        f"Слайдов: {s.count}   Оформление: {style_label(s)}\n"
        "Настрой и жми «Создать». Или напиши новую тему сообщением.\n"
        "Команды: /help — помощь, /content — текст презентации"
    )


async def refresh(q, s):
    try:
        await q.edit_message_text(settings_text(s), reply_markup=settings_keyboard(s))
    except Exception:
        pass


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    sessions[update.effective_user.id] = Session()
    await update.message.reply_text(
        "Привет! Я делаю красивые презентации .pptx: дизайн под тему, "
        "картинки, диаграммы, факты и цитаты.\n\n"
        "Напиши тему одним сообщением.\n"
        "Например: «Искусственный интеллект в медицине».\n\n"
        f"(версия {VER}) Команды: /help"
    )


async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📚 *Справка по боту*\n\n"
        "1️⃣ Напиши тему презентации\n"
        "2️⃣ Выбери количество слайдов (3-15)\n"
        "3️⃣ Выбери язык и стиль оформления\n"
        "4️⃣ Нажми «🚀 Создать презентацию»\n\n"
        "📝 *Команды:*\n"
        "/start — начать заново\n"
        "/help — эта справка\n"
        "/content — текстовое содержание последней презентации\n"
        "/cancel — отменить редактирование слайда\n\n"
        "✏️ После генерации можно нажать на номер слайда и изменить его заголовок\n\n"
        f"Версия {VER}",
        parse_mode="Markdown"
    )


async def cmd_content(update: Update, context: ContextTypes.DEFAULT_TYPE):
    s = get_session(update.effective_user.id)
    if not s.last_content:
        await update.message.reply_text("❌ Сначала создай презентацию")
        return
    text = f"📋 *{s.last_content.get('title', s.topic)}*\n\n"
    text += f"_{s.last_content.get('subtitle', '')}_\n\n"
    for i, slide in enumerate(s.last_content.get("slides", []), 1):
        text += f"*{i}. {slide.get('title', '')}*\n"
        text += f"_{slide.get('lead', '')}_\n"
        for b in slide.get("bullets", [])[:3]:
            text += f"• {b}\n"
        if slide.get("fact"):
            text += f"💡 {slide['fact']}\n"
        text += "\n"
    if len(text) > 4000:
        text = text[:4000] + "..."
    await update.message.reply_text(text, parse_mode="Markdown")


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
        if s.last_content and 1 <= s.editing_slide <= len(s.last_content.get("slides", [])):
            s.last_content["slides"][s.editing_slide - 1]["title"] = text[:120]
            s.editing_slide = None
            await rebuild_and_send(update, context, s)
        else:
            s.editing_slide = None
            await update.message.reply_text("❌ Слайд не найден, отменено")
        return
    if s.busy:
        await update.message.reply_text("⏳ Подожди, я ещё делаю предыдущую презентацию.")
        return
    s.topic = text[:300]
    if s.topic not in s.history:
        s.history.append(s.topic)
        if len(s.history) > 5:
            s.history.pop(0)
    await update.message.reply_text(settings_text(s), reply_markup=settings_keyboard(s))


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

    elif data.startswith("hidx:"):
        try:
            idx = int(data.split(":", 1)[1]) - 1
            recent_topics = s.history[-3:]
            if 0 <= idx < len(recent_topics):
                s.topic = recent_topics[idx]
                await q.answer()
                await refresh(q, s)
            else:
                await q.answer("Тема уже удалена из истории")
        except ValueError:
            await q.answer("Ошибка данных")

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
        await q.message.reply_text(f"✏️ Напиши новый заголовок для слайда {s.editing_slide} (или /cancel)")

    elif data == "go":
        if s.busy:
            await q.answer("Уже генерирую, подожди 🙂")
            return
        if not s.topic:
            await q.answer("Сначала напиши тему сообщением.")
            return
        await q.answer()
        await generate_simple_flow(update, context, s)

    elif data == "regen":
        if s.busy:
            await q.answer("Уже генерирую")
            return
        if not s.topic:
            await q.answer("Нет темы")
            return
        await q.answer()
        await generate_simple_flow(update, context, s)


PROMPT = (
    "You are a professional presentation designer and expert content writer. "
    "Return ONLY valid JSON, no markdown fences, no explanations outside JSON. "
    'JSON structure: {"title": string (max 6 words), "subtitle": string (one catchy phrase), '
    '"img": string (3-5 ENGLISH words describing the main photo for title slide), '
    '"intro": string (2-3 engaging sentences introducing the topic), '
    '"slides": array of EXACTLY {count} objects. Each slide object MUST have these fields: '
    '"title" (string, 2-5 words, descriptive NOT generic like "Основная часть"), '
    '"lead" (string, ONE compelling sentence that hooks attention), '
    '"img" (string, 3-5 ENGLISH words with specific visual subject first, e.g. "surgeon operating robot arm"), '
    '"bullets" (array of EXACTLY 4 informative strings, each 10-20 words, containing REAL facts/data/examples about THIS SPECIFIC sub-topic), '
    '"scheme" (array of EXACTLY 3 short strings representing logical steps/phases/aspects of this slide\'s topic, each 3-8 words), '
    '"fact" (string, one surprising statistic or memorable fact related to this slide), '
    '"speaker_notes" (string, 2-3 sentences coaching the presenter what to emphasize). '
    'Also include top-level: "stats" (array of 3 objects with "value" and "label"), '
    '"chart" (object with "title", "type" as pie or bar, "data" array of 4-6 objects with "label" and numeric "value"), '
    '"conclusion" (string, 2-3 inspiring closing sentences). '
    "CRITICAL RULES: "
    "1. Write EVERYTHING strictly in {lang}. Never mix languages. "
    "2. Slide titles must be SPECIFIC and INFORMATIVE (e.g., 'Диагностика рака с помощью нейросетей' NOT 'Основная часть'). "
    "3. Bullets must contain CONCRETE information: numbers, names, dates, examples. No vague filler. "
    "4. Scheme items must form a LOGICAL SEQUENCE relevant to that slide's content. "
    "5. Create EXACTLY {count} slides. Count them carefully before responding. "
    "6. Make it genuinely INTERESTING and EDUCATIONAL. Topic: {topic}"
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


async def get_gigachat_access_token():
    """Получаем временный access token по Client ID и Secret"""
    auth_header = f"Basic {base64.b64encode(f'{GIGACHAT_CLIENT_ID}:{GIGACHAT_AUTH_KEY}'.encode()).decode()}"
    
    payload = {
        "grant_type": "client_credentials",
        "scope": "GIGACHAT_API_PERS"
    }
    
    headers = {
        "Authorization": auth_header,
        "Content-Type": "application/json"
    }
    
    timeout = aiohttp.ClientTimeout(total=10)
    async with aiohttp.ClientSession(timeout=timeout) as sess:
        try:
            log.debug(f"Requesting GigaChat OAuth token... URL: {GIGACHAT_OAUTH_URL}")
            async with sess.post(GIGACHAT_OAUTH_URL, json=payload, headers=headers) as r:
                body = await r.text()
                if r.status != 200:
                    log.error(f"GigaChat OAuth FAILED ({r.status}): {body[:500]}")
                    raise ValueError(f"GigaChat OAuth error {r.status}: {body[:200]}")
                
                data = await r.json()
                token = data.get("access_token")
                if not token:
                    raise ValueError("No access_token in response")
                log.info("✓ GigaChat access token obtained successfully")
                return token
        except Exception as e:
            log.error(f"GigaChat OAuth exception: {e}")
            raise


async def ask_gigachat(prompt_text):
    """Отправляем запрос к GigaChat напрямую через aiohttp"""
    access_token = await get_gigachat_access_token()
    
    messages = [
        {"role": "system", "content": "You are an expert presentation writer. Always respond with valid JSON only."},
        {"role": "user", "content": prompt_text}
    ]
    
    # Пробуем разные модели, начиная с Pro
    models_to_try = ["GigaChat-Pro", "GigaChat-Max", "SberGigaChat-2-Pro"]
    
    for model_name in models_to_try:
        payload = {
            "model": model_name,
            "messages": messages,
            "temperature": 0.7,
            "max_tokens": 4000,
        }
        
        headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
        timeout = aiohttp.ClientTimeout(total=90)
        
        async with aiohttp.ClientSession(timeout=timeout) as sess:
            try:
                log.debug(f"Trying GigaChat model: {model_name}...")
                async with sess.post(REAL_GIGACHAT_API_URL, json=payload, headers=headers) as r:
                    err_body = await r.text()
                    
                    if r.status == 200:
                        resp_data = await r.json()
                        choices = resp_data.get("choices", [])
                        if not choices:
                            raise RuntimeError("Empty choices from GigaChat")
                        
                        content = choices[0].get("message", {}).get("content", "")
                        log.info(f"✓ GigaChat [{model_name}] responded with {len(content)} chars")
                        return content
                    
                    elif r.status == 400:
                        log.warning(f"Model {model_name} rejected request (400 Bad Request): {err_body[:200]}")
                        continue # Try next model
                    else:
                        log.error(f"GigaChat API Error {r.status} for {model_name}: {err_body[:500]}")
                        raise RuntimeError(f"GigaChat returned HTTP {r.status}")
                        
            except asyncio.TimeoutError:
                log.warning(f"GigaChat [{model_name}] timed out after 90s")
                continue
            except Exception as e:
                log.exception(f"Unexpected error calling GigaChat [{model_name}]")
                continue
    
    raise RuntimeError("All GigaChat models failed to respond")


async def ask_ai_multi(s):
    lang = "Russian" if s.lang == "ru" else "English"
    prompt = PROMPT.format(lang=lang, count=s.count, topic=s.topic)
    
    for attempt in range(3):
        try:
            raw_response = await ask_gigachat(prompt)
            json_str = extract_json(raw_response)
            if not json_str:
                log.warning(f"Attempt {attempt+1}: No JSON found in response")
                continue
            
            content = json.loads(json_str)
            slides = content.get("slides", [])
            
            if isinstance(slides, list) and len(slides) >= s.count:
                log.info(f"Success on attempt {attempt+1}: got {len(slides)} slides")
                return content
            else:
                log.warning(f"Attempt {attempt+1}: Only {len(slides) if isinstance(slides, list) else 0} slides, need {s.count}")
        except Exception as e:
            log.warning(f"Attempt {attempt+1} failed: {e}")
            if attempt == 2:
                raise
    
    raise RuntimeError("All 3 attempts to get AI content failed")


# УЛУЧШЕННЫЙ fallback из Википедии — теперь с реальной структурой
async def smart_fallback_from_wikipedia(s):
    api = WIKI_API if s.lang == "ru" else EN_WIKI_API
    params = {
        "action": "query",
        "format": "json",
        "generator": "search",
        "gsrsearch": s.topic,
        "gsrlimit": "1",
        "gsrnamespace": "0",
        "prop": "extracts|pageimages",
        "exintro": "0",  # НЕ только введение, а весь текст
        "explaintext": "1",
        "piprop": "thumbnail",
        "pithumbsize": "900",
    }
    timeout = aiohttp.ClientTimeout(total=15)
    async with aiohttp.ClientSession(timeout=timeout, headers=HTTP_HEADERS) as sess:
        async with sess.get(api, params=params) as r:
            data = await r.json(content_type=None)
    pages = (data.get("query") or {}).get("pages") or {}
    if not pages:
        raise ValueError("no Wikipedia article found")
    page = list(pages.values())[0]
    article_title = page.get("title", s.topic)
    full_extract = (page.get("extract") or "").strip()
    thumbnail_url = (page.get("thumbnail") or {}).get("source")
    
    # Разбиваем на предложения
    all_sentences = [x.strip() for x in re.split(r"(?<=[.!?])\s+", full_extract) if x.strip()]
    
    # Фильтруем хорошие предложения (длинные, информативные)
    good_sentences = [sent for sent in all_sentences if 25 < len(sent) < 250]
    
    if not good_sentences:
        good_sentences = ["Основные сведения о теме " + article_title, 
                          "Подробное рассмотрение предмета",
                          "Практическое значение и применение",
                          "Перспективы развития направления"]
    
    slides = []
    
    # Определяем разделы статьи через заголовки (если есть)
    sections = re.split(r"\n==+\s*(.+?)\s*==+\n?", full_extract)
    
    if len(sections) >= 4:
        # Есть реальные секции из Википедии
        section_titles = []
        section_texts = []
        for i in range(1, len(sections), 2):
            sec_title = sections[i].strip().replace("=", "").strip()
            sec_text = sections[i+1] if i+1 < len(sections) else ""
            if sec_title and len(sec_text) > 50:
                section_titles.append(sec_title)
                section_texts.append(sec_text)
        
        num_slides_needed = min(s.count - 2, len(section_titles), 8)
        
        for i in range(num_slides_needed):
            sec_sents = [x.strip() for x in re.split(r"(?<=[.!?])\s+", section_texts[i]) if x.strip()]
            bullets = [sent for sent in sec_sents if 25 < len(sent) < 220][:4]
            
            if len(bullets) < 3:
                extra = [sent for sent in good_sentences if sent not in bullets][:3-len(bullets)]
                bullets.extend(extra)
            
            # ИСПРАВЛЕНО: Схема всегда заполняется реальными данными из буллетов
            scheme_items = []
            for j in range(min(3, len(bullets))):
                item = bullets[j][:60] if len(bullets[j]) > 60 else bullets[j]
                scheme_items.append(item)
            
            slides.append({
                "title": section_titles[i][:80],
                "lead": bullets[0][:150] if bullets else f"Раздел «{section_titles[i]}»",
                "img": f"{article_title} {section_titles[i]}",
                "bullets": bullets[:4],
                "scheme": scheme_items if len(scheme_items) >= 2 else [],
                "fact": "",
                "speaker_notes": f"Расскажите подробнее про «{section_titles[i]}».",
            })
    else:
        # Нет секций — делим текст равномерно
        chunk_size = max(3, len(good_sentences) // (s.count - 2)) if s.count > 2 else 3
        num_sections = min(s.count - 2, max(3, len(good_sentences) // chunk_size))
        
        generic_titles_ru = [
            "Определение и суть", "История возникновения", "Ключевые принципы",
            "Практическое применение", "Преимущества и возможности",
            "Вызовы и ограничения", "Тренды и будущее", "Известные примеры",
            "Методы и технологии", "Влияние на общество"
        ]
        generic_titles_en = [
            "Definition & Overview", "History & Origins", "Core Principles",
            "Practical Applications", "Advantages & Capabilities",
            "Challenges & Limitations", "Future Trends", "Notable Examples",
            "Methods & Technologies", "Societal Impact"
        ]
        titles_pool = generic_titles_ru if s.lang == "ru" else generic_titles_en
        
        for i in range(num_sections):
            start_idx = i * chunk_size
            end_idx = start_idx + chunk_size
            chunk_sents = good_sentences[start_idx:end_idx]
            
            bullets = []
            for sent in chunk_sents:
                cleaned = sent.replace("\n", " ").strip()
                if 25 < len(cleaned) < 220:
                    bullets.append(cleaned)
            
            if len(bullets) < 3:
                fillers = ["Конкретные данные по этому аспекту", 
                           "Практические примеры применения",
                           "Экспертные оценки significance"]
                for f in fillers:
                    if len(bullets) >= 4:
                        break
                    if f not in bullets:
                        bullets.append(f)
            
            # ИСПРАВЛЕНО: Схема всегда заполняется реальными данными из буллетов
            scheme_items = [b[:55] for b in bullets[:3]]
            
            title = titles_pool[i] if i < len(titles_pool) else f"Аспект {i+1}"
            
            slides.append({
                "title": title,
                "lead": bullets[0][:150] if bullets else f"Рассмотрение темы: {title}",
                "img": f"{article_title} {title}",
                "bullets": bullets[:4],
                "scheme": scheme_items if len(scheme_items) >= 2 else [],
                "fact": "",
                "speaker_notes": f"Раскройте тему «{title}» подробнее.",
            })
    
    # Добавляем заключение если нужно
    if len(slides) < s.count - 1:
        last_bullets = good_sentences[-3:] if len(good_sentences) >= 3 else ["Итоги рассмотрения", "Запомните главное"]
        slides.append({
            "title": "Итоги и выводы" if s.lang == "ru" else "Summary & Conclusions",
            "lead": f"Обобщение ключевых моментов по теме «{article_title}»",
            "img": article_title,
            "bullets": last_bullets[:4],
            "scheme": [],
            "fact": "",
            "speaker_notes": "Подведите итоги и поблагодарите аудиторию.",
        })
    
    stats = [
        {"value": str(len(all_sentences)), "label": "предложений в источнике"},
        {"value": str(len(slides)), "label": "тематических блоков"},
        {"value": "Wikipedia", "label": "источник данных"},
    ]
    
    chart = {
        "title": f"Структура: {article_title}",
        "type": "bar",
        "data": [{"label": sl.get("title", "")[:20], "value": len(sl.get("bullets", []))} for sl in slides[:6]],
    }
    
    result = {
        "title": article_title,
        "subtitle": f"На основе материалов Википедии • {datetime.now().strftime('%d.%m.%Y')}",
        "img": article_title,
        "intro": f"Презентация подготовлена автоматически на основе энциклопедических источников о теме «{article_title}». Рассмотрены ключевые аспекты, практическое значение и перспективы.",
        "slides": slides[:s.count],
        "stats": stats,
        "chart": chart,
        "conclusion": f"Мы рассмотрели основные грани темы «{article_title}». Эти знания помогут глубже понять предмет и применить их на практике.",
    }
    
    log.info(f"Wikipedia fallback: created {len(result['slides'])} rich slides from '{article_title}'")
    return result, thumbnail_url


def prepare_image(raw, w=900, h=600, quality=85):
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
    img = img.resize((w, h), Image.LANCZOS)
    out = io.BytesIO()
    img.save(out, format="JPEG", quality=quality)
    return out.getvalue()


def _tokens(q):
    return set(re.findall(r"[a-zа-яё]{3,}", (q or "").lower()))


def _relevant(query, title):
    qw = _tokens(query)
    tw = _tokens(title)
    if not qw or not tw:
        return True
    overlap = qw & tw
    return len(overlap) >= 1


async def _download_image(sess, url, quality=85):
    try:
        async with sess.get(url) as r:
            if r.status != 200:
                return None
            ct = r.headers.get("content-type", "")
            if "image" not in ct and "octet-stream" not in ct:
                return None
            raw = await r.read()
        if not raw or len(raw) < 5000:
            return None
        return prepare_image(raw, quality=quality)
    except Exception as e:
        log.debug("image download failed %s: %s", url[:80], e)
        return None


async def fetch_topic_image(query, quality=85):
    timeout = aiohttp.ClientTimeout(total=6, connect=3)
    async with aiohttp.ClientSession(timeout=timeout, headers=HTTP_HEADERS) as sess:
        # Источник 1: Wikimedia Commons (самый надёжный)
        try:
            params = {
                "action": "query", "format": "json",
                "generator": "search",
                "gsrnamespace": "6",
                "gsrsearch": f"filetype:bitmap {query}",
                "gsrlimit": "8",
                "prop": "imageinfo",
                "iiprop": "url|mime|size",
                "iiurlwidth": "900",
            }
            async with sess.get(COMMONS_API, params=params) as r:
                data = await r.json(content_type=None)
            pages = sorted(
                ((data.get("query") or {}).get("pages") or {}).values(),
                key=lambda p: -(p.get("index", 0))
            )
            for page in pages:
                if not _relevant(query, page.get("title", "")):
                    continue
                info_list = page.get("imageinfo") or [{}]
                info = info_list[0]
                mime = info.get("mime", "")
                if mime not in ("image/jpeg", "image/png"):
                    continue
                width = info.get("width", 0)
                height = info.get("height", 0)
                if width < 400 or height < 300:
                    continue
                url = info.get("thumburl") or info.get("url")
                if url:
                    img = await _download_image(sess, url, quality)
                    if img:
                        log.info("✓ image from commons: %s (%dx%d)", query[:30], width, height)
                        return img
        except Exception as e:
            log.debug("commons failed for %s: %s", query[:30], e)
        
        # Источник 2: Wikipedia article thumbnails
        try:
            params = {
                "action": "query", "format": "json",
                "generator": "search",
                "gsrsearch": query,
                "gsrlimit": "5",
                "gsrnamespace": "0",
                "prop": "pageimages",
                "piprop": "thumbnail",
                "pithumbsize": "900",
            }
            async with sess.get(WIKI_API, params=params) as r:
                data = await r.json(content_type=None)
            pages = (data.get("query") or {}).get("pages") or {}
            for page in pages.values():
                thumb = page.get("thumbnail") or {}
                url = thumb.get("source", "")
                if url:
                    img = await _download_image(sess, url, quality)
                    if img:
                        log.info("✓ image from wiki: %s", query[:30])
                        return img
        except Exception as e:
            log.debug("wiki image failed for %s: %s", query[:30], e)
        
        # Источник 3: Openverse
        try:
            async with sess.get(OPENVERSE_API, params={"q": query, "page_size": "8", "license_type": "all"}) as r:
                if r.status == 200:
                    data = await r.json(content_type=None)
                    results = data.get("results") or []
                    for res in results:
                        url = res.get("url") or res.get("thumbnail") or ""
                        if url and url.endswith((".jpg", ".jpeg", ".png", ".webp")):
                            img = await _download_image(sess, url, quality)
                            if img:
                                log.info("✓ image from openverse: %s", query[:30])
                                return img
        except Exception as e:
            log.debug("openverse failed for %s: %s", query[:30], e)
    
    return None


async def fetch_images_parallel(queries, limit=6, quality=85):
    """Параллельная загрузка картинок с общим таймаутом"""
    tasks = []
    task_indices = []
    for i, q in enumerate(queries):
        if i >= limit or not q:
            continue
        tasks.append(fetch_topic_image(q, quality))
        task_indices.append(i)
    
    if not tasks:
        return [None] * len(queries)
    
    try:
        results = await asyncio.wait_for(
            asyncio.gather(*tasks, return_exceptions=True),
            timeout=20  # Общий таймаут 20 секунд
        )
    except asyncio.TimeoutError:
        log.warning("Image loading global timeout exceeded (20s)")
        results = [None] * len(tasks)
    
    out = [None] * len(queries)
    for idx, r in zip(task_indices, results):
        if isinstance(r, BaseException):
            out[idx] = None
        else:
            out[idx] = r
    
    loaded_count = sum(1 for x in out if x is not None)
    log.info(f"Images loaded: {loaded_count}/{min(limit, len(queries))} requested")
    return out


# ==================== ДИЗАЙН СЛАЙДОВ (ПОЛНОСТЬЮ ПЕРЕРАБОТАН) ====================

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


def add_rounded_rect(slide, x, y, w, h, fill_color, border_color=None, border_width=Pt(1)):
    sp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    sp.fill.solid()
    sp.fill.fore_color.rgb = fill_color
    if border_color:
        sp.line.color.rgb = border_color
        sp.line.width = border_width
    else:
        sp.line.fill.background()
    sp.shadow.inherit = False
    return sp


def add_text_box(slide, x, y, w, h, text, size, color, bold=False, italic=False, align=PP_ALIGN.LEFT, font="Calibri", line_spacing=1.15):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = align
    p.line_spacing = line_spacing
    run = p.runs[0] if p.runs else p.add_run()
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = font
    return tb


def truncate_text(text, max_chars):
    """Безопасно обрезает текст, сохраняя слова целыми"""
    if len(text) <= max_chars:
        return text
    truncated = text[:max_chars]
    last_space = truncated.rfind(" ")
    if last_space > max_chars * 0.6:
        truncated = truncated[:last_space]
    return truncated.rstrip(',.;:!?') + '...'


def fit_font_size(text, box_width_inches, target_height_pt, base_size, min_size=10):
    """Подбирает размер шрифта чтобы текст помещался"""
    approx_chars_per_line = int(box_width_inches * 72 / (base_size * 0.55))
    lines_needed = max(1, len(text) // approx_chars_per_line + 1)
    estimated_height = lines_needed * base_size * 1.3
    if estimated_height <= target_height_pt:
        return base_size
    scale = target_height_pt / estimated_height
    return max(min_size, int(base_size * scale))


# --- ТИТУЛЬНЫЙ СЛАЙД (с большим заголовком!) ---
def add_title_slide(prs, content, pal, s, img, date_str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["bg"])
    
    # Декоративные круги
    add_rect(slide, Inches(-1.8), Inches(4.8), Inches(5.5), Inches(5.5), pal["accent2"], MSO_SHAPE.OVAL)
    add_rect(slide, Inches(10.8), Inches(-2.5), Inches(5.5), Inches(5.5), pal["accent"], MSO_SHAPE.OVAL)
    
    title_text = str(content.get("title") or s.topic)
    subtitle_text = str(content.get("subtitle") or "")
    intro_text = str(content.get("intro") or "")
    
    # Если есть картинка — заголовок слева, картинка справа
    if img:
        text_area_w = Inches(7.0)
        # БОЛЬШОЙ заголовок
        title_size = fit_font_size(title_text, 7.0, 180, 54, min_size=28)
        add_text_box(slide, Inches(0.8), Inches(1.2), text_area_w, Inches(2.8),
                     title_text, title_size, pal["text"], bold=True, font=pal["hfont"], line_spacing=1.1)
        
        # Акцентная линия под заголовком
        add_rect(slide, Inches(0.85), Inches(4.1), Inches(3.2), Inches(0.08), pal["accent"])
        
        # Подзаголовок
        if subtitle_text:
            sub_size = fit_font_size(subtitle_text, 7.0, 50, 22, min_size=14)
            add_text_box(slide, Inches(0.85), Inches(4.35), text_area_w, Inches(0.7),
                         subtitle_text, sub_size, pal["muted"], font=pal["bfont"])
        
        # Интро
        if intro_text:
            intro_truncated = truncate_text(intro_text, 200)
            intro_size = fit_font_size(intro_truncated, 7.0, 80, 18, min_size=12)
            add_text_box(slide, Inches(0.85), Inches(5.1), text_area_w, Inches(1.2),
                         intro_truncated, intro_size, pal["accent"], italic=True, font=pal["bfont"])
        
        # Картинка справа (большая!)
        pic_x = Inches(8.2)
        pic_y = Inches(1.0)
        pic_w = Inches(4.6)
        pic_h = Inches(5.2)
        # Рамка вокруг картинки
        add_rect(slide, pic_x - Inches(0.06), pic_y - Inches(0.06), pic_w + Inches(0.12), pic_h + Inches(0.12), pal["accent"])
        slide.shapes.add_picture(io.BytesIO(img), pic_x, pic_y, width=pic_w, height=pic_h)
    else:
        # Без картинки — заголовок по центру на всю ширину
        title_size = fit_font_size(title_text, 11.5, 200, 60, min_size=32)
        add_text_box(slide, Inches(0.9), Inches(1.5), Inches(11.5), Inches(3.0),
                     title_text, title_size, pal["text"], bold=True, align=PP_ALIGN.CENTER, font=pal["hfont"], line_spacing=1.1)
        
        add_rect(slide, Inches(5.0), Inches(4.6), Inches(3.3), Inches(0.08), pal["accent"])
        
        if subtitle_text:
            sub_size = fit_font_size(subtitle_text, 11.5, 50, 24, min_size=16)
            add_text_box(slide, Inches(0.9), Inches(4.85), Inches(11.5), Inches(0.8),
                         subtitle_text, sub_size, pal["muted"], align=PP_ALIGN.CENTER, font=pal["bfont"])
        
        if intro_text:
            intro_truncated = truncate_text(intro_text, 250)
            intro_size = fit_font_size(intro_truncated, 11.5, 80, 20, min_size=14)
            add_text_box(slide, Inches(1.5), Inches(5.7), Inches(10.3), Inches(1.0),
                         intro_truncated, intro_size, pal["accent"], italic=True, align=PP_ALIGN.CENTER, font=pal["bfont"])
    
    # Дата внизу
    add_text_box(slide, Inches(0.5), Inches(7.05), Inches(5.0), Inches(0.35),
                 f"📅 {date_str}", 11, pal["muted"], font=pal["bfont"])


# --- ОГЛАВЛЕНИЕ (компактное) ---
def add_agenda_slide(prs, titles, pal, date_str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["bg"])
    add_rect(slide, 0, 0, Inches(0.15), prs.slide_height, pal["accent2"])
    add_rect(slide, 0, 0, prs.slide_width, Inches(0.1), pal["accent"])
    
    add_text_box(slide, Inches(0.6), Inches(0.3), Inches(12.1), Inches(0.8),
                 "Содержание", 36, pal["text"], bold=True, font=pal["hfont"])
    add_rect(slide, Inches(0.62), Inches(1.15), Inches(2.4), Inches(0.06), pal["accent"])
    
    display_titles = titles[:10]
    half = (len(display_titles) + 1) // 2
    col1 = display_titles[:half]
    col2 = display_titles[half:]
    
    # Колонка 1
    tb1 = slide.shapes.add_textbox(Inches(0.7), Inches(1.5), Inches(5.8), Inches(5.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    for i, t in enumerate(col1):
        p = tf1.paragraphs[0] if i == 0 else tf1.add_paragraph()
        p.space_after = Pt(12)
        r1 = p.add_run()
        r1.text = f"{i+1:02d}  "
        r1.font.size = Pt(18)
        r1.font.bold = True
        r1.font.color.rgb = pal["accent"]
        r1.font.name = pal["hfont"]
        r2 = p.add_run()
        r2.text = truncate_text(str(t), 45)
        r2.font.size = Pt(18)
        r2.font.color.rgb = pal["text"]
        r2.font.name = pal["bfont"]
    
    # Колонка 2
    if col2:
        tb2 = slide.shapes.add_textbox(Inches(6.8), Inches(1.5), Inches(5.8), Inches(5.5))
        tf2 = tb2.text_frame
        tf2.word_wrap = True
        for i, t in enumerate(col2):
            p = tf2.paragraphs[0] if i == 0 else tf2.add_paragraph()
            p.space_after = Pt(12)
            r1 = p.add_run()
            r1.text = f"{half+i+1:02d}  "
            r1.font.size = Pt(18)
            r1.font.bold = True
            r1.font.color.rgb = pal["accent"]
            r1.font.name = pal["hfont"]
            r2 = p.add_run()
            r2.text = truncate_text(str(t), 45)
            r2.font.size = Pt(18)
            r2.font.color.rgb = pal["text"]
            r2.font.name = pal["bfont"]
    
    add_text_box(slide, Inches(0.5), Inches(7.05), Inches(5.0), Inches(0.35),
                 f"📅 {date_str}", 11, pal["muted"], font=pal["bfont"])


# --- КОНТЕНТНЫЙ СЛАЙД (главный тип, много информации) ---
def add_content_slide(prs, item, pal, num, total, img, date_str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["bg"])
    
    # Левая акцентная полоса
    add_rect(slide, 0, 0, Inches(0.15), prs.slide_height, pal["accent2"])
    # Верхняя тонкая линия
    add_rect(slide, 0, 0, prs.slide_width, Inches(0.08), pal["accent"])
    
    title_text = str(item.get("title") or f"Слайд {num}")
    lead_text = str(item.get("lead") or "")
    bullets_raw = clean_bullets(item.get("bullets"), 4)
    scheme_raw = item.get("scheme") or []
    fact_text = str(item.get("fact") or "")
    
    # === ЗАГОЛОВОК (крупный, нумерованный) ===
    num_str = f"{num - 2:02d}"
    header_tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(12.5), Inches(0.95))
    hdr_tf = header_tb.text_frame
    hdr_tf.word_wrap = True
    hp = hdr_tf.paragraphs[0]
    hr1 = hp.add_run()
    hr1.text = f"{num_str}  "
    hr1.font.size = Pt(36)
    hr1.font.bold = True
    hr1.font.color.rgb = pal["accent"]
    hr1.font.name = pal["hfont"]
    hr2 = hp.add_run()
    hr2.text = truncate_text(title_text, 60)
    hr2.font.size = Pt(36)
    hr2.font.bold = True
    hr2.font.color.rgb = pal["text"]
    hr2.font.name = pal["hfont"]
    
    # Линия под заголовком
    add_rect(slide, Inches(0.55), Inches(1.18), Inches(2.8), Inches(0.06), pal["accent"])
    
    # === ЛЕЙД (подзаголовок-введение к слайду) ===
    current_y = Inches(1.35)
    if lead_text:
        lead_truncated = truncate_text(lead_text, 180)
        lead_size = fit_font_size(lead_truncated, 12.0, 35, 17, min_size=13)
        add_text_box(slide, Inches(0.55), current_y, Inches(12.0), Inches(0.5),
                     lead_truncated, lead_size, pal["muted"], italic=True, font=pal["bfont"])
        current_y = Inches(1.85)
    
    # === ОСНОВНАЯ ОБЛАСТЬ: буллеты + схема + картинка ===
    
    # Определяем размеры зон
    if img:
        # Картинка справа сверху
        bullets_w = Inches(7.8)
        pic_x = Inches(8.6)
        pic_y = current_y
        pic_w = Inches(4.2)
        pic_h = Inches(3.0)
    else:
        bullets_w = Inches(12.0)
        pic_x = None
    
    # Буллеты
    bullet_count = len(bullets_raw)
    if bullet_count > 0:
        bullets_available_h = Inches(3.2) if img else Inches(4.5)
        per_bullet_h = bullets_available_h / max(bullet_count, 1)
        bullet_font_size = min(17, max(13, int(per_bullet_h.pt / 3.5)))
        
        bt_tb = slide.shapes.add_textbox(Inches(0.55), current_y, bullets_w, bullets_available_h)
        bt_tf = bt_tb.text_frame
        bt_tf.word_wrap = True
        
        for bi, bullet in enumerate(bullets_raw):
            bp = bt_tf.paragraphs[0] if bi == 0 else bt_tf.add_paragraph()
            bp.space_after = Pt(8)
            bp.line_spacing = 1.2
            
            br1 = bp.add_run()
            br1.text = "▸ "
            br1.font.size = Pt(bullet_font_size)
            br1.font.bold = True
            br1.font.color.rgb = pal["accent"]
            br1.font.name = pal["bfont"]
            
            br2 = bp.add_run()
            # Обрезаем каждый буллет до разумной длины
            max_bullet_chars = 160 if img else 200
            br2.text = truncate_text(bullet, max_bullet_chars)
            br2.font.size = Pt(bullet_font_size)
            br2.font.color.rgb = pal["text"]
            br2.font.name = pal["bfont"]
    
    # Картинка (если есть)
    if img and pic_x:
        # Белая рамка-фон
        add_rect(slide, pic_x - Inches(0.05), pic_y - Inches(0.05), pic_w + Inches(0.1), pic_h + Inches(0.1), pal["panel"])
        # Акцентная граница
        add_rect(slide, pic_x - Inches(0.03), pic_y - Inches(0.03), pic_w + Inches(0.06), pic_h + Inches(0.06), pal["accent"])
        slide.shapes.add_picture(io.BytesIO(img), pic_x, pic_y, width=pic_w, height=pic_h)
    
    # === СХЕМА (нижняя часть слайда) ===
    scheme_y = Inches(5.0)
    
    # ИСПРАВЛЕНО: Проверяем что схема содержит реальные данные
    valid_scheme = []
    for si in (scheme_raw or []):
        s_clean = str(si).strip()
        # Пропускаем заглушки типа "Шаг 1..."
        if len(s_clean) >= 5 and not s_clean.startswith("Шаг"):
            valid_scheme.append(truncate_text(s_clean, 40))
    
    if valid_scheme and len(valid_scheme) >= 2:
        n_boxes = min(len(valid_scheme), 4)
        box_w = Inches(2.6)
        gap = Inches(0.5)
        arrow_w = Inches(0.35)
        total_scheme_w = n_boxes * box_w + (n_boxes - 1) * (gap + arrow_w)
        start_x = int((Inches(13.333) - total_scheme_w) / 2)
        
        for si, step_text in enumerate(valid_scheme[:4]):
            bx = start_x + si * (box_w + gap + arrow_w)
            
            # Rounded rectangle box
            box = add_rounded_rect(slide, bx, scheme_y, box_w, Inches(1.1),
                                   pal["panel"], pal["accent"], Pt(1.5))
            box_tf = box.text_frame
            box_tf.word_wrap = True
            box_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            box_tf.margin_left = Inches(0.1)
            box_tf.margin_right = Inches(0.1)
            box_tf.margin_top = Inches(0.05)
            box_tf.margin_bottom = Inches(0.05)
            
            # Номер шага
            step_num_p = box_tf.paragraphs[0]
            step_num_r = step_num_p.add_run()
            step_num_r.text = f"{si+1}. "
            step_num_r.font.size = Pt(14)
            step_num_r.font.bold = True
            step_num_r.font.color.rgb = pal["accent"]
            step_num_r.font.name = pal["bfont"]
            
            # Текст шага
            step_text_p = box_tf.add_paragraph()
            step_text_r = step_text_p.add_run()
            step_text_r.text = step_text
            step_text_r.font.size = Pt(13)
            step_text_r.font.bold = True
            step_text_r.font.color.rgb = pal["text"]
            step_text_r.font.name = pal["bfont"]
            step_text_p.alignment = PP_ALIGN.CENTER
            
            # Стрелка между блоками
            if si < n_boxes - 1:
                arrow_x = bx + box_w + Inches(0.05)
                arrow_y_pos = scheme_y + Inches(0.4)
                add_rect(slide, arrow_x, arrow_y_pos, arrow_w, Inches(0.25),
                         pal["accent"], MSO_SHAPE.RIGHT_ARROW)
    else:
        # Если нет валидной схемы — показываем факт крупнее
        if fact_text:
            fact_card_y = scheme_y
            add_rounded_rect(slide, Inches(0.5), fact_card_y, Inches(12.3), Inches(1.0),
                             pal["panel"], pal["accent2"], Pt(1.5))
            add_text_box(slide, Inches(0.8), fact_card_y + Inches(0.15), Inches(11.7), Inches(0.7),
                         "💡 " + truncate_text(fact_text, 200), 16, pal["text"], italic=True, font=pal["bfont"])
    
    # === ФАКТ (если схема заняла место — ставим компактно над схемой) ===
    if fact_text and valid_scheme and len(valid_scheme) >= 2:
        fact_bar_y = Inches(4.55)
        add_rect(slide, Inches(0.5), fact_bar_y, Inches(12.3), Inches(0.35), pal["panel"], MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text_box(slide, Inches(0.7), fact_bar_y + Inches(0.03), Inches(11.9), Inches(0.3),
                     "💡 " + truncate_text(fact_text, 150), 13, pal["accent"], italic=True, font=pal["bfont"])
    
    # === НОМЕР СТРАНИЦЫ ===
    add_text_box(slide, Inches(11.9), Inches(7.1), Inches(1.3), Inches(0.3),
                 f"{num} / {total}", 11, pal["muted"], align=PP_ALIGN.RIGHT, font=pal["bfont"])
    
    # Заметки докладчика
    notes = slide.notes_slide.notes_text_frame
    speaker = str(item.get("speaker_notes") or "")
    notes.text = speaker if speaker else f"Слайд {num}: {title_text}"


# --- СЛАЙД СТАТИСТИКИ ---
def add_stats_slide(prs, stats, pal, num, total, date_str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["bg"])
    add_rect(slide, 0, 0, Inches(0.15), prs.slide_height, pal["accent2"])
    add_rect(slide, 0, 0, prs.slide_width, Inches(0.08), pal["accent"])
    
    add_text_box(slide, Inches(0.6), Inches(0.3), Inches(12.1), Inches(0.8),
                 "Ключевые цифры", 36, pal["text"], bold=True, font=pal["hfont"])
    add_rect(slide, Inches(0.62), Inches(1.15), Inches(2.4), Inches(0.06), pal["accent"])
    
    n = min(len(stats), 3)
    card_w = Inches(3.6)
    gap = Inches(0.5)
    total_w = n * card_w + (n - 1) * gap
    start_x = int((prs.slide_width - total_w) / 2)
    
    for i, st in enumerate(stats[:3]):
        cx = start_x + i * (card_w + gap)
        cy = Inches(1.8)
        ch = Inches(3.5)
        
        add_rounded_rect(slide, cx, cy, card_w, ch, pal["panel"], pal["accent"], Pt(2))
        
        if isinstance(st, dict):
            value = str(st.get("value", "?"))
            label = str(st.get("label", ""))
        else:
            value = str(st)
            label = ""
        
        val_size = fit_font_size(value, 3.2, 80, 52, min_size=28)
        add_text_box(slide, cx + Inches(0.2), cy + Inches(0.5), card_w - Inches(0.4), Inches(1.5),
                     value, val_size, pal["accent"], bold=True, align=PP_ALIGN.CENTER, font=pal["hfont"])
        
        lbl_size = fit_font_size(label, 3.2, 60, 18, min_size=12)
        add_text_box(slide, cx + Inches(0.2), cy + Inches(2.2), card_w - Inches(0.4), Inches(1.0),
                     label, lbl_size, pal["muted"], align=PP_ALIGN.CENTER, font=pal["bfont"])
    
    add_text_box(slide, Inches(11.9), Inches(7.1), Inches(1.3), Inches(0.3),
                 f"{num} / {total}", 11, pal["muted"], align=PP_ALIGN.RIGHT, font=pal["bfont"])
    add_text_box(slide, Inches(0.5), Inches(7.05), Inches(5.0), Inches(0.35),
                 f"📅 {date_str}", 11, pal["muted"], font=pal["bfont"])


# --- СЛАЙД ДИАГРАММЫ ---
def add_chart_slide(prs, chart, rows, pal, num, total, date_str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["bg"])
    add_rect(slide, 0, 0, Inches(0.15), prs.slide_height, pal["accent2"])
    add_rect(slide, 0, 0, prs.slide_width, Inches(0.08), pal["accent"])
    
    chart_title = str(chart.get("title") or "Данные")
    add_text_box(slide, Inches(0.6), Inches(0.3), Inches(12.1), Inches(0.8),
                 chart_title, 36, pal["text"], bold=True, font=pal["hfont"])
    add_rect(slide, Inches(0.62), Inches(1.15), Inches(2.4), Inches(0.06), pal["accent"])
    
    cd = CategoryChartData()
    categories = [truncate_text(r[0], 30) for r in rows]
    values_list = [r[1] for r in rows]
    cd.categories = categories
    cd.add_series("Значения", values_list)
    
    ctype = XL_CHART_TYPE.PIE
    if str(chart.get("type", "")).lower() == "bar":
        ctype = XL_CHART_TYPE.BAR_CLUSTERED
    
    gf = slide.shapes.add_chart(ctype, Inches(1.2), Inches(1.5), Inches(10.9), Inches(5.3), cd)
    ch = gf.chart
    try:
        ch.has_legend = True
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
        ch.font.size = Pt(13)
        ch.font.color.rgb = pal["text"]
        plot = ch.plots[0]
        plot.has_data_labels = True
        plot.data_labels.font.size = Pt(12)
        plot.data_labels.font.color.rgb = pal["text"]
        if ctype == XL_CHART_TYPE.PIE:
            colors_cycle = (pal["accent"], pal["accent2"], pal["muted"], 
                          pal["accent"], pal["accent2"], pal["muted"])
            for pi, point in enumerate(plot.series[0].points):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = colors_cycle[pi % len(colors_cycle)]
    except Exception as e:
        log.warning("Chart styling skipped: %s", e)
    
    add_text_box(slide, Inches(11.9), Inches(7.1), Inches(1.3), Inches(0.3),
                 f"{num} / {total}", 11, pal["muted"], align=PP_ALIGN.RIGHT, font=pal["bfont"])
    add_text_box(slide, Inches(0.5), Inches(7.05), Inches(5.0), Inches(0.35),
                 f"📅 {date_str}", 11, pal["muted"], font=pal["bfont"])


# --- ФИНАЛЬНЫЙ СЛАЙД ---
def add_final_slide(prs, content, pal, num, total, date_str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["accent2"])
    
    add_rect(slide, Inches(-1.5), Inches(-2.0), Inches(5.0), Inches(5.0), pal["accent"], MSO_SHAPE.OVAL)
    add_rect(slide, Inches(10.0), Inches(5.0), Inches(5.0), Inches(5.0), pal["bg"], MSO_SHAPE.OVAL)
    
    conclusion = str(content.get("conclusion") or "Спасибо за внимание!")
    concl_truncated = truncate_text(conclusion, 300)
    concl_size = fit_font_size(concl_truncated, 10.0, 120, 26, min_size=18)
    
    add_text_box(slide, Inches(1.5), Inches(1.8), Inches(10.3), Inches(2.0),
                 concl_truncated, concl_size, WHITE, align=PP_ALIGN.CENTER, font=pal["bfont"], line_spacing=1.3)
    
    add_text_box(slide, Inches(1.5), Inches(4.0), Inches(10.3), Inches(1.5),
                 "СПАСИБО ЗА ВНИМАНИЕ!", 52, WHITE, bold=True, align=PP_ALIGN.CENTER, font=pal["hfont"])
    
    add_text_box(slide, Inches(11.9), Inches(7.1), Inches(1.3), Inches(0.3),
                 f"{num} / {total}", 11, WHITE, align=PP_ALIGN.RIGHT, font=pal["bfont"])
    add_text_box(slide, Inches(0.5), Inches(7.05), Inches(5.0), Inches(0.35),
                 f"📅 {date_str}", 11, WHITE, font=pal["bfont"])
    
    notes = slide.notes_slide.notes_text_frame
    notes.text = "Финальный слайд. Поблагодарите аудиторию и предложите задать вопросы."


def clean_bullets(raw, limit=4):
    out = []
    for b in raw or []:
        t = str(b).strip()
        t = t.strip('"\'""«»`·—–')
        if len(t) < 5:
            continue
        out.append(t[:250])
        if len(out) >= limit:
            break
    return out


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
        if label and value != 0:
            rows.append((label, abs(value)))
        if len(rows) >= 6:
            break
    return rows if len(rows) >= 2 else []


def build_pptx(content, s, slide_imgs, title_img, quality=85):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = str(content.get("title") or s.topic)
    prs.core_properties.subject = s.topic
    prs.core_properties.author = "Presentation Bot v23.1"
    prs.core_properties.created = datetime.now()
    date_str = datetime.now().strftime("%d.%m.%Y")
    
    key = resolve_style(s)
    pal = dict(PALETTES.get(key, PALETTES["dark"]))
    hfont, bfont = THEME_FONTS.get(key, ("Segoe UI", "Calibri"))
    pal["hfont"] = hfont
    pal["bfont"] = bfont
    
    slides = content.get("slides", [])[:s.count]
    stats = content.get("stats") or []
    has_stats = isinstance(stats, list) and bool(stats)
    rows = chart_rows(content.get("chart"))
    has_chart = bool(rows)
    
    total = 2 + len(slides) + (1 if has_stats else 0) + (1 if has_chart else 0) + 1
    
    add_title_slide(prs, content, pal, s, title_img, date_str)
    add_agenda_slide(prs, [it.get("title", "") for it in slides], pal, date_str)
    
    num = 3
    for i, item in enumerate(slides):
        img = slide_imgs[i] if i < len(slide_imgs) else None
        add_content_slide(prs, item, pal, num, total, img, date_str)
        num += 1
    
    if has_stats:
        add_stats_slide(prs, stats, pal, num, total, date_str)
        num += 1
    if has_chart:
        add_chart_slide(prs, content.get("chart"), rows, pal, num, total, date_str)
        num += 1
    
    add_final_slide(prs, content, pal, num, total, date_str)
    
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def edit_keyboard(count):
    rows = []
    row = []
    for i in range(1, min(count, 8) + 1):
        row.append(InlineKeyboardButton(text=str(i), callback_data=f"edit:{i}"))
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


STAGE_MESSAGES = [
    "🧠 Пишу текст презентации…",
    "🖼 Ищу подходящие фотографии…",
    "🎨 Оформляю слайды…",
    "📦 Готовлю файл…",
    "✅ Презентация готова!",
]


async def send_stage(status_msg, stage_idx):
    msg = STAGE_MESSAGES[min(stage_idx, len(STAGE_MESSAGES) - 1)]
    try:
        await status_msg.edit_text(msg)
    except Exception:
        pass


async def generate_simple_flow(update, context, s):
    chat_id = update.effective_chat.id
    status_msg = update.effective_message
    s.busy = True
    t0 = time.monotonic()
    
    try:
        # Этап 1: Текст
        await send_stage(status_msg, 0)
        source = "AI"
        wiki_thumb = None
        
        try:
            content = await ask_ai_multi(s)
        except Exception as e:
            log.warning(f"AI generation failed, using Wikipedia fallback: {e}")
            try:
                content, wiki_thumb = await smart_fallback_from_wikipedia(s)
                source = "wiki"
            except Exception as we:
                log.error(f"Wikipedia fallback also failed: {we}")
                # Последний шанс — минимальная заглушка
                content = {
                    "title": s.topic,
                    "subtitle": "Автоматически сгенерированная презентация",
                    "img": s.topic,
                    "intro": f"Презентация по теме «{s.topic}».",
                    "slides": [
                        {"title": "Введение", "lead": f"Знакомство с темой {s.topic}", "img": s.topic,
                         "bullets": ["Актуальность данной темы", "Основные понятия и определения",
                                     "Исторический контекст вопроса", "Цели и задачи рассмотрения"],
                         "scheme": ["Определение термина", "Контекст проблемы", "Актуальность сегодня"],
                         "fact": "", "speaker_notes": "Представьте тему аудитории."},
                        {"title": "Ключевые аспекты", "lead": "Рассмотрение основных сторон вопроса", "img": f"{s.topic} details",
                         "bullets": ["Первый важный аспект требует детального изучения",
                                     "Второй компонент играет значительную роль в общей картине",
                                     "Третий фактор определяет направление дальнейшего развития",
                                     "Все элементы взаимосвязаны и образуют целостную систему"],
                         "scheme": ["Анализ компонентов", "Взаимосвязи элементов", "Системный подход"],
                         "fact": "", "speaker_notes": "Раскройте каждый аспект подробнее."},
                        {"title": "Практическое значение", "lead": "Где это применяется в жизни", "img": f"{s.topic} application",
                         "bullets": ["Практическое внедрение даёт измеримые результаты",
                                     "Реальные кейсы подтверждают эффективность подхода",
                                     "Экономический эффект оценивается в миллиардах ежегодно",
                                     "Социальное влияние распространяется на миллионы людей"],
                         "scheme": ["Теория к практике", "Пилотные проекты", "Масштабирование"],
                         "fact": "", "speaker_notes": "Приведите конкретные примеры."},
                        {"title": "Выводы", "lead": f"Итоги рассмотрения темы «{s.topic}»", "img": s.topic,
                         "bullets": ["Тема обладает высоким потенциалом для развития",
                                     "Необходим дальнейший мониторинг тенденций",
                                     "Междисциплинарный подход усиливает понимание",
                                     "Каждый может внести вклад в развитие области"],
                         "scheme": ["Обобщение знаний", "Оценка перспектив", "Призыв к действию"],
                         "fact": "", "speaker_notes": "Подведите итог и поблагодарите."},
                    ],
                    "stats": [{"value": "4", "label": "раздела"}, {"value": "16", "label": "тезисов"}, {"value": "∞", "label": "возможностей"}],
                    "chart": {"title": "Структура презентации", "type": "bar",
                              "data": [{"label": "Введение", "value": 4}, {"label": "Аспекты", "value": 4},
                                       {"label": "Практика", "value": 4}, {"label": "Выводы", "value": 4}]},
                    "conclusion": f"Мы рассмотрели ключевые аспекты темы «{s.topic}». Спасибо за внимание!"
                }
                source = "stub"
        
        s.last_content = content
        slides = content.get("slides", [])[:s.count]
        
        # Этап 2: Фото (параллельно, с увеличенным лимитом)
        await send_stage(status_msg, 1)
        queries = [str(it.get("img") or it.get("title") or s.topic) for it in slides]
        slide_imgs = await fetch_images_parallel(queries, limit=6, quality=85)
        
        title_img = None
        if wiki_thumb:
            try:
                async with aiohttp.ClientSession(headers=HTTP_HEADERS, timeout=aiohttp.ClientTimeout(total=8)) as sess:
                    title_img = await _download_image(sess, wiki_thumb, 85)
            except Exception:
                pass
        
        if not title_img:
            try:
                res = await fetch_images_parallel([str(content.get("img") or s.topic)], limit=1, quality=85)
                title_img = res[0] if res else None
            except Exception:
                pass
        
        pics_loaded = sum(1 for x in slide_imgs if x) + (1 if title_img else 0)
        s.last_imgs = slide_imgs
        s.last_title_img = title_img
        
        # Этап 3: Сборка
        await send_stage(status_msg, 2)
        data = build_pptx(content, s, slide_imgs, title_img)
        
        file_size_mb = len(data) / (1024 * 1024)
        if file_size_mb > 7:
            log.info(f"Recompressing images (was {file_size_mb:.1f}MB)")
            quality_low = 55
            slide_imgs_recomp = await fetch_images_parallel(queries, limit=6, quality=quality_low)
            title_img_recomp = None
            try:
                res = await fetch_images_parallel([str(content.get("img") or s.topic)], limit=1, quality=quality_low)
                title_img_recomp = res[0] if res else None
            except Exception:
                pass
            s.last_imgs = slide_imgs_recomp
            s.last_title_img = title_img_recomp
            data = build_pptx(content, s, slide_imgs_recomp, title_img_recomp)
            file_size_mb = len(data) / (1024 * 1024)
        
        # Этап 4: Отправка
        await send_stage(status_msg, 3)
        name = re.sub(r"[^\w\- ]+", "", s.topic).strip()[:40]
        name = name or "presentation"
        pal_name = PALETTE_NAMES.get(resolve_style(s), "")
        secs = int(time.monotonic() - t0)
        total_words = sum(len(str(b).split()) for sl in slides for b in sl.get("bullets", []))
        
        caption_parts = [f"✅ Готово за {secs}с!"]
        caption_parts.append(f"🎨 {pal_name} | 🖼 {pics_loaded} фото | 📝 {total_words} слов")
        caption_parts.append(f"v{VER} | {len(slides)} контента + оформление | {file_size_mb:.1f} MB")
        if source == "wiki":
            caption_parts.append("⚠️ AI недоступен → Википедия")
        elif source == "stub":
            caption_parts.append("⚠️ Резервная генерация")
        caption = "\n".join(caption_parts)
        
        await context.bot.send_document(
            chat_id=chat_id,
            document=InputFile(io.BytesIO(data), filename=f"{name}.pptx"),
            caption=caption,
        )
        
        # ✅ ИСПРАВЛЕНО: Безопасное создание клавиатуры редактирования
        kb_rows = []
        btn_row = []
        for i in range(1, min(len(slides), 8) + 1):
            btn_row.append(InlineKeyboardButton(text=str(i), callback_data=f"edit:{i}"))
            if len(btn_row) == 4:
                kb_rows.append(btn_row)
                btn_row = []
        if btn_row:
            kb_rows.append(btn_row)
        
        # Добавляем кнопку перегенерации отдельным рядом
        kb_rows.append([InlineKeyboardButton("🔄 Перегенерировать", callback_data="regen")])
        
        final_kb = InlineKeyboardMarkup(kb_rows)
        
        await context.bot.send_message(
            chat_id,
            "✏️ Нажми номер слайда и напиши новый заголовок:",
            reply_markup=final_kb,
        )
        
        await send_stage(status_msg, 4)
        
    except Exception as e:
        log.exception("generation failed completely")
        reason = f"{type(e).__name__}: {e}"
        if len(reason) > 300:
            reason = reason[:300] + "..."
        try:
            await status_msg.edit_text(
                f"😔 Не получилось создать презентацию.\n"
                f"Причина: {reason}\n\n"
                f"Попробуй ещё раз: /start"
            )
        except Exception:
            pass
    finally:
        s.busy = False


async def health(_request):
    return web.Response(text="OK")


async def main():
    if not BOT_TOKEN:
        log.error("Set BOT_TOKEN environment variable!")
        return
    if not GIGACHAT_CLIENT_ID or not GIGACHAT_AUTH_KEY:
        log.error("GIGACHAT_CLIENT_ID and GIGACHAT_AUTH_KEY must both be set!")
        return

    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help", cmd_help))
    app.add_handler(CommandHandler("content", cmd_content))
    app.add_handler(CommandHandler("cancel", cmd_cancel))
    app.add_handler(CallbackQueryHandler(on_callback))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_message))

    await app.initialize()
    await app.start()
    await app.updater.start_polling(drop_pending_updates=True)
    log.info("Bot started (polling), version %s", VER)
    log.info("Using GigaChat-Pro model via Sber Devices API")

    webapp = web.Application()
    webapp.router.add_get("/", health)
    webapp.router.add_get("/health", health)
    runner = web.AppRunner(webapp)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", PORT)
    await site.start()
    log.info("Health check server on port %s", PORT)

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