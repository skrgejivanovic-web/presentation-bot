import asyncio
import io
import json
import logging
import os
import re
import time
from datetime import datetime
import base64  # <-- ВАЖНО: добавлено для авторизации GigaChat

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

VER = "20.1"

# --- Настройки окружения ---
BOT_TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_TOKEN", "")
GIGACHAT_CLIENT_ID = os.getenv("GIGACHAT_CLIENT_ID", "")
GIGACHAT_AUTH_KEY = os.getenv("GIGACHAT_AUTH_KEY", "") 

PORT = int(os.getenv("PORT", "10000"))

# URL для получения токена доступа (OAuth)
GIGACHAT_OAUTH_URL = "https://ngrok.devices.sberbank.ru/api/v1/oauth/token" 
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


# ИСПРАВЛЕНО: Безопасная клавиатура настроек
def settings_keyboard(s):
    rows = []
    # Выбор количества слайдов
    for row in (COUNTS[:3], COUNTS[3:]):
        rows.append([InlineKeyboardButton(mark(s.count == n, str(n)), callback_data=f"n:{n}") for n in row])
    
    # Язык
    rows.append([
        InlineKeyboardButton(mark(s.lang == "ru", "🇷🇺 Русский"), callback_data="lang:ru"),
        InlineKeyboardButton(mark(s.lang == "en", "🇬🇧 English"), callback_data="lang:en"),
    ])
    
    # Стиль оформления
    rows.append([
        InlineKeyboardButton(mark(s.style == "auto", "🪄 Авто"), callback_data="style:auto"),
        InlineKeyboardButton(mark(s.style == "dark", "🌙 Тёмная"), callback_data="style:dark"),
        InlineKeyboardButton(mark(s.style == "light", "☀️ Светлая"), callback_data="style:light"),
        InlineKeyboardButton(mark(s.style == "color", "🎨 Яркая"), callback_data="style:color"),
    ])
    
    # История тем (используем индексы, чтобы не было ошибок Button_data_invalid)
    if s.history:
        hist_row = []
        # Берем последние 3 темы из истории
        recent_topics = s.history[-3:]
        for i, h in enumerate(recent_topics, start=1):
            short_name = h[:12] + "…" if len(h) > 12 else h
            # Передаем только номер (hidx:1, hidx:2...), это безопасно для Telegram
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
        text = text[:4000] + "…"
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
            s.history.pop(0) # Удаляем самую старую
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

    # ИСПРАВЛЕНО: Обработка истории через индексы
    elif data.startswith("hidx:"):
        try:
            idx = int(data.split(":", 1)[1]) - 1
            # Проверяем, что индекс корректный относительно последних 3 элементов
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
    "You are a top-tier presentation designer and writer. "
    "Return ONLY valid JSON, no markdown, no explanations, "
    'with keys: "title" (string), "subtitle" (string), '
    '"img" (3-5 ENGLISH words describing a photo for the title slide), '
    '"intro" (1-2 catchy sentences), '
    '"slides" (array of EXACTLY {count} objects - you MUST create exactly '
    '{count} slides, no more, no less, each with: "title" '
    '(string), "lead" (one engaging sentence), '
    '"img" (3-5 ENGLISH words with the main subject first), '
    '"bullets" (array of 4 NON-EMPTY informative strings, each 8-18 '
    'words), '
    '"fact" (one surprising fact or witty remark), '
    '"speaker_notes" (2-3 sentences with tips for the presenter), '
    '"stats" (array of 3 objects with "value" (short number or percent) '
    'and "label" (short description)), '
    '"chart" (object with "title" (string), "type" ("pie" or "bar"), '
    '"data" (array of 4-6 objects with "label" (short string) and '
    '"value" (number; percents 0-100 when type is pie))), '
    '"conclusion" (2-3 inspiring sentences). '
    "Write EVERY string strictly in {lang}, never mix languages. "
    "CRITICAL: You MUST create EXACTLY {count} slides. Count them before "
    "returning. Topic: {topic}"
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
            async with sess.post(GIGACHAT_OAUTH_URL, json=payload, headers=headers) as r:
                if r.status != 200:
                    log.error(f"GigaChat OAuth failed: {await r.text()}")
                    raise ValueError("Failed to get access token")
                data = await r.json()
                return data.get("access_token")
        except Exception as e:
            log.error(f"GigaChat OAuth error: {e}")
            raise


async def ask_gigachat(prompt_text):
    """Отправляем запрос к GigaChat напрямую через aiohttp"""
    access_token = await get_gigachat_access_token()
    
    messages = [
        {"role": "system", "content": "You are a helpful assistant that generates presentations."},
        {"role": "user", "content": prompt_text}
    ]
    
    payload = {
        "model": "GigaChat-Max",
        "messages": messages,
        "temperature": 0.7
    }
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    timeout = aiohttp.ClientTimeout(total=60) 
    
    async with aiohttp.ClientSession(timeout=timeout) as sess:
        try:
            async with sess.post(REAL_GIGACHAT_API_URL, json=payload, headers=headers) as r:
                if r.status != 200:
                    err_body = await r.text()
                    log.error(f"GigaChat API Error {r.status}: {err_body}")
                    raise RuntimeError(f"GigaChat returned {r.status}")
                
                resp_data = await r.json()
                content = resp_data.get("choices", [{}])[0].get("message", {}).get("content", "")
                return content
                
        except asyncio.TimeoutError:
            raise TimeoutError("GigaChat timed out after 60s")
        except Exception as e:
            log.exception("Unexpected error calling GigaChat")
            raise RuntimeError(f"GigaChat call failed: {e}")


async def ask_ai_multi(s):
    lang = "Russian" if s.lang == "ru" else "English"
    prompt = PROMPT.format(lang=lang, count=s.count, topic=s.topic)
    
    try:
        raw_response = await ask_gigachat(prompt)
        json_str = extract_json(raw_response)
        
        if not json_str:
            raise ValueError("No JSON found in response")
            
        content = json.loads(json_str)
        slides = content.get("slides", [])
        
        if not isinstance(slides, list) or len(slides) < s.count:
             log.warning(f"Got {len(slides)} slides, need {s.count}. Trying again...")
             raw_response = await ask_gigachat(prompt)
             json_str = extract_json(raw_response)
             content = json.loads(json_str)
             
        return content
        
    except Exception as e:
        log.error(f"Main AI flow failed: {e}")
        raise e


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
        "exintro": "1",
        "explaintext": "1",
        "piprop": "thumbnail",
        "pithumbsize": "900",
    }
    timeout = aiohttp.ClientTimeout(total=10)
    async with aiohttp.ClientSession(timeout=timeout, headers=HTTP_HEADERS) as sess:
        async with sess.get(api, params=params) as r:
            data = await r.json(content_type=None)
    pages = (data.get("query") or {}).get("pages") or {}
    if not pages:
        raise ValueError("no Wikipedia article found")
    page = list(pages.values())[0]
    article_title = page.get("title", s.topic)
    extract = (page.get("extract") or "").strip()
    thumbnail_url = (page.get("thumbnail") or {}).get("source")
    sentences = [x.strip() for x in re.split(r"(?<=[.!?])\s+", extract) if x.strip()]
    slides = []
    intro_bullets = [sent for sent in sentences[:4] if 20 < len(sent) < 200]
    if not intro_bullets:
        intro_bullets = ["Основные сведения о теме", "Краткое описание предмета рассмотрения"]
    slides.append({
        "title": "Введение",
        "lead": f"Знакомство с темой: {article_title}",
        "img": article_title,
        "bullets": intro_bullets[:4],
        "fact": "",
        "speaker_notes": "Представьте тему аудитории кратко и ясно.",
    })
    detail_sentences = [sent for sent in sentences if len(sent) > 30][:20]
    chunk_size = max(2, len(detail_sentences) // (s.count - 2)) if s.count > 2 else 2
    section_names = ["Основная часть", "Детали и аспекты", "Практическое применение", 
                     "Преимущества и особенности", "Вызовы и решения", "Перспективы развития"]
    for i in range(min(s.count - 2, 5)):
        start_idx = i * chunk_size
        end_idx = start_idx + chunk_size
        chunk_sents = detail_sentences[start_idx:end_idx]
        bullets = []
        for sent in chunk_sents:
            cleaned = sent.replace("\n", " ").strip()
            if 20 < len(cleaned) < 180:
                bullets.append(cleaned)
        if not bullets:
            bullets = [f"Аспект {i+1} темы {article_title}", "Дополнительные сведения по разделу"]
        section_name = section_names[i] if i < len(section_names) else f"Раздел {i+1}"
        slides.append({
            "title": section_name,
            "lead": f"Подробное рассмотрение: {section_name.lower()} темы {article_title}",
            "img": f"{article_title} {section_name}",
            "bullets": bullets[:4],
            "fact": "",
            "speaker_notes": f"Расскажите подробнее про {section_name.lower()}.",
        })
    conclusion_bullets = []
    remaining_sents = detail_sentences[len(section_names)*chunk_size:]
    for sent in remaining_sents[:3]:
        if len(sent) > 25:
            conclusion_bullets.append(sent)
    if not conclusion_bullets:
        conclusion_bullets = [
            "Подведём итоги рассмотренной темы",
            "Запомните ключевые моменты презентации",
            "Эта информация поможет лучше понять предмет"
        ]
    slides.append({
        "title": "Заключение",
        "lead": f"Итоги по теме: {article_title}",
        "img": article_title,
        "bullets": conclusion_bullets[:4],
        "fact": "",
        "speaker_notes": "Сделайте выводы и поблагодарите аудиторию.",
    })
    stats = [
        {"value": f"{len(sentences)}", "label": "предложений найдено"},
        {"value": f"{len(slides)}", "label": "слайдов создано"},
        {"value": "Wiki", "label": "источник данных"},
    ]
    chart = {
        "title": f"Структура презентации: {article_title}",
        "type": "bar",
        "data": [
            {"label": "Введение", "value": 1},
            {"label": "Основная часть", "value": len(slides)-2},
            {"label": "Заключение", "value": 1},
        ],
    }
    result = {
        "title": article_title,
        "subtitle": f"Материалы собраны из Википедии ({datetime.now().strftime('%d.%m.%Y')})",
        "img": article_title,
        "intro": f"Презентация подготовлена автоматически на основе открытых источников о теме «{article_title}».",
        "slides": slides,
        "stats": stats,
        "chart": chart,
        "conclusion": f"Мы рассмотрели основные аспекты темы «{article_title}». Спасибо за внимание!",
    }
    log.info(f"Wikipedia fallback created {len(slides)} structured slides from '{article_title}'")
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
    img = img.resize((w, h))
    out = io.BytesIO()
    img.save(out, format="JPEG", quality=quality)
    return out.getvalue()


def _tokens(q):
    return set(re.findall(r"[a-z0-9а-яё]{4,}", (q or "").lower()))


def _relevant(query, title):
    qw = _tokens(query)
    tw = _tokens(title)
    if not qw or not tw:
        return True
    return bool(qw & tw)


async def _download_image(sess, url, quality=85):
    try:
        async with sess.get(url) as r:
            if r.status != 200:
                return None
            raw = await r.read()
        if not raw:
            return None
        return prepare_image(raw, quality=quality)
    except Exception as e:
        log.warning("image download failed %s: %s", url, e)
        return None


async def fetch_topic_image(query, quality=85):
    timeout = aiohttp.ClientTimeout(total=4, connect=2)
    async with aiohttp.ClientSession(timeout=timeout, headers=HTTP_HEADERS) as sess:
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
                    img = await _download_image(sess, url, quality)
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
                    img = await _download_image(sess, url, quality)
                    if img:
                        log.info("image from wikipedia: %s", query)
                        return img
        except Exception as e:
            log.warning("wiki image failed for %s: %s", query, e)
        try:
            async with sess.get(OPENVERSE_API, params={"q": query, "page_size": "5"}) as r:
                if r.status == 200:
                    data = await r.json(content_type=None)
                    for res in (data.get("results") or []):
                        if not _relevant(query, res.get("title")):
                            continue
                        url = res.get("url") or ""
                        if url:
                            img = await _download_image(sess, url, quality)
                            if img:
                                log.info("image from openverse: %s", query)
                                return img
        except Exception as e:
            log.warning("openverse failed for %s: %s", query, e)
    return None


async def fetch_images_parallel(queries, limit=4, quality=85):
    tasks = []
    for i, q in enumerate(queries):
        if i >= limit or not q:
            continue
        tasks.append(fetch_topic_image(q, quality))
    if not tasks:
        return [None] * len(queries)
    try:
        results = await asyncio.wait_for(
            asyncio.gather(*tasks, return_exceptions=True),
            timeout=12
        )
    except asyncio.TimeoutError:
        log.warning("Image loading timed out after 12s")
        results = [None] * len(tasks)
    out = []
    for r in results:
        if isinstance(r, BaseException):
            out.append(None)
        else:
            out.append(r)
    while len(out) < len(queries):
        out.append(None)
    return out


def make_progress_bar(percent, width=20):
    filled = int(width * percent / 100)
    bar = "█" * filled + "░" * (width - filled)
    return f"[{bar}] {percent}%"


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


def add_text(slide, x, y, w, h, text, size, color, bold=False, italic=False, align=PP_ALIGN.LEFT, font="Calibri"):
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
    add_rect(slide, x - Inches(0.07), y - Inches(0.07), w + Inches(0.14), h + Inches(0.14), pal["accent"])
    slide.shapes.add_picture(io.BytesIO(img), x, y, width=w, height=h)


def clean_bullets(raw, limit=4):
    out = []
    for b in raw or []:
        t = str(b).strip()
        t = t.strip("\"'""«»`)")
        if len(t) < 3 or t in (",", "."):
            continue
        out.append(t[:200])
        if len(out) >= limit:
            break
    return out


def add_bullets(slide, x, y, w, h, bullets, pal, size=18):
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


def ensure_scheme(item):
    scheme = item.get("scheme")
    if scheme and isinstance(scheme, list) and len(scheme) >= 2:
        return scheme
    title = str(item.get("title", "")).strip()
    if not title:
        return None
    default_steps = [
        f"Шаг 1: {title[:30]}...",
        "Шаг 2: Детализация процесса",
        "Шаг 3: Результат и вывод",
    ]
    return default_steps


def add_scheme(slide, steps, pal, top):
    n = min(len(steps), 4)
    box_w = Inches(2.7)
    gap = Inches(0.45)
    total = n * box_w + (n - 1) * gap
    start = int((Inches(13.333) - total) / 2)
    for i, step in enumerate(steps[:4]):
        x = start + i * (box_w + gap)
        box = add_rect(
            slide, x, top, box_w, Inches(1.2),
            pal["panel"], MSO_SHAPE.ROUNDED_RECTANGLE,
        )
        box.line.color.rgb = pal["accent"]
        box.line.width = Pt(1.5)
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.12)
        tf.margin_right = Inches(0.12)
        p = tf.paragraphs[0]
        p.text = f"{i + 1}. {step}"
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = pal["text"]
        p.font.name = pal["bfont"]
        if i < n - 1:
            add_rect(
                slide, x + box_w + Inches(0.05), top + Inches(0.45),
                Inches(0.35), Inches(0.3), pal["accent"],
                MSO_SHAPE.RIGHT_ARROW,
            )


def add_title_slide(prs, content, pal, s, img, date_str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["bg"])
    add_rect(slide, Inches(-1.6), Inches(4.6), Inches(5.2), Inches(5.2), pal["accent2"], MSO_SHAPE.OVAL)
    add_rect(slide, Inches(10.4), Inches(-2.2), Inches(5.0), Inches(5.0), pal["accent"], MSO_SHAPE.OVAL)
    tw = Inches(7.4) if img else Inches(11.3)
    add_text(slide, Inches(1.0), Inches(1.5), tw, Inches(2.5), str(content.get("title") or s.topic), 60, pal["text"], bold=True, font=pal["hfont"])
    add_rect(slide, Inches(1.05), Inches(4.2), Inches(3.5), Inches(0.1), pal["accent"])
    add_text(slide, Inches(1.05), Inches(4.5), tw, Inches(0.8), str(content.get("subtitle") or ""), 24, pal["muted"], font=pal["bfont"])
    intro = str(content.get("intro") or "")
    if intro:
        add_text(slide, Inches(1.05), Inches(5.4), tw, Inches(1.1), intro, 20, pal["accent"], italic=True, font=pal["bfont"])
    if img:
        add_picture(slide, img, Inches(8.8), Inches(2.0), Inches(3.8), Inches(2.9), pal)
    add_text(slide, Inches(0.5), Inches(7.0), Inches(5.0), Inches(0.4), f"📅 {date_str}", 11, pal["muted"], font=pal["bfont"])


def add_agenda_slide(prs, titles, pal, date_str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["bg"])
    add_rect(slide, 0, 0, Inches(0.18), prs.slide_height, pal["accent2"])
    add_rect(slide, 0, 0, prs.slide_width, Inches(0.12), pal["accent"])
    add_text(slide, Inches(0.6), Inches(0.4), Inches(12.1), Inches(0.9), "Что внутри", 40, pal["text"], bold=True, font=pal["hfont"])
    add_rect(slide, Inches(0.62), Inches(1.45), Inches(2.6), Inches(0.08), pal["accent"])
    tb = slide.shapes.add_textbox(Inches(0.65), Inches(1.9), Inches(12.0), Inches(5.1))
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
    add_text(slide, Inches(0.5), Inches(7.0), Inches(5.0), Inches(0.4), f"📅 {date_str}", 11, pal["muted"], font=pal["bfont"])


def add_content_slide(prs, item, pal, num, total, img, date_str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["bg"])
    add_rect(slide, 0, 0, Inches(0.18), prs.slide_height, pal["accent2"])
    add_rect(slide, 0, 0, prs.slide_width, Inches(0.12), pal["accent"])
    add_numbered_title(slide, Inches(0.6), Inches(0.35), Inches(12.1), Inches(1.0), f"{num - 2:02d}", str(item.get("title") or ""), 40, pal)
    add_rect(slide, Inches(0.62), Inches(1.42), Inches(2.6), Inches(0.08), pal["accent"])
    lead = str(item.get("lead") or "")
    if lead:
        add_text(slide, Inches(0.62), Inches(1.62), Inches(12.1), Inches(0.6), lead, 19, pal["muted"], italic=True, font=pal["bfont"])
    bullets = clean_bullets(item.get("bullets"), 4)
    scheme = ensure_scheme(item)
    wide = Inches(12.0)
    if img:
        wide = Inches(8.2)
        add_picture(slide, img, Inches(9.3), Inches(2.25), Inches(3.45), Inches(2.6), pal)
    if scheme:
        add_bullets(slide, Inches(0.65), Inches(2.3), wide, Inches(2.0), bullets[:3], pal, size=17)
        add_scheme(slide, scheme, pal, Inches(4.55))
    else:
        bsize = 18 if (img or len(bullets) >= 3) else 20
        add_bullets(slide, Inches(0.65), Inches(2.3), wide, Inches(3.4), bullets, pal, size=bsize)
    fact = str(item.get("fact") or "")
    if fact:
        add_rect(slide, Inches(0.6), Inches(6.1), Inches(12.1), Inches(1.0), pal["panel"], MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(slide, Inches(0.6), Inches(6.1), Inches(0.1), Inches(1.0), pal["accent"])
        add_text(slide, Inches(0.95), Inches(6.3), Inches(11.5), Inches(0.65), "💡 " + fact, 17, pal["text"], italic=True, font=pal["bfont"])
    add_text(slide, Inches(11.9), Inches(7.12), Inches(1.3), Inches(0.35), f"{num} / {total}", 12, pal["muted"], align=PP_ALIGN.RIGHT, font=pal["bfont"])
    notes = slide.notes_slide.notes_text_frame
    notes.text = str(item.get("speaker_notes") or f"Слайд {num}: {item.get('title', '')}")


def add_stats_slide(prs, stats, pal, num, total, date_str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["bg"])
    add_rect(slide, 0, 0, Inches(0.18), prs.slide_height, pal["accent2"])
    add_rect(slide, 0, 0, prs.slide_width, Inches(0.12), pal["accent"])
    add_text(slide, Inches(0.6), Inches(0.4), Inches(12.1), Inches(0.9), "Цифры и факты", 40, pal["text"], bold=True, font=pal["hfont"])
    add_rect(slide, Inches(0.62), Inches(1.45), Inches(2.6), Inches(0.08), pal["accent"])
    n = min(len(stats), 3)
    card_w = Inches(3.7)
    gap = Inches(0.45)
    total_w = n * card_w + (n - 1) * gap
    start = int((prs.slide_width - total_w) / 2)
    for i, st in enumerate(stats[:3]):
        x = start + i * (card_w + gap)
        add_rect(slide, x, Inches(2.1), card_w, Inches(3.0), pal["panel"], MSO_SHAPE.ROUNDED_RECTANGLE)
        if isinstance(st, dict):
            value = str(st.get("value", ""))
            label = str(st.get("label", ""))
        else:
            value = str(st)
            label = ""
        add_text(slide, x + Inches(0.2), Inches(2.6), card_w - Inches(0.4), Inches(1.2), value, 54, pal["accent"], bold=True, align=PP_ALIGN.CENTER, font=pal["hfont"])
        add_text(slide, x + Inches(0.2), Inches(3.9), card_w - Inches(0.4), Inches(1.0), label, 18, pal["muted"], align=PP_ALIGN.CENTER, font=pal["bfont"])
    add_text(slide, Inches(11.9), Inches(7.12), Inches(1.3), Inches(0.35), f"{num} / {total}", 12, pal["muted"], align=PP_ALIGN.RIGHT, font=pal["bfont"])
    add_text(slide, Inches(0.5), Inches(7.0), Inches(5.0), Inches(0.4), f"📅 {date_str}", 11, pal["muted"], font=pal["bfont"])


def add_chart_slide(prs, chart, rows, pal, num, total, date_str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["bg"])
    add_rect(slide, 0, 0, Inches(0.18), prs.slide_height, pal["accent2"])
    add_rect(slide, 0, 0, prs.slide_width, Inches(0.12), pal["accent"])
    add_text(slide, Inches(0.6), Inches(0.4), Inches(12.1), Inches(0.9), str(chart.get("title") or "Диаграмма"), 40, pal["text"], bold=True, font=pal["hfont"])
    add_rect(slide, Inches(0.62), Inches(1.45), Inches(2.6), Inches(0.08), pal["accent"])
    cd = CategoryChartData()
    categories = [r[0] for r in rows]
    values_list = [r[1] for r in rows]
    cd.categories = categories
    cd.add_series("Значения", values_list)
    ctype = XL_CHART_TYPE.PIE
    if str(chart.get("type", "")).lower() == "bar":
        ctype = XL_CHART_TYPE.BAR_CLUSTERED
    gf = slide.shapes.add_chart(ctype, Inches(1.0), Inches(1.8), Inches(11.3), Inches(5.1), cd)
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
            colors = (pal["accent"], pal["accent2"], pal["muted"], pal["accent"], pal["accent2"], pal["muted"])
            for i, point in enumerate(plot.series[0].points):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = colors[i % len(colors)]
    except Exception as e:
        log.warning("chart styling skipped: %s", e)
    add_text(slide, Inches(11.9), Inches(7.12), Inches(1.3), Inches(0.35), f"{num} / {total}", 12, pal["muted"], align=PP_ALIGN.RIGHT, font=pal["bfont"])
    add_text(slide, Inches(0.5), Inches(7.0), Inches(5.0), Inches(0.4), f"📅 {date_str}", 11, pal["muted"], font=pal["bfont"])


def add_final_slide(prs, content, pal, num, total, date_str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["accent2"])
    add_rect(slide, Inches(-1.4), Inches(-1.8), Inches(4.5), Inches(4.5), pal["accent"], MSO_SHAPE.OVAL)
    add_rect(slide, Inches(10.6), Inches(5.0), Inches(4.5), Inches(4.5), pal["bg"], MSO_SHAPE.OVAL)
    add_text(slide, Inches(1.2), Inches(1.9), Inches(10.9), Inches(1.7), str(content.get("conclusion") or ""), 28, WHITE, align=PP_ALIGN.CENTER, font=pal["bfont"])
    add_text(slide, Inches(1.2), Inches(3.8), Inches(10.9), Inches(1.6), "СПАСИБО ЗА ВНИМАНИЕ!", 56, WHITE, bold=True, align=PP_ALIGN.CENTER, font=pal["hfont"])
    add_text(slide, Inches(11.9), Inches(7.12), Inches(1.3), Inches(0.35), f"{num} / {total}", 12, WHITE, align=PP_ALIGN.RIGHT, font=pal["bfont"])
    add_text(slide, Inches(0.5), Inches(7.0), Inches(5.0), Inches(0.4), f"📅 {date_str}", 11, WHITE, font=pal["bfont"])
    notes = slide.notes_slide.notes_text_frame
    notes.text = "Финальный слайд. Поблагодарите аудиторию и ответьте на вопросы."


def build_pptx(content, s, slide_imgs, title_img, quality=85):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = str(content.get("title") or s.topic)
    prs.core_properties.subject = s.topic
    prs.core_properties.author = "Presentation Bot"
    prs.core_properties.created = datetime.now()
    date_str = datetime.now().strftime("%d.%m.%Y")
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
    msg = STAGE_MESSAGES[min(stage_idx, len(STAGE_MESSAGES)-1)]
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
            log.warning("All AI models failed, falling back to Wikipedia: %s", e)
            content, wiki_thumb = await smart_fallback_from_wikipedia(s)
            source = "wiki"
        
        s.last_content = content
        slides = content.get("slides", [])[: s.count]
        
        # Этап 2: Фото
        await send_stage(status_msg, 1)
        queries = [str(it.get("img") or it.get("title") or s.topic) for it in slides]
        slide_imgs = await fetch_images_parallel(queries, limit=4, quality=85)
        
        title_img = None
        if wiki_thumb:
            try:
                async with aiohttp.ClientSession(headers=HTTP_HEADERS) as sess:
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
        
        # Проверка размера
        file_size_mb = len(data) / (1024 * 1024)
        if file_size_mb > 5:
            log.info(f"Recompressing due to large size ({file_size_mb:.1f}MB)")
            quality = 60
            slide_imgs = await fetch_images_parallel(queries, limit=4, quality=quality)
            title_img = None
            try:
                res = await fetch_images_parallel([str(content.get("img") or s.topic)], limit=1, quality=quality)
                title_img = res[0] if res else None
            except Exception:
                pass
            s.last_imgs = slide_imgs
            s.last_title_img = title_img
            data = build_pptx(content, s, slide_imgs, title_img)
        
        # Этап 4: Финализация
        await send_stage(status_msg, 3)
        name = re.sub(r"[^\w\- ]+", "", s.topic).strip()[:40]
        name = name or "presentation"
        pal_name = PALETTE_NAMES.get(resolve_style(s), "")
        secs = int(time.monotonic() - t0)
        total_words = sum(len(str(b).split()) for sl in slides for b in sl.get("bullets", []))
        
        caption = f"✅ Готово за {secs}с!\n"
        caption += f"🎨 {pal_name} | 🖼 {pics_loaded} фото | 📝 {total_words} слов\n"
        caption += f"v{VER} | {len(slides)} слайдов | {file_size_mb:.1f} MB"
        if source == "wiki":
            caption += "\n⚠️ Использован резервный источник (Википедия)"
        
        await context.bot.send_document(
            chat_id=chat_id,
            document=InputFile(io.BytesIO(data), filename=f"{name}.pptx"),
            caption=caption,
        )
        
        buttons = edit_keyboard(len(slides))
        buttons.inline_keyboard.append([InlineKeyboardButton("🔄 Перегенерировать", callback_data="regen")])
        await context.bot.send_message(
            chat_id,
            "✏️ Нажми номер слайда и напиши новый заголовок:",
            reply_markup=buttons,
        )
        
        await send_stage(status_msg, 4)
        
    except Exception as e:
        log.exception("generation failed")
        reason = f"{type(e).__name__}: {e}"
        if len(reason) > 300:
            reason = reason[:300] + "…"
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
    if not any([GIGACHAT_CLIENT_ID, GIGACHAT_AUTH_KEY]):
        log.error("At least one AI API key must be provided!")
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
    
    # ИСПРАВЛЕНО: убрали ссылку на несуществующий MODEL_CHAIN
    log.info("Using GigaChat API")

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