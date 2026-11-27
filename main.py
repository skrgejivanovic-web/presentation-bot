import asyncio
import io
import json
import logging
import os
import re

import aiohttp
from aiohttp import web
from openai import AsyncOpenAI
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
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

PALETTE_NAMES = {
    "dark": "🌙 тёмная", "light": "☀️ светлая", "color": "🎨 яркая",
    "med": "💊 медицинская", "space": "🚀 космическая",
    "nature": "🌿 природная", "finance": "💰 деловая",
    "tech": "🤖 техно", "edu": "🎓 учебная",
    "sport": "🏆 спортивная", "art": "🎭 творческая",
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
        "картинки, схемы, факты и цитаты.\n\n"
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

    elif data.startswith("style:"):
        s.style = data.split(":", 1)[1]
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
            f"⏳ Рисую «{s.topic}» ({s.count} слайдов, "
            f"{style_label(s)})…\nПодбираю картинки и шрифты, "
            "это займёт до пары минут."
        )
        await generate_and_send(update, context, s)


PROMPT = (
    "You are a top-tier presentation designer and writer. "
    "Return ONLY valid JSON, no markdown, no explanations, "
    'with keys: "title" (string), "subtitle" (string), "intro" (1-2 catchy sentences), '
    '"slides" (array of exactly {count} objects, each with: "title" '
    '(string), "lead" (one engaging sentence), "bullets" (array of 5-6 '
    "informative strings, each 8-18 words), "
    '"fact" (one surprising fact, quote or witty remark about the slide '
    'topic), "scheme" (array of 3-4 short step names, include ONLY if the '
    "slide explains a process, otherwise omit)), "
    '"stats" (array of 3 objects with "value" (short number or percent) '
    'and "label" (short description), striking real-world statistics), '
    '"conclusion" (2-3 inspiring sentences). '
    "Write everything in {lang}. Topic: {topic}"
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
                    last_error = ValueError(
                        f"{model}: no JSON ({len(raw)} chars)"
                    )
                    log.warning(str(last_error))
                    continue
                content = json.loads(json_str)
                if not isinstance(content.get("slides"), list):
                    last_error = ValueError(f"{model}: no slides list")
                    log.warning(str(last_error))
                    continue
                if not content["slides"]:
                    last_error = ValueError(f"{model}: empty slides")
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


async def fetch_topic_image(query):
    api = "https://commons.wikimedia.org/w/api.php"
    params = {
        "action": "query",
        "format": "json",
        "generator": "search",
        "gsrnamespace": "6",
        "gsrsearch": f"filetype:bitmap {query}",
        "gsrlimit": "1",
        "prop": "imageinfo",
        "iiprop": "url|mime",
        "iiurlwidth": "900",
    }
    headers = {"User-Agent": "presentation-bot/1.0 (telegram bot)"}
    timeout = aiohttp.ClientTimeout(total=15)
    async with aiohttp.ClientSession(timeout=timeout, headers=headers) as sess:
        async with sess.get(api, params=params) as r:
            data = await r.json()
        pages = (data.get("query") or {}).get("pages") or {}
        for page in pages.values():
            info = (page.get("imageinfo") or [{}])[0]
            url = info.get("thumburl") or info.get("url")
            mime = info.get("mime") or ""
            if url and mime in ("image/jpeg", "image/png"):
                async with sess.get(url) as ir:
                    if ir.status != 200:
                        return None
                    raw = await ir.read()
                return prepare_image(raw)
    return None


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


def add_picture(slide, img, x, y, w, h, pal):
    add_rect(slide, x - Inches(0.07), y - Inches(0.07),
             w + Inches(0.14), h + Inches(0.14), pal["accent"])
    slide.shapes.add_picture(io.BytesIO(img), x, y, width=w, height=h)


def add_bullets(slide, x, y, w, h, bullets, pal):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(10)
        r1 = p.add_run()
        r1.text = "▸ "
        r1.font.size = Pt(16)
        r1.font.bold = True
        r1.font.color.rgb = pal["accent"]
        r1.font.name = pal["bfont"]
        r2 = p.add_run()
        r2.text = str(b)
        r2.font.size = Pt(16)
        r2.font.color.rgb = pal["text"]
        r2.font.name = pal["bfont"]
    return tb


def add_scheme(slide, steps, pal, top):
    n = min(len(steps), 4)
    box_w = Inches(2.7)
    gap = Inches(0.45)
    total = n * box_w + (n - 1) * gap
    start = int((Inches(13.333) - total) / 2)
    for i, step in enumerate(steps[:4]):
        x = start + i * (box_w + gap)
        box = add_rect(
            slide, x, top, box_w, Inches(1.15),
            pal["panel"], MSO_SHAPE.ROUNDED_RECTANGLE,
        )
        box.line.color.rgb = pal["accent"]
        box.line.width = Pt(1.25)
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.1)
        tf.margin_right = Inches(0.1)
        p = tf.paragraphs[0]
        p.text = f"{i + 1}. {step}"
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = pal["text"]
        p.font.name = pal["bfont"]
        if i < n - 1:
            add_rect(
                slide, x + box_w + Inches(0.05), top + Inches(0.42),
                Inches(0.35), Inches(0.3), pal["accent"],
                MSO_SHAPE.RIGHT_ARROW,
            )


def add_title_slide(prs, content, pal, s, img):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["bg"])
    add_rect(slide, Inches(-1.6), Inches(4.6), Inches(5.2), Inches(5.2),
             pal["accent2"], MSO_SHAPE.OVAL)
    add_rect(slide, Inches(10.4), Inches(-2.2), Inches(5.0), Inches(5.0),
             pal["accent"], MSO_SHAPE.OVAL)
    tw = Inches(7.4) if img else Inches(11.3)
    add_text(slide, Inches(1.0), Inches(2.3), tw, Inches(1.7),
             str(content.get("title") or s.topic), 40, pal["text"],
             bold=True, font=pal["hfont"])
    add_text(slide, Inches(1.05), Inches(4.0), tw, Inches(0.7),
             str(content.get("subtitle") or ""), 18, pal["muted"],
             font=pal["bfont"])
    intro = str(content.get("intro") or "")
    if intro:
        add_text(slide, Inches(1.05), Inches(4.8), tw, Inches(1.0),
                 intro, 15, pal["accent"], italic=True, font=pal["bfont"])
    if img:
        add_picture(slide, img, Inches(8.9), Inches(2.4),
                    Inches(3.7), Inches(2.8), pal)


def add_content_slide(prs, item, pal, num, img):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["bg"])
    add_rect(slide, 0, 0, prs.slide_width, Inches(0.12), pal["accent"])
    add_text(slide, Inches(0.6), Inches(0.45), Inches(12.1), Inches(0.9),
             str(item.get("title") or ""), 28, pal["text"], bold=True,
             font=pal["hfont"])
    lead = str(item.get("lead") or "")
    if lead:
        add_text(slide, Inches(0.62), Inches(1.35), Inches(12.1),
                 Inches(0.5), lead, 15, pal["muted"], italic=True,
                 font=pal["bfont"])
    scheme = item.get("scheme") or []
    bullets = item.get("bullets") or []
    if isinstance(bullets, str):
        bullets = [bullets]
    wide = Inches(12.0)
    if img and not scheme:
        wide = Inches(8.3)
        add_picture(slide, img, Inches(9.35), Inches(1.95),
                    Inches(3.4), Inches(2.55), pal)
    if scheme:
        add_bullets(slide, Inches(0.65), Inches(1.95), wide,
                    Inches(2.2), bullets[:4], pal)
        add_scheme(slide, scheme, pal, Inches(4.35))
    else:
        add_bullets(slide, Inches(0.65), Inches(1.95), wide,
                    Inches(3.9), bullets[:6], pal)
    fact = str(item.get("fact") or "")
    if fact:
        add_rect(slide, Inches(0.6), Inches(6.1), Inches(12.1), Inches(0.95),
                 pal["panel"], MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(slide, Inches(0.6), Inches(6.1), Inches(0.09), Inches(0.95),
                 pal["accent"])
        add_text(slide, Inches(0.9), Inches(6.28), Inches(11.6), Inches(0.6),
                 "💡 " + fact, 13, pal["text"], italic=True,
                 font=pal["bfont"])
    add_text(slide, Inches(12.5), Inches(7.05), Inches(0.7), Inches(0.35),
             str(num), 10, pal["muted"], align=PP_ALIGN.RIGHT,
             font=pal["bfont"])


def add_stats_slide(prs, stats, pal):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["bg"])
    add_rect(slide, 0, 0, prs.slide_width, Inches(0.12), pal["accent"])
    add_text(slide, Inches(0.6), Inches(0.5), Inches(12.1), Inches(0.9),
             "Цифры и факты", 30, pal["text"], bold=True, font=pal["hfont"])
    n = min(len(stats), 3)
    card_w = Inches(3.7)
    gap = Inches(0.45)
    total = n * card_w + (n - 1) * gap
    start = int((prs.slide_width - total) / 2)
    for i, st in enumerate(stats[:3]):
        x = start + i * (card_w + gap)
        add_rect(slide, x, Inches(2.2), card_w, Inches(2.8), pal["panel"],
                 MSO_SHAPE.ROUNDED_RECTANGLE)
        if isinstance(st, dict):
            value = str(st.get("value", ""))
            label = str(st.get("label", ""))
        else:
            value = str(st)
            label = ""
        add_text(slide, x + Inches(0.2), Inches(2.7), card_w - Inches(0.4),
                 Inches(1.0), value, 40, pal["accent"], bold=True,
                 align=PP_ALIGN.CENTER, font=pal["hfont"])
        add_text(slide, x + Inches(0.2), Inches(3.8), card_w - Inches(0.4),
                 Inches(1.0), label, 14, pal["muted"], align=PP_ALIGN.CENTER,
                 font=pal["bfont"])


def add_final_slide(prs, content, pal):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, pal["accent2"])
    add_rect(slide, Inches(-1.4), Inches(-1.8), Inches(4.5), Inches(4.5),
             pal["accent"], MSO_SHAPE.OVAL)
    add_rect(slide, Inches(10.6), Inches(5.0), Inches(4.5), Inches(4.5),
             pal["bg"], MSO_SHAPE.OVAL)
    add_text(slide, Inches(1.2), Inches(2.0), Inches(10.9), Inches(1.6),
             str(content.get("conclusion") or ""), 22, WHITE,
             align=PP_ALIGN.CENTER, font=pal["bfont"])
    add_text(slide, Inches(1.2), Inches(3.8), Inches(10.9), Inches(1.4),
             "СПАСИБО ЗА ВНИМАНИЕ!", 44, WHITE, bold=True,
             align=PP_ALIGN.CENTER, font=pal["hfont"])


def build_pptx(content, s, slide_imgs, title_img):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    key = resolve_style(s)
    pal = dict(PALETTES.get(key, PALETTES["dark"]))
    hfont, bfont = THEME_FONTS.get(key, ("Trebuchet MS", "Calibri"))
    pal["hfont"] = hfont
    pal["bfont"] = bfont
    add_title_slide(prs, content, pal, s, title_img)
    num = 2
    slides = content.get("slides", [])[: s.count]
    for i, item in enumerate(slides):
        img = slide_imgs[i] if i < len(slide_imgs) else None
        add_content_slide(prs, item, pal, num, img)
        num += 1
    stats = content.get("stats") or []
    if isinstance(stats, list) and stats:
        add_stats_slide(prs, stats, pal)
        num += 1
    add_final_slide(prs, content, pal)
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
        slides = content.get("slides", [])[: s.count]
        slide_imgs = []
        fetched = 0
        for item in slides:
            img = None
            if fetched < 5:
                try:
                    img = await fetch_topic_image(
                        str(item.get("title") or s.topic)
                    )
                except Exception:
                    img = None
                if img:
                    fetched += 1
            slide_imgs.append(img)
        title_img = None
        try:
            title_img = await fetch_topic_image(s.topic)
        except Exception:
            title_img = None
        data = build_pptx(content, s, slide_imgs, title_img)
        name = re.sub(r"[^\w\- ]+", "", s.topic).strip()[:40]
        name = name or "presentation"
        pal_name = PALETTE_NAMES.get(resolve_style(s), "")
        await context.bot.send_document(
            chat_id=chat_id,
            document=InputFile(io.BytesIO(data), filename=f"{name}.pptx"),
            caption=f"✅ Готово! Оформление: {pal_name}",
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
    app.add_handler(CommandHandler("cancel", cmd_start))
    app.add_handler(CallbackQueryHandler(on_callback))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))

    await app.initialize()
    await app.start()
    await app.updater.start_polling(drop_pending_updates=True)
    log.info("Bot started (polling)")
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