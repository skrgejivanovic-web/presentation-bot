# -*- coding: utf-8 -*-
"""
🎯 Presentation Agent Bot — автономный файл.
Стек: aiogram 3 + openai (любой совместимый API) + python-pptx + aiohttp.
ENV: BOT_TOKEN, OPENAI_API_KEY, (опц.) OPENAI_BASE_URL, OPENAI_MODEL, PORT
"""
import asyncio
import json
import logging
import os
import re
from dataclasses import dataclass
from datetime import datetime
from html import escape
from io import BytesIO

from aiohttp import web
from aiogram import Bot, Dispatcher, F
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import Command
from aiogram.types import (BufferedInputFile, CallbackQuery,
                           InlineKeyboardButton, InlineKeyboardMarkup, Message)
from openai import AsyncOpenAI
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("agent")

# ─────────────────────────── CONFIG ───────────────────────────
BOT_TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_BOT_TOKEN", "")
API_KEY = os.getenv("OPENAI_API_KEY", "")
BASE_URL = os.getenv("OPENAI_BASE_URL") or None
MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
PORT = int(os.getenv("PORT", "10000"))

client = AsyncOpenAI(api_key=API_KEY or "missing", base_url=BASE_URL, timeout=180)

TONES = {
    "inspire": ("🔥 Вдохновляющий", "вдохновляющий, энергичный, как выступление на TED"),
    "business": ("💼 Деловой", "деловой, уверенный, конкретный, без пафоса"),
    "friendly": ("😊 Лёгкий", "дружелюбный, тёплый, с мягким юмором"),
    "bold": ("⚡ Дерзкий", "провокационный, острый, ломающий стереотипы"),
}

THEMES = {
    "noir":  {"label": "🌌 Noir",  "bg": "101322", "panel": "1B2138", "text": "F4F6FB",
              "muted": "9AA3B8", "accent": "7C6CFF", "accent2": "22D3EE"},
    "paper": {"label": "📄 Paper", "bg": "FAF6EF", "panel": "EFE7D8", "text": "23262E",
              "muted": "8A8578", "accent": "E4572E", "accent2": "2E86AB"},
    "ocean": {"label": "🌊 Ocean", "bg": "0B2545", "panel": "13315C", "text": "EAF4FB",
              "muted": "9FB9CC", "accent": "57C4E5", "accent2": "F4A259"},
}

FONT, QUOTE_FONT = "Arial", "Georgia"


@dataclass
class Session:
    topic: str
    count: int = 8
    tone: str = "inspire"
    theme: str = "noir"
    busy: bool = False
    cancel: bool = False


sessions: dict[int, Session] = {}


class GenerationCancelled(Exception):
    pass


# ─────────────────────────── LLM ───────────────────────────
def parse_json(text: str) -> dict:
    text = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.M).strip()
    try:
        return json.loads(text)
    except Exception:
        m = re.search(r"\{.*\}", text, re.S)
        if m:
            return json.loads(m.group(0))
        raise


async def llm(system: str, user: str, temperature: float = 0.8) -> dict:
    last = None
    for attempt in range(3):
        for use_json in (True, False):
            try:
                kwargs = dict(
                    model=MODEL, temperature=temperature,
                    messages=[{"role": "system", "content": system},
                              {"role": "user", "content": user}],
                )
                if use_json:
                    kwargs["response_format"] = {"type": "json_object"}
                r = await client.chat.completions.create(**kwargs)
                return parse_json(r.choices[0].message.content or "")
            except Exception as e:
                last = e
                if use_json and "response_format" in str(e):
                    continue  # API не поддерживает json mode — пробуем без него
                break
        await asyncio.sleep(1.5 * (attempt + 1))
    raise RuntimeError(f"LLM недоступна: {last}")


CONCEPT_SYS = """Ты — креативный директор презентационного бюро уровня TED и Apple keynote.
Придумай концепцию презентации. Верни ТОЛЬКО JSON:
{"title":"...","subtitle":"...","hook":"...","audience":"...","insight":"...","narrative":"...","closing_call":"..."}
Правила: title ≤ 8 слов, живой, без канцелярита. hook — одна интригующая фраза (парадокс, вопрос или сцена).
insight — неочевидная правда об аудитории. narrative — дуга истории в 1-2 предложениях (напряжение → инсайт → решение).
closing_call — конкретное действие для аудитории. Язык — русский (или язык темы)."""

OUTLINE_SYS = """Ты — сторителлер-структуратор презентаций. Верни ТОЛЬКО JSON:
{"slides":[{"type":"content|quote|stats|section","title_hint":"...","idea":"..."}]}
Ровно K слайдов (титульный и финальный уже существуют, их НЕ включай).
Дуга: первый content — проблема/интрига; далее развитие; stats — доказательства; quote — эмоциональная пауза;
ближе к концу — решение и конкретные шаги.
quote — ровно 1, если K>=5. stats — 1, если K>=4. section — только если K>=8 (максимум 1).
idea — 1-2 предложения с конкретикой: что именно раскрыть на слайде."""

CONTENT_SYS = """Ты — спичрайтер уровня TED. Превращаешь сухие тезисы в живые слайды.
АБСОЛЮТНО ЗАПРЕЩЕНЫ: «в современном мире», «является», «играет важную роль», «неотъемлемая часть»,
«стремительно развивается», «широкий спектр», «ключевой аспект», вода и вводные конструкции.
Правила:
- title ≤ 7 слов, как заголовок газеты: глагол, образ или контраст. Не «Преимущества X», а «X экономит три часа в день».
- content: 3-4 буллета, каждый ≤ 12 слов, законченная мысль с конкретикой, цифрой, образом или действием.
- Показывай «было → стало», используй метафоры и мини-сцены.
- quote: авторская афористичная фраза ≤ 20 слов. НЕ приписывай реальным людям; author — обобщённая роль или пусто.
- stats: 2-3 пары {"value","label"}; только правдоподобные общеизвестные величины; сомнительное округляй
  и помечай в notes как оценку, которую нужно проверить.
- section: title 2-4 слова + subtitle — одна фраза.
- notes к каждому слайду: 2-4 разговорных предложения для докладчика, можно вопрос залу.
Перед ответом перечитай каждый буллет: звучит как методичка — перепиши.
Верни ТОЛЬКО JSON: {"slides":[{"i":<номер>,"title":"...","bullets":["..."],"quote":"...","author":"...",
"stats":[{"value":"...","label":"..."}],"subtitle":"...","notes":"..."}]} — поля только по типу слайда."""


def check_cancel(s: Session):
    if s.cancel:
        raise GenerationCancelled()


async def generate(sess: Session, status) -> tuple[dict, bytes]:
    await status("🧠 Анализирую тему и аудиторию…")
    concept = await llm(CONCEPT_SYS,
                        f"Тема: {sess.topic}\nВсего слайдов: {sess.count}\nТон: {TONES[sess.tone][1]}", 0.9)
    check_cancel(sess)

    await status("🏗 Выстраиваю драматургию и структуру…")
    k = sess.count - 2
    outline = await llm(OUTLINE_SYS,
                        f"K={k}\nКонцепция: {json.dumps(concept, ensure_ascii=False)}\nТон: {TONES[sess.tone][1]}", 0.7)
    metas = [m for m in outline.get("slides", []) if isinstance(m, dict)][:k]
    while len(metas) < k:
        metas.append({"type": "content", "title_hint": f"Аспект {len(metas)+1}", "idea": sess.topic})
    check_cancel(sess)

    slides = []
    for start in range(0, k, 4):
        batch = [{"i": start + j + 2, **m} for j, m in enumerate(metas[start:start + 4])]
        await status(f"✍️ Пишу слайды {batch[0]['i']}–{batch[-1]['i']} из {sess.count}…")
        res = await llm(CONTENT_SYS,
                        f"Концепция: {json.dumps(concept, ensure_ascii=False)}\n"
                        f"Тон: {TONES[sess.tone][1]}\nСлайды: {json.dumps(batch, ensure_ascii=False)}", 0.85)
        got = [s for s in res.get("slides", []) if isinstance(s, dict)]
        by_i = {s.get("i"): s for s in got}
        for j, meta in enumerate(batch):
            data = by_i.get(meta["i"]) or (got[j] if j < len(got) else {})
            slides.append({**meta, **data})
        check_cancel(sess)

    await status("🎨 Верстаю слайды и дизайн…")
    data = build_pptx(concept, slides, sess.theme)
    return concept, data


# ─────────────────────────── PPTX ───────────────────────────
def C(h): return RGBColor.from_string(h.lstrip("#"))


def fit(text, big, mid, small, t1=45, t2=80):
    n = len(text or "")
    return big if n <= t1 else mid if n <= t2 else small


def new_slide(prs, th):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = C(th["bg"])
    return s


def rect(slide, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE):
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.fill.solid()
    sp.fill.fore_color.rgb = C(color)
    sp.line.fill.background()
    sp.shadow.inherit = False
    return sp


def text(slide, x, y, w, h, txt, size, color, bold=False, italic=False,
         font=FONT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = txt or ""
    f = r.font
    f.size, f.bold, f.italic, f.name = Pt(size), bold, italic, font
    f.color.rgb = C(color)
    return tb


def footer(slide, th, title, idx, total):
    text(slide, 0.65, 7.06, 9.5, 0.3, title[:60], 9, th["muted"])
    text(slide, 11.4, 7.06, 1.28, 0.3, f"{idx:02d} / {total:02d}", 9, th["muted"], align=PP_ALIGN.RIGHT)
    rect(slide, 0, 7.44, 13.333, 0.06, th["panel"])
    rect(slide, 0, 7.44, 13.333 * idx / total, 0.06, th["accent"])


def notes(slide, txt):
    if txt:
        slide.notes_slide.notes_text_frame.text = txt


def header(slide, th, title, idx, total):
    text(slide, 11.4, 0.32, 1.28, 0.3, f"{idx:02d} / {total:02d}", 11, th["muted"], align=PP_ALIGN.RIGHT)
    text(slide, 0.65, 0.55, 11.0, 1.15, title, fit(title, 32, 28, 24, 45, 70), th["text"], bold=True)
    rect(slide, 0.65, 1.78, 0.9, 0.05, th["accent"])


def slide_cover(prs, th, concept):
    s = new_slide(prs, th)
    rect(s, 10.4, -1.4, 4.6, 4.6, th["panel"], MSO_SHAPE.OVAL)
    rect(s, 12.1, 5.7, 1.7, 1.7, th["panel"], MSO_SHAPE.OVAL)
    rect(s, 12.55, 6.15, 0.8, 0.8, th["accent2"], MSO_SHAPE.OVAL)
    hook = concept.get("hook", "")
    if hook:
        text(s, 0.9, 1.1, 11.0, 0.7, hook, 15, th["accent"], italic=True)
    rect(s, 0.9, 1.95, 1.5, 0.055, th["accent"])
    title = concept.get("title", "Презентация")
    text(s, 0.9, 2.2, 11.6, 2.2, title, fit(title, 50, 40, 32, 40, 75), th["text"], bold=True)
    text(s, 0.9, 4.6, 10.5, 1.0, concept.get("subtitle", ""), 19, th["muted"])
    meta = datetime.now().strftime("%d.%m.%Y")
    if concept.get("audience"):
        meta += " · " + concept["audience"][:60]
    text(s, 0.9, 6.85, 11.0, 0.4, meta, 12, th["muted"])
    notes(s, f"Хук: {hook}\nИнсайт об аудитории: {concept.get('insight','')}\nДуга: {concept.get('narrative','')}")


def slide_content(prs, th, d, idx, total, deck_title):
    s = new_slide(prs, th)
    title = d.get("title") or d.get("title_hint") or "Ключевая мысль"
    header(s, th, title, idx, total)
    items = [b for b in (d.get("bullets") or []) if b][:5] or [d.get("idea", "")]
    n, longest = len(items), max((len(i) for i in items), default=0)
    size = 20 if n <= 4 and longest <= 95 else 18 if n <= 5 else 16
    tb = s.shapes.add_textbox(Inches(0.65), Inches(2.05), Inches(12.0), Inches(4.6))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after, p.line_spacing = Pt(14), 1.05
        r1 = p.add_run(); r1.text = "●  "
        r1.font.size, r1.font.name = Pt(size - 6), FONT
        r1.font.color.rgb = C(th["accent"])
        r2 = p.add_run(); r2.text = item
        r2.font.size, r2.font.name = Pt(size), FONT
        r2.font.color.rgb = C(th["text"])
    footer(s, th, deck_title, idx, total)
    notes(s, d.get("notes", ""))


def slide_quote(prs, th, d, idx, total, deck_title):
    s = new_slide(prs, th)
    text(s, 0.75, 0.25, 3, 2.2, "“", 160, th["accent"], bold=True, font=QUOTE_FONT)
    q = d.get("quote") or d.get("idea", "")
    text(s, 1.0, 2.3, 11.3, 3.0, q, fit(q, 34, 30, 26, 90, 150),
         th["text"], italic=True, font=QUOTE_FONT, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if d.get("author"):
        text(s, 1.0, 5.55, 11.3, 0.5, "— " + d["author"], 14, th["muted"], align=PP_ALIGN.CENTER)
    footer(s, th, deck_title, idx, total)
    notes(s, d.get("notes", ""))


def slide_stats(prs, th, d, idx, total, deck_title):
    s = new_slide(prs, th)
    header(s, th, d.get("title") or "Цифры, которые решают", idx, total)
    stats = (d.get("stats") or [])[:3] or [{"value": "—", "label": d.get("idea", "")[:80]}]
    n = len(stats)
    col_w = 12.0 / n
    for i, st in enumerate(stats):
        x = 0.65 + i * col_w
        text(s, x, 2.5, col_w, 1.3, str(st.get("value", "")), 58, th["accent"],
             bold=True, align=PP_ALIGN.CENTER)
        text(s, x + 0.3, 3.95, col_w - 0.6, 1.3, str(st.get("label", "")), 16,
             th["text"], align=PP_ALIGN.CENTER)
        if i > 0:
            rect(s, x - 0.01, 2.6, 0.02, 2.4, th["panel"])
    footer(s, th, deck_title, idx, total)
    notes(s, d.get("notes", ""))


def slide_section(prs, th, d, idx, total, deck_title):
    s = new_slide(prs, th)
    text(s, 0.55, 0.7, 6, 2.4, f"{idx-1:02d}", 120, th["panel"], bold=True)
    rect(s, 0.65, 3.35, 1.2, 0.05, th["accent"])
    text(s, 0.65, 3.55, 11.5, 1.1, d.get("title") or d.get("title_hint", ""), 40, th["text"], bold=True)
    text(s, 0.65, 4.75, 10.5, 0.9, d.get("subtitle", ""), 18, th["muted"])
    footer(s, th, deck_title, idx, total)
    notes(s, d.get("notes", ""))


def slide_final(prs, th, concept, total, deck_title):
    s = new_slide(prs, th)
    rect(s, -1.5, 4.8, 4.2, 4.2, th["panel"], MSO_SHAPE.OVAL)
    rect(s, 11.9, -1.2, 3.2, 3.2, th["panel"], MSO_SHAPE.OVAL)
    text(s, 1.0, 1.5, 11.3, 0.5, "ГЛАВНАЯ МЫСЛЬ", 13, th["accent"], bold=True, align=PP_ALIGN.CENTER)
    call = concept.get("closing_call") or concept.get("title", "")
    text(s, 1.0, 2.1, 11.3, 2.3, call, fit(call, 30, 26, 22, 60, 110),
         th["text"], bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 1.0, 5.1, 11.3, 0.6, "Спасибо! Готов ответить на вопросы.", 18, th["muted"], align=PP_ALIGN.CENTER)
    footer(s, th, deck_title, total, total)
    notes(s, f"Финальный призыв: {call}. Поблагодари аудиторию и открой обсуждение.")


def build_pptx(concept, slides, theme_key) -> bytes:
    th = THEMES.get(theme_key, THEMES["noir"])
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)  # 16:9
    total = len(slides) + 2
    deck_title = concept.get("title", "Презентация")
    slide_cover(prs, th, concept)
    for idx, d in enumerate(slides, start=2):
        t = d.get("type", "content")
        if t == "quote":
            slide_quote(prs, th, d, idx, total, deck_title)
        elif t == "stats":
            slide_stats(prs, th, d, idx, total, deck_title)
        elif t == "section":
            slide_section(prs, th, d, idx, total, deck_title)
        else:
            slide_content(prs, th, d, idx, total, deck_title)
    slide_final(prs, th, concept, total, deck_title)
    buf = BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ─────────────────────────── TELEGRAM ───────────────────────────
async def safe_edit(msg: Message, txt: str, kb=None):
    try:
        await msg.edit_text(txt, reply_markup=kb)
    except Exception:
        pass


def settings_kb(s: Session) -> InlineKeyboardMarkup:
    mark = lambda cond, label: ("✅ " if cond else "") + label
    rows = [
        [InlineKeyboardButton(text=mark(s.count == n, str(n)), callback_data=f"n:{n}")