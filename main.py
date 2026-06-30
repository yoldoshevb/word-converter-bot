import os
import io
import re
import logging
import tempfile
import asyncio
import sqlite3
from datetime import datetime
from io import BytesIO

import fitz  # PyMuPDF — pip install pymupdf
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image, ImageEnhance
import pytesseract
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    filters, ContextTypes,
)

# =============== LOGGING ===============
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# =============== KONFIGURATSIYA ===============
# MUHIM: Tokenni hech qachon kod ichida yozmang! Faqat environment variable orqali bering.
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
if not TOKEN:
    raise RuntimeError(
        "TELEGRAM_BOT_TOKEN environment variable topilmadi. "
        "Avvalgi tokenni BotFather orqali /revoke qiling va yangisini xavfsiz joyga (.env yoki "
        "server secrets) saqlang."
    )

ADMIN_ID = int(os.getenv("ADMIN_ID", "8805822154"))
ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "@yoldoshev_3")
BOT_NAME = "📄 Professional File Converter"
MAX_FILE_SIZE = 50 * 1024 * 1024

# OCR tillari: o'zbek (lotin/krill), rus, ingliz. Tesseract'da mos til paketlari o'rnatilgan bo'lishi kerak:
# sudo apt install tesseract-ocr-uzb tesseract-ocr-rus tesseract-ocr-eng
OCR_LANGS = os.getenv("OCR_LANGS", "uzb+rus+eng")

# =============== SQLITE STATISTIKA ===============
DB_PATH = os.getenv("DB_PATH", "bot_stats.db")


def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            username TEXT,
            first_seen TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS conversions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            file_type TEXT,
            file_name TEXT,
            success INTEGER,
            created_at TEXT
        )
    """)
    conn.commit()
    conn.close()


def log_user(user_id: int, username: str):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        "INSERT OR IGNORE INTO users (user_id, username, first_seen) VALUES (?, ?, ?)",
        (user_id, username, datetime.now().isoformat())
    )
    conn.commit()
    conn.close()


def log_conversion(user_id: int, file_type: str, file_name: str, success: bool):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO conversions (user_id, file_type, file_name, success, created_at) VALUES (?, ?, ?, ?, ?)",
        (user_id, file_type, file_name, int(success), datetime.now().isoformat())
    )
    conn.commit()
    conn.close()


def get_stats() -> dict:
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM users")
    total_users = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM conversions")
    total_conversions = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM conversions WHERE success = 1")
    successful = cur.fetchone()[0]
    cur.execute("""
        SELECT file_type, COUNT(*) c FROM conversions
        GROUP BY file_type ORDER BY c DESC LIMIT 5
    """)
    top_types = cur.fetchall()
    conn.close()
    return {
        "total_users": total_users,
        "total_conversions": total_conversions,
        "successful": successful,
        "top_types": top_types,
    }


# =============== WORD HUJJAT YARATISH ===============
def create_doc(title: str, original_name: str = "") -> Document:
    doc = Document()

    section = doc.sections[0]
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)

    title_heading = doc.add_heading(title, level=0)
    title_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title_heading.runs:
        run.font.size = Pt(18)
        run.font.color.rgb = RGBColor(0, 51, 102)

    date_para = doc.add_paragraph()
    date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_run = date_para.add_run(f"Sana: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    date_run.font.size = Pt(9)
    date_run.font.color.rgb = RGBColor(128, 128, 128)

    if original_name:
        file_para = doc.add_paragraph()
        file_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        file_run = file_para.add_run(f"Asl fayl: {original_name}")
        file_run.font.size = Pt(9)
        file_run.font.color.rgb = RGBColor(128, 128, 128)

    line_para = doc.add_paragraph()
    line_run = line_para.add_run("─" * 70)
    line_run.font.size = Pt(6)
    line_run.font.color.rgb = RGBColor(200, 200, 200)
    doc.add_paragraph()

    return doc


def clean_text(text: str) -> str:
    """Suv belgilari va keraksiz belgilarni tozalash"""
    if not text or len(text.strip()) < 2:
        return ""

    text = text.strip()

    if not any(c.isalnum() for c in text) and len(text) < 15:
        return ""

    if re.match(r'^[\.\-\s=_~|]{5,}$', text):
        return ""

    text = re.sub(r'[ \t]+', ' ', text)
    return text


# =============== PDF → WORD (PyMuPDF bilan: matn + formatlash + rasmlar) ===============
def pdf_to_word(pdf_content: bytes, original_name: str) -> BytesIO:
    """PDF ni Word ga o'tkazish: matn formatlash (bold/o'lcham), jadval, va sahifadagi rasmlarni saqlab"""
    doc = create_doc("📄 PDF Hujjat", original_name)

    pdf = fitz.open(stream=pdf_content, filetype="pdf")
    total_pages = len(pdf)

    for page_num in range(total_pages):
        page = pdf[page_num]

        if total_pages > 1:
            doc.add_heading(f'Sahifa {page_num + 1}', level=1)

        # ---------- 0. Avval jadval hududlarini aniqlab olamiz (matn bilan dublikat bo'lmasligi uchun) ----------
        table_rects = []
        tabs = None
        try:
            tabs = page.find_tables()
            for table_obj in tabs.tables:
                table_rects.append(fitz.Rect(table_obj.bbox))
        except Exception as e:
            logger.warning(f"Jadval qidirishda xatolik (sahifa {page_num + 1}, {original_name}): {e}")

        def inside_table(bbox):
            r = fitz.Rect(bbox)
            return any(r.intersects(tr) for tr in table_rects)

        # ---------- 1. Matn bloklari, rasmlar va jadvallarni bitta ro'yxatga yig'amiz (joylashuv bo'yicha) ----------
        items = []  # (y_top, x_left, type, payload)

        try:
            blocks = page.get_text("dict")["blocks"]
            for block in blocks:
                if block.get("type") != 0:  # 0 = matn bloki
                    continue
                bbox = block.get("bbox", (0, 0, 0, 0))
                if inside_table(bbox):
                    continue  # bu matn allaqachon jadval sifatida chiqariladi
                items.append((bbox[1], bbox[0], "text_block", block))
        except Exception as e:
            logger.warning(f"Matn olishda xatolik (sahifa {page_num + 1}, fayl {original_name}): {e}")

        try:
            for img_index, img in enumerate(page.get_images(full=True), 1):
                xref = img[0]
                try:
                    img_rects = page.get_image_rects(xref)
                    bbox = img_rects[0] if img_rects else fitz.Rect(0, 0, 1, 1)
                    items.append((bbox.y0, bbox.x0, "image", (xref, img_index)))
                except Exception as e:
                    logger.warning(f"Rasm joyini topishda xatolik (sahifa {page_num + 1}, {original_name}): {e}")
        except Exception as e:
            logger.warning(f"Rasmlar ro'yxatini olishda xatolik (sahifa {page_num + 1}, {original_name}): {e}")

        if tabs is not None:
            for table_obj in tabs.tables:
                bbox = table_obj.bbox
                items.append((bbox[1], bbox[0], "table", table_obj))

        # ---------- Ustunlarni aniqlash (multi-column PDF uchun) ----------
        # Matn bloklarining x0 koordinatalarini klasterlab, sahifa nechta ustundan iboratligini taxmin qilamiz
        text_x_positions = sorted(
            it[1] for it in items if it[2] == "text_block"
        )
        column_bounds = []
        if text_x_positions:
            clusters = [[text_x_positions[0]]]
            for x in text_x_positions[1:]:
                if x - clusters[-1][-1] > 40:  # 40pt'dan katta bo'shliq = yangi ustun
                    clusters.append([x])
                else:
                    clusters[-1].append(x)
            # Faqat yetarlicha blok bilan tasdiqlangan klasterlarni ustun deb hisoblaymiz
            clusters = [c for c in clusters if len(c) >= 2]
            if len(clusters) >= 2:
                for c in clusters:
                    column_bounds.append(min(c))
                column_bounds.sort()

        def column_index(x_val):
            if not column_bounds:
                return 0
            best = 0
            best_dist = abs(x_val - column_bounds[0])
            for idx, cb in enumerate(column_bounds):
                d = abs(x_val - cb)
                if d < best_dist:
                    best_dist = d
                    best = idx
            return best

        # Avval ustun bo'yicha, keyin har ustun ichida yuqoridan-pastga tartiblash
        items.sort(key=lambda it: (column_index(it[1]), round(it[0] / 5)))

        # ---------- 2. Tartib bo'yicha hujjatga qo'shish ----------
        for y, x, kind, payload in items:
            if kind == "text_block":
                block = payload
                # Bir blokdagi barcha qatorlarni bitta paragrafga birlashtiramiz (bo'linib ketmasligi uchun)
                para = doc.add_paragraph()
                any_text = False
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        span_text = span["text"]
                        if not span_text.strip():
                            continue
                        cleaned_span = re.sub(r'\s+', ' ', span_text)
                        run = para.add_run(cleaned_span + " ")
                        any_text = True
                        font_size = span.get("size", 11)
                        run.font.size = Pt(max(6, min(36, round(font_size))))
                        font_name = span.get("font", "").lower()
                        flags = span.get("flags", 0)
                        run.font.bold = bool(flags & 2 ** 4) or "bold" in font_name
                        run.font.italic = bool(flags & 2 ** 1) or "italic" in font_name or "oblique" in font_name
                if not any_text:
                    p_elem = para._element
                    p_elem.getparent().remove(p_elem)

            elif kind == "image":
                xref, img_index = payload
                try:
                    base_image = pdf.extract_image(xref)
                    image_bytes = base_image["image"]
                    pil_img = Image.open(BytesIO(image_bytes))
                    if pil_img.width < 40 or pil_img.height < 40:
                        continue
                    doc.add_picture(BytesIO(image_bytes), width=Inches(6))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                except Exception as e:
                    logger.warning(
                        f"Rasm chiqarishda xatolik (sahifa {page_num + 1}, img {img_index}, {original_name}): {e}"
                    )

            elif kind == "table":
                table_obj = payload
                try:
                    table_data = table_obj.extract()
                    if not table_data or len(table_data) < 1:
                        continue
                    has_content = any(
                        cell and str(cell).strip()
                        for row in table_data
                        for cell in row if cell
                    )
                    if not has_content:
                        continue

                    rows = len(table_data)
                    cols = max(len(row) for row in table_data if row)
                    if cols == 0:
                        continue

                    table = doc.add_table(rows=rows, cols=cols)
                    table.style = 'Light Grid Accent 1'

                    for i, row in enumerate(table_data):
                        for j in range(cols):
                            if j < len(row) and row[j]:
                                cell_text = clean_text(str(row[j]))
                                if cell_text:
                                    table.cell(i, j).text = cell_text
                    doc.add_paragraph()
                except Exception as e:
                    logger.warning(f"Jadval qo'shishda xatolik (sahifa {page_num + 1}, {original_name}): {e}")

        if page_num < total_pages - 1:
            doc.add_page_break()

    pdf.close()

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# =============== EXCEL → WORD ===============
def excel_to_word(file_content: bytes, file_ext: str, original_name: str) -> BytesIO:
    import pandas as pd

    doc = create_doc("📊 Excel Ma'lumotlari", original_name)

    try:
        if file_ext == 'csv':
            df = pd.read_csv(BytesIO(file_content), encoding='utf-8')
        else:
            df = pd.read_excel(BytesIO(file_content))

        df = df.dropna(how='all', axis=1)
        df = df.dropna(how='all', axis=0)

        if len(df) > 0:
            info_para = doc.add_paragraph()
            info_run = info_para.add_run(f"Qatorlar: {len(df)} | Ustunlar: {len(df.columns)}")
            info_run.font.size = Pt(10)
            info_run.font.color.rgb = RGBColor(100, 100, 100)

            doc.add_paragraph()

            table = doc.add_table(rows=len(df) + 1, cols=len(df.columns))
            table.style = 'Light Grid Accent 1'

            for j, col in enumerate(df.columns):
                cell = table.cell(0, j)
                cell.text = str(col)
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in paragraph.runs:
                        run.font.bold = True
                        run.font.size = Pt(10)

            for i, (_, row) in enumerate(df.iterrows()):
                for j, value in enumerate(row):
                    cell = table.cell(i + 1, j)
                    if pd.notna(value):
                        cell.text = str(value)
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(10)
        else:
            doc.add_paragraph("Ma'lumot topilmadi.")
    except Exception as e:
        logger.error(f"Excel konvertatsiya xatolik ({original_name}): {e}")
        doc.add_paragraph(f"Xatolik: {str(e)[:100]}")

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# =============== MATN → WORD ===============
def text_to_word(text: str) -> BytesIO:
    doc = create_doc("📝 Matn Hujjati")

    for para in text.split('\n'):
        cleaned = clean_text(para)
        if cleaned:
            doc.add_paragraph(cleaned)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# =============== HTML → WORD ===============
def html_to_word(file_content: bytes, original_name: str) -> BytesIO:
    from bs4 import BeautifulSoup

    doc = create_doc("🌐 HTML Hujjat", original_name)

    html = file_content.decode('utf-8', errors='ignore')
    soup = BeautifulSoup(html, 'html.parser')

    for tag in soup(["script", "style"]):
        tag.decompose()

    # Sarlavhalarni ham saqlab qolish (h1-h3)
    for element in soup.find_all(['h1', 'h2', 'h3', 'p', 'li', 'div']):
        text = element.get_text(separator=' ', strip=True)
        cleaned = clean_text(text)
        if not cleaned:
            continue
        if element.name in ('h1', 'h2', 'h3'):
            level = int(element.name[1])
            doc.add_heading(cleaned, level=level)
        else:
            doc.add_paragraph(cleaned)

    # Rasmlarni ham qo'shishga harakat qilish (faqat tashqi URL bo'lsa o'tkazib yuboriladi, base64 bo'lsa olinadi)
    import base64
    for img_tag in soup.find_all('img'):
        src = img_tag.get('src', '')
        if src.startswith('data:image'):
            try:
                header, encoded = src.split(',', 1)
                img_bytes = base64.b64decode(encoded)
                doc.add_picture(BytesIO(img_bytes), width=Inches(5))
            except Exception as e:
                logger.warning(f"HTML rasm qo'shishda xatolik ({original_name}): {e}")

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# =============== EPUB → WORD ===============
def epub_to_word(file_content: bytes, original_name: str) -> BytesIO:
    from ebooklib import epub
    import ebooklib
    from bs4 import BeautifulSoup

    doc = create_doc("📚 Elektron Kitob", original_name)

    temp_epub = tempfile.NamedTemporaryFile(suffix='.epub', delete=False)
    temp_epub.write(file_content)
    temp_epub.close()

    try:
        book = epub.read_epub(temp_epub.name)
        chapter_num = 0

        for item in book.get_items():
            if item.get_type() == ebooklib.ITEM_DOCUMENT:
                soup = BeautifulSoup(item.get_content(), 'html.parser')
                for tag in soup(["script", "style"]):
                    tag.decompose()

                text = soup.get_text()
                if text.strip():
                    chapter_num += 1
                    doc.add_heading(f'Bo\'lim {chapter_num}', level=1)
                    for para in text.split('\n'):
                        cleaned = clean_text(para)
                        if cleaned:
                            doc.add_paragraph(cleaned)
                    doc.add_page_break()
            elif item.get_type() == ebooklib.ITEM_IMAGE:
                # Kitobdagi rasmlarni alohida ilova qilmaymiz (matn bilan bog'lash murakkab),
                # lekin kerak bo'lsa shu yerga qo'shish mumkin.
                pass
    except Exception as e:
        logger.error(f"EPUB konvertatsiya xatolik ({original_name}): {e}")
        doc.add_paragraph(f"Xatolik: {str(e)[:100]}")
    finally:
        os.unlink(temp_epub.name)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# =============== RASM OCR (yaxshilangan) ===============
def preprocess_image_for_ocr(image: Image.Image) -> Image.Image:
    """OCR aniqligini oshirish uchun rasmni qayta ishlash"""
    if image.mode != 'L':
        image = image.convert('L')

    # Kontrastni oshirish
    image = ImageEnhance.Contrast(image).enhance(2.0)
    # O'tkirlikni oshirish
    image = ImageEnhance.Sharpness(image).enhance(1.5)

    # Agar rasm kichik bo'lsa, kattalashtirish (OCR aniqligi pasaymasligi uchun)
    if image.width < 1000:
        scale = 1000 / image.width
        new_size = (int(image.width * scale), int(image.height * scale))
        image = image.resize(new_size, Image.LANCZOS)

    return image


def image_ocr(image_content: bytes) -> str:
    """Rasmdan matn olish — bir necha PSM rejimini sinab, eng yaxshi natijani tanlash"""
    try:
        image = Image.open(BytesIO(image_content))
        image = preprocess_image_for_ocr(image)

        best_text = ""
        for psm in ('6', '3', '11'):
            try:
                config = f'--psm {psm}'
                text = pytesseract.image_to_string(image, lang=OCR_LANGS, config=config)
                if len(text.strip()) > len(best_text.strip()):
                    best_text = text
            except Exception as e:
                logger.warning(f"OCR psm={psm} xatolik: {e}")

        return best_text.strip()
    except Exception as e:
        logger.error(f"OCR xatolik: {e}")
        return ""


# =============== BOT HANDLERLARI ===============

# =============== TUGMALAR (INLINE KEYBOARD) ===============
def main_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📖 Yordam", callback_data="help"),
         InlineKeyboardButton("📋 Formatlar", callback_data="formats")],
        [InlineKeyboardButton("ℹ️ Bot haqida", callback_data="about")],
    ])


def back_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Orqaga", callback_data="menu")]])


HELP_TEXT = f"""📖 YORDAM

📝 1. MATN — matn yozing, Word faylga o'tkaziladi
📄 2. PDF — PDF fayl yuboring (matn, rasm va jadvallar bilan)
📊 3. EXCEL — .xlsx, .xls, .csv yuboring
🌐 4. HTML — .html fayl yuboring
📚 5. EPUB — elektron kitob yuboring
🖼 6. RASM — matnli rasm yuboring

⚠️ Fayl hajmi 50 MB dan oshmasin
📞 Admin: {ADMIN_USERNAME}"""

FORMATS_TEXT = """📋 FORMATLAR

📄 .pdf — PDF hujjatlar (rasm va jadval bilan)
📊 .xlsx, .xls, .csv — Excel
🌐 .html, .htm — Web sahifalar
📚 .epub — Elektron kitoblar
📝 .txt, .md, .json, .xml — Matn
🖼 .jpg, .png, .bmp — Rasmlar

✅ Barchasi → .docx"""

ABOUT_TEXT = f"""ℹ️ BOT HAQIDA

🏷 {BOT_NAME}

💼 Fayllarni Microsoft Word (.docx) formatiga o'tkazish uchun professional bot.
PDF'dagi rasmlar, ko'p ustunli matn va shrift formatlash saqlanadi.

👨‍💼 {ADMIN_USERNAME}
📞 {ADMIN_USERNAME}"""


async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    log_user(user.id, user.username or user.first_name)

    message = f"""✨ Assalomu alaykum, {user.first_name}!

💼 {BOT_NAME} ga xush kelibsiz!

📌 Quyidagi formatlarni Word (.docx) ga o'tkazadi:

📄 PDF hujjatlar (matn + rasmlar + jadvallar)
📊 Excel jadvallari (.xlsx, .xls, .csv)
🌐 HTML sahifalar
📚 EPUB elektron kitoblar
📝 Matn fayllari (.txt, .md, .json, .xml)
🖼 Rasmlardagi matnlar (OCR)

⚙️ Foydalanish:
1️⃣ Fayl yoki matn yuboring
2️⃣ Konvertatsiya avtomatik
3️⃣ Word faylni yuklab oling

📌 Maksimal hajm: 50 MB | 📌 24/7 ishlaydi | 📌 Bepul ✅

Quyidagi tugmalardan foydalaning 👇"""

    await update.message.reply_text(message, reply_markup=main_menu_keyboard())


async def menu_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Inline tugmalar bosilganda ishlaydigan handler"""
    query = update.callback_query
    await query.answer()

    if query.data == "help":
        await query.edit_message_text(HELP_TEXT, reply_markup=back_keyboard())
    elif query.data == "formats":
        await query.edit_message_text(FORMATS_TEXT, reply_markup=back_keyboard())
    elif query.data == "about":
        await query.edit_message_text(ABOUT_TEXT, reply_markup=back_keyboard())
    elif query.data == "menu":
        user = query.from_user
        message = f"""✨ {user.first_name}, asosiy menyu

💼 {BOT_NAME}

Fayl yoki matn yuboring — avtomatik Word'ga o'tkazaman.
Yoki quyidagi tugmalardan foydalaning 👇"""
        await query.edit_message_text(message, reply_markup=main_menu_keyboard())


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(HELP_TEXT, reply_markup=back_keyboard())


async def formats_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(FORMATS_TEXT, reply_markup=back_keyboard())


async def admin_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("⛔ Faqat admin uchun.")
        return

    await update.message.reply_text(f"👑 ADMIN PANEL\n\n🟢 Holat: Aktiv\n👨‍💼 {ADMIN_USERNAME}")


async def stats_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("⛔ Faqat admin uchun.")
        return

    stats = get_stats()
    top_lines = "\n".join(f"  • {ftype}: {count}" for ftype, count in stats["top_types"]) or "  (hali yo'q)"
    message = f"""📊 STATISTIKA

👥 Foydalanuvchilar: {stats['total_users']}
🔄 Jami konvertatsiyalar: {stats['total_conversions']}
✅ Muvaffaqiyatli: {stats['successful']}
❌ Xatolik: {stats['total_conversions'] - stats['successful']}

🏆 Eng ko'p ishlatilgan formatlar:
{top_lines}"""
    await update.message.reply_text(message)


async def about_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(ABOUT_TEXT, reply_markup=back_keyboard())


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    file_name = document.file_name
    user = update.effective_user
    log_user(user.id, user.username or user.first_name)

    if document.file_size > MAX_FILE_SIZE:
        await update.message.reply_text("❌ Fayl hajmi 50 MB dan oshmasligi kerak.")
        return

    file_ext = file_name.lower().split('.')[-1] if '.' in file_name else ''
    supported = ['pdf', 'xlsx', 'xls', 'csv', 'html', 'htm', 'epub', 'txt', 'md', 'json', 'xml']

    if file_ext not in supported:
        await update.message.reply_text(f"❌ .{file_ext} qo'llab-quvvatlanmaydi.\n📋 /formats")
        return

    msg = await update.message.reply_text(f"⏳ {file_name} yuklab olinmoqda...")

    try:
        file = await document.get_file()
        content = await file.download_as_bytearray()
        output_name = file_name.rsplit('.', 1)[0] + '.docx'

        logger.info(f"Konvertatsiya: {file_name} ({file_ext}) — {len(content)} bytes")
        await msg.edit_text(f"⚙️ {file_name} konvertatsiya qilinmoqda... Bu biroz vaqt olishi mumkin.")

        # Og'ir CPU ishini alohida thread'ga chiqaramiz — bot shu vaqtda boshqa
        # foydalanuvchilarga ham javob bera oladi (bloklanib qolmaydi)
        if file_ext == 'pdf':
            output = await asyncio.to_thread(pdf_to_word, bytes(content), file_name)
        elif file_ext in ['xlsx', 'xls', 'csv']:
            output = await asyncio.to_thread(excel_to_word, bytes(content), file_ext, file_name)
        elif file_ext in ['html', 'htm']:
            output = await asyncio.to_thread(html_to_word, bytes(content), file_name)
        elif file_ext == 'epub':
            output = await asyncio.to_thread(epub_to_word, bytes(content), file_name)
        else:
            text = content.decode('utf-8', errors='ignore')
            output = await asyncio.to_thread(text_to_word, text)

        await msg.delete()
        await update.message.reply_document(
            document=output,
            filename=output_name,
            caption=f"✅ Tayyor!\n📎 {file_name}\n📄 {output_name}"
        )
        log_conversion(user.id, file_ext, file_name, success=True)

    except Exception as e:
        await msg.delete()
        logger.error(f"Xatolik ({file_name}): {e}", exc_info=True)
        log_conversion(user.id, file_ext, file_name, success=False)
        await update.message.reply_text(
            f"❌ Xatolik yuz berdi.\n\n"
            f"📎 Fayl: {file_name}\n"
            f"⚠️ Sabab: {str(e)[:150]}\n"
            f"📞 Admin: {ADMIN_USERNAME}"
        )


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    log_user(user.id, user.username or user.first_name)
    msg = await update.message.reply_text("🖼 Rasm qayta ishlanmoqda (OCR)...")

    try:
        photo = update.message.photo[-1]
        file = await photo.get_file()
        content = await file.download_as_bytearray()

        text = await asyncio.to_thread(image_ocr, bytes(content))

        if text and len(text.strip()) > 10:
            output = await asyncio.to_thread(text_to_word, text)
            await msg.delete()
            await update.message.reply_document(
                document=output,
                filename="rasmdagi_matn.docx",
                caption=f"✅ Tayyor!\n📏 {len(text)} belgi"
            )
            log_conversion(user.id, "photo_ocr", "photo.jpg", success=True)
        else:
            await msg.delete()
            log_conversion(user.id, "photo_ocr", "photo.jpg", success=False)
            await update.message.reply_text(
                "⚠️ Rasmdan matn topilmadi.\n\n"
                "📌 Sabablar:\n"
                "• Rasmda matn yo'q\n"
                "• Rasm sifatsiz\n"
                "• Qo'l yozuvi\n\n"
                "💡 Aniq, bosma matnli rasm yuboring."
            )
    except Exception as e:
        await msg.delete()
        logger.error(f"Rasm xatolik: {e}", exc_info=True)
        log_conversion(user.id, "photo_ocr", "photo.jpg", success=False)
        await update.message.reply_text(f"❌ Xatolik.\n📞 {ADMIN_USERNAME}")


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    user = update.effective_user
    log_user(user.id, user.username or user.first_name)

    if len(text) < 10:
        await update.message.reply_text("⚠️ Kamida 10 ta belgi kerak.")
        return

    msg = await update.message.reply_text("⏳ Matn qayta ishlanmoqda...")

    try:
        output = await asyncio.to_thread(text_to_word, text)
        await msg.delete()
        await update.message.reply_document(
            document=output,
            filename="matn.docx",
            caption=f"✅ Tayyor!\n📏 {len(text)} belgi"
        )
        log_conversion(user.id, "text", "matn.txt", success=True)
    except Exception as e:
        await msg.delete()
        logger.error(f"Matn xatolik: {e}", exc_info=True)
        log_conversion(user.id, "text", "matn.txt", success=False)
        await update.message.reply_text("❌ Xatolik yuz berdi.")


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    logger.error(f"Xatolik: {context.error}", exc_info=context.error)


# =============== MAIN ===============
def main():
    logger.info(f"🤖 {BOT_NAME} ishga tushmoqda...")
    init_db()

    app = Application.builder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("formats", formats_command))
    app.add_handler(CommandHandler("admin", admin_command))
    app.add_handler(CommandHandler("stats", stats_command))
    app.add_handler(CommandHandler("about", about_command))

    app.add_handler(CallbackQueryHandler(menu_callback))

    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    app.add_error_handler(error_handler)

    WEBHOOK_URL = os.getenv("WEBHOOK_URL", "")
    PORT = int(os.getenv("PORT", "8080"))

    if WEBHOOK_URL:
        logger.info(f"🌐 Webhook: {WEBHOOK_URL}")
        app.run_webhook(listen="0.0.0.0", port=PORT, webhook_url=f"{WEBHOOK_URL}/webhook", drop_pending_updates=True)
    else:
        logger.info("📡 Polling rejimi")
        app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
