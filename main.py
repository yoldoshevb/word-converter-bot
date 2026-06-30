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


def looks_like_heading(block, body_font_size: float = 11.0) -> bool:
    """Bir qatorli, qisqa, asosan bosh harfli yoki o'rtacha shriftdan kattaroq matnni
    sarlavha (masalan, lug'at so'zi yoki bo'lim nomi) deb taxmin qiladi."""
    lines = block.get("lines", [])
    if len(lines) != 1:
        return False

    spans = lines[0].get("spans", [])
    text = "".join(s["text"] for s in spans).strip()
    if not text or len(text) > 70:
        return False

    letters = [c for c in text if c.isalpha()]
    if not letters:
        return False

    upper_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
    max_size = max((s.get("size", body_font_size) for s in spans), default=body_font_size)
    is_bold = any(bool(s.get("flags", 0) & 2 ** 4) for s in spans)

    return upper_ratio > 0.8 or max_size >= body_font_size + 2 or (is_bold and len(text.split()) <= 6)


# =============== PDF → WORD (PyMuPDF bilan: matn + formatlash + rasmlar) ===============
def pdf_to_word(pdf_content: bytes, original_name: str, progress_cb=None) -> BytesIO:
    """PDF ni Word ga o'tkazish: matn formatlash (bold/o'lcham), jadval, va sahifadagi rasmlarni saqlab.

    progress_cb(current_page, total_pages) — har sahifadan keyin chaqiriladi (progress-bar uchun).
    """
    doc = create_doc("📄 PDF Hujjat", original_name)

    pdf = fitz.open(stream=pdf_content, filetype="pdf")
    total_pages = len(pdf)

    for page_num in range(total_pages):
        page = pdf[page_num]

        if total_pages > 1:
            # MUHIM: "Sahifa N" endi Heading emas — katta hujjatlarda (yuzlab sahifa) Heading sifatida
            # qo'shilsa Word'ning Table of Contents'i butunlay sahifa raqamlaridan iborat bo'lib qoladi
            # va haqiqiy mazmun (bo'lim sarlavhalari) ko'rinmay qoladi. Endi shunchaki kichik, sezilmas belgi.
            marker = doc.add_paragraph()
            marker.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            marker_run = marker.add_run(f"— {page_num + 1} —")
            marker_run.font.size = Pt(7)
            marker_run.font.color.rgb = RGBColor(190, 190, 190)

        # ---------- 0. Avval jadval hududlarini aniqlab olamiz (matn bilan dublikat bo'lmasligi uchun) ----------
        # MUHIM: find_tables() ba'zi murakkab (ko'p vektor grafikali) sahifalarda juda sekin ishlaydi
        # va butun botni "osilib qolgandek" his qildiradi. Shuning uchun har sahifa uchun vaqt chegarasi
        # qo'yamiz: agar sahifa juda murakkab bo'lsa, jadval qidirishni shunchaki o'tkazib yuboramiz
        # (matn va rasmlar baribir chiqadi — hujjat "chala" emas, faqat jadval formatlanmagan bo'ladi).
        table_rects = []
        tabs = None
        page_start = datetime.now()
        try:
            # Sahifada juda ko'p chizilgan element (vektor grafika) bo'lsa, find_tables() sekinlashadi —
            # bunday hollarda oldindan o'tkazib yuboramiz.
            drawings_count = len(page.get_drawings())
            if drawings_count < 800:
                tabs = page.find_tables()
                for table_obj in tabs.tables:
                    table_rects.append(fitz.Rect(table_obj.bbox))
            else:
                logger.warning(
                    f"Sahifa {page_num + 1} juda murakkab ({drawings_count} chizilgan element) — "
                    f"jadval qidirish o'tkazib yuborildi ({original_name})"
                )
        except Exception as e:
            logger.warning(f"Jadval qidirishda xatolik (sahifa {page_num + 1}, {original_name}): {e}")
        finally:
            elapsed = (datetime.now() - page_start).total_seconds()
            if elapsed > 5:
                logger.warning(f"Sahifa {page_num + 1} jadval qidirish {elapsed:.1f}s vaqt oldi ({original_name})")

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
                if looks_like_heading(block):
                    text = "".join(
                        s["text"] for line in block.get("lines", []) for s in line.get("spans", [])
                    ).strip()
                    text = re.sub(r'\s+', ' ', text)
                    if text:
                        doc.add_heading(text, level=2)
                    continue

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

        if progress_cb:
            try:
                progress_cb(page_num + 1, total_pages)
            except Exception as e:
                logger.warning(f"Progress callback xatolik: {e}")

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


# =============== RASM(LAR) → PDF ===============
def images_to_pdf(images_content: list) -> BytesIO:
    """Bir nechta rasmni bitta PDF faylga birlashtirish"""
    from PIL import Image
    
    if not images_content:
        raise ValueError("Rasmlar topilmadi")
    
    pil_images = []
    for content in images_content:
        img = Image.open(BytesIO(content))
        if img.mode != 'RGB':
            img = img.convert('RGB')
        pil_images.append(img)
    
    output = BytesIO()
    
    if len(pil_images) == 1:
        # Bitta rasm bo'lsa
        pil_images[0].save(output, format='PDF')
    else:
        # Ko'p rasm bo'lsa
        pil_images[0].save(
            output,
            format='PDF',
            save_all=True,
            append_images=pil_images[1:]
        )
    
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


def result_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[InlineKeyboardButton("✏️ Nomini o'zgartirish", callback_data="rename")]])


async def send_result_document(update: Update, context: ContextTypes.DEFAULT_TYPE,
                                output: BytesIO, filename: str, caption: str):
    """Konvertatsiya natijasini yuborish va keyinroq nomini o'zgartirish imkonini saqlash"""
    data = output.getvalue()
    context.user_data['last_output_bytes'] = data
    context.user_data['last_output_name'] = filename

    await update.message.reply_document(
        document=BytesIO(data),
        filename=filename,
        caption=caption,
        reply_markup=result_keyboard()
    )


# =============== BOT HANDLERLARI ===============

# =============== TUGMALAR (INLINE KEYBOARD) ===============
def main_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📖 Yordam", callback_data="help"),
         InlineKeyboardButton("📋 Formatlar", callback_data="formats")],
        [InlineKeyboardButton("ℹ️ Bot haqida", callback_data="about"),
         InlineKeyboardButton("📩 Murojaat", callback_data="contact_admin")],
    ])


def back_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Orqaga", callback_data="menu")]])


HELP_TEXT = f"""📖 YORDAM

📝 1. MATN — matn yozing, Word faylga o'tkaziladi
📄 2. PDF — PDF fayl yuboring (matn, rasm va jadvallar bilan)
📊 3. EXCEL — .xlsx, .xls, .csv yuboring
🌐 4. HTML — .html fayl yuboring
📚 5. EPUB — elektron kitob yuboring
🖼 6. RASM — matnli rasm yuboring (OCR → Word)
📄 7. RASM(LAR) → PDF — /rasm_pdf buyrug'ini yuboring, keyin rasmlarni jo'nating
✏️ 8. Har bir natija ostida "Nomini o'zgartirish" tugmasi bor
📩 9. Admin bilan bog'lanish — "📩 Murojaat" tugmasi yoki istalgan vaqt /cancel bilan bekor qiling

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

    elif query.data == "contact_admin":
        context.user_data['awaiting_contact_message'] = True
        await query.message.reply_text(
            "📩 Murojaat yuborish\n\n"
            "Xabaringizni yozing (taklif, shikoyat, savol va h.k.) — to'g'ridan-to'g'ri adminga yuboriladi.\n"
            "Bekor qilish uchun /cancel yuboring."
        )

    elif query.data == "rename":
        if not context.user_data.get('last_output_bytes'):
            await query.message.reply_text("⚠️ O'zgartiriladigan fayl topilmadi.")
            return
        context.user_data['awaiting_rename'] = True
        await query.message.reply_text(
            "✏️ Yangi fayl nomini yozing (kengaytmasiz, masalan: \"Mening hujjatim\"):"
        )

    elif query.data == "make_pdf":
        images = context.user_data.get('pdf_images', [])
        if not images:
            await query.message.reply_text("⚠️ Hech qanday rasm topilmadi. Avval rasm yuboring.")
            return
        await query.edit_message_reply_markup(reply_markup=None)
        progress = await query.message.reply_text(f"⚙️ {len(images)} ta rasmdan PDF yaratilmoqda...")
        try:
            output = await asyncio.to_thread(images_to_pdf, images)
            await progress.delete()
            await send_result_document(
                update, context, output, "rasmlar.pdf",
                caption=f"✅ Tayyor!\n🖼 {len(images)} ta rasmdan PDF yaratildi."
            )
            log_conversion(query.from_user.id, "images_to_pdf", "rasmlar.pdf", success=True)
        except Exception as e:
            await progress.delete()
            logger.error(f"Rasm→PDF xatolik: {e}", exc_info=True)
            log_conversion(query.from_user.id, "images_to_pdf", "rasmlar.pdf", success=False)
            await query.message.reply_text(f"❌ Xatolik yuz berdi.\n📞 {ADMIN_USERNAME}")
        finally:
            context.user_data['collecting_pdf'] = False
            context.user_data['pdf_images'] = []

    elif query.data == "cancel_pdf":
        context.user_data['collecting_pdf'] = False
        context.user_data['pdf_images'] = []
        await query.edit_message_reply_markup(reply_markup=None)
        await query.message.reply_text("❌ Bekor qilindi. Rasmlar tozalandi.")


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


def make_progress_bar(current: int, total: int, width: int = 12) -> str:
    if total <= 0:
        return ""
    ratio = current / total
    filled = int(ratio * width)
    bar = "■" * filled + "□" * (width - filled)
    percent = int(ratio * 100)
    return f"[{bar}] {current}/{total} sahifa ({percent}%)"


PDF_TIMEOUT_SECONDS = int(os.getenv("PDF_TIMEOUT_SECONDS", "240"))


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

        if file_ext == 'pdf':
            # ---------- Real progress-bar: worker thread'dan asosiy loopga xabar yuborish ----------
            loop = asyncio.get_running_loop()
            last_edit = {"time": datetime.min, "text": ""}

            async def update_progress_msg(text: str):
                # Telegram FloodWait'ga uchramaslik uchun kamida 2 soniyada bir marta yangilaymiz,
                # va bir xil matnni qayta yubormaymiz ("message is not modified" xatosi oldini olish)
                now = datetime.now()
                if text == last_edit["text"] or (now - last_edit["time"]).total_seconds() < 2:
                    return
                last_edit["time"] = now
                last_edit["text"] = text
                try:
                    await msg.edit_text(text)
                except Exception:
                    pass  # mas. "message is not modified" — e'tiborsiz qoldiramiz

            def sync_progress(current: int, total: int):
                bar = make_progress_bar(current, total)
                text = f"⚙️ {file_name} konvertatsiya qilinmoqda...\n{bar}"
                asyncio.run_coroutine_threadsafe(update_progress_msg(text), loop)

            await msg.edit_text(f"⚙️ {file_name} konvertatsiya qilinmoqda...\n{make_progress_bar(0, 1)}")

            try:
                output = await asyncio.wait_for(
                    asyncio.to_thread(pdf_to_word, bytes(content), file_name, sync_progress),
                    timeout=PDF_TIMEOUT_SECONDS
                )
            except asyncio.TimeoutError:
                await msg.delete()
                log_conversion(user.id, file_ext, file_name, success=False)
                keyboard = InlineKeyboardMarkup([
                    [InlineKeyboardButton("📩 Admin bilan bog'lanish", callback_data="contact_admin")]
                ])
                await update.message.reply_text(
                    f"⏱ {file_name} konvertatsiyasi {PDF_TIMEOUT_SECONDS} soniyadan ko'p vaqt oldi va "
                    f"to'xtatildi.\n\n"
                    f"📌 Sabab odatda: PDF juda murakkab (ko'p grafika/skanerlangan sahifalar) yoki juda katta.\n"
                    f"💡 Faylni kichikroq qismlarga bo'lib yuborib ko'ring, yoki admin bilan bog'laning.",
                    reply_markup=keyboard
                )
                return
        elif file_ext in ['xlsx', 'xls', 'csv']:
            await msg.edit_text(f"⚙️ {file_name} konvertatsiya qilinmoqda...")
            output = await asyncio.to_thread(excel_to_word, bytes(content), file_ext, file_name)
        elif file_ext in ['html', 'htm']:
            await msg.edit_text(f"⚙️ {file_name} konvertatsiya qilinmoqda...")
            output = await asyncio.to_thread(html_to_word, bytes(content), file_name)
        elif file_ext == 'epub':
            await msg.edit_text(f"⚙️ {file_name} konvertatsiya qilinmoqda...")
            output = await asyncio.to_thread(epub_to_word, bytes(content), file_name)
        else:
            await msg.edit_text(f"⚙️ {file_name} konvertatsiya qilinmoqda...")
            text = content.decode('utf-8', errors='ignore')
            output = await asyncio.to_thread(text_to_word, text)

        await msg.delete()
        await send_result_document(
            update, context, output, output_name,
            caption=f"✅ Tayyor!\n📎 {file_name}\n📄 {output_name}"
        )
        log_conversion(user.id, file_ext, file_name, success=True)

    except Exception as e:
        await msg.delete()
        logger.error(f"Xatolik ({file_name}): {e}", exc_info=True)
        log_conversion(user.id, file_ext, file_name, success=False)
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("📩 Admin bilan bog'lanish", callback_data="contact_admin")]
        ])
        await update.message.reply_text(
            f"❌ Xatolik yuz berdi.\n\n"
            f"📎 Fayl: {file_name}\n"
            f"⚠️ Sabab: {str(e)[:150]}",
            reply_markup=keyboard
        )


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    log_user(user.id, user.username or user.first_name)

    photo = update.message.photo[-1]
    file = await photo.get_file()
    content = await file.download_as_bytearray()

    # ---------- Rejim 1: Rasm(lar)ni PDF qilish to'plash rejimi (/rasm_pdf bilan boshlangan) ----------
    if context.user_data.get('collecting_pdf'):
        context.user_data.setdefault('pdf_images', []).append(bytes(content))
        count = len(context.user_data['pdf_images'])
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton(f"✅ PDF qilish ({count} ta rasm)", callback_data="make_pdf")],
            [InlineKeyboardButton("❌ Bekor qilish", callback_data="cancel_pdf")],
        ])
        await update.message.reply_text(
            f"📥 {count}-rasm qabul qilindi.\nYana rasm yuboring yoki tugmani bosing 👇",
            reply_markup=keyboard
        )
        return

    # ---------- Rejim 2: Oddiy OCR → Word (standart) ----------
    msg = await update.message.reply_text("🖼 Rasm qayta ishlanmoqda (OCR)...")

    try:
        text = await asyncio.to_thread(image_ocr, bytes(content))

        if text and len(text.strip()) > 10:
            output = await asyncio.to_thread(text_to_word, text)
            await msg.delete()
            await send_result_document(
                update, context, output, "rasmdagi_matn.docx",
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
                "💡 Aniq, bosma matnli rasm yuboring.\n"
                "📄 Yoki rasmni PDF qilish uchun /rasm_pdf buyrug'ini ishlating."
            )
    except Exception as e:
        await msg.delete()
        logger.error(f"Rasm xatolik: {e}", exc_info=True)
        log_conversion(user.id, "photo_ocr", "photo.jpg", success=False)
        await update.message.reply_text(f"❌ Xatolik.\n📞 {ADMIN_USERNAME}")


async def rasm_pdf_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Bir nechta rasmni PDF'ga birlashtirish rejimini boshlash"""
    context.user_data['collecting_pdf'] = True
    context.user_data['pdf_images'] = []
    await update.message.reply_text(
        "📄 Rasm → PDF rejimi yoqildi.\n\n"
        "Endi bir nechta rasm yuboring (tartib bo'yicha), so'ng \"✅ PDF qilish\" tugmasini bosing."
    )


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    user = update.effective_user
    log_user(user.id, user.username or user.first_name)

    # ---------- Agar foydalanuvchidan murojaat (admin'ga xabar) kutilayotgan bo'lsa ----------
    if context.user_data.get('awaiting_contact_message'):
        context.user_data['awaiting_contact_message'] = False
        username_display = f"@{user.username}" if user.username else user.first_name
        try:
            await context.bot.send_message(
                ADMIN_ID,
                f"📩 YANGI MUROJAAT\n\n"
                f"👤 Foydalanuvchi: {username_display}\n"
                f"🆔 ID: {user.id}\n\n"
                f"💬 Xabar:\n{text}\n\n"
                f"↩️ Javob berish uchun: /reply {user.id} <matn>"
            )
            await update.message.reply_text(
                "✅ Murojaatingiz adminga yuborildi. Tez orada javob olasiz.",
                reply_markup=back_keyboard()
            )
        except Exception as e:
            logger.error(f"Murojaatni yuborishda xatolik: {e}", exc_info=True)
            await update.message.reply_text("❌ Murojaatni yuborib bo'lmadi. Birozdan keyin qayta urinib ko'ring.")
        return

    # ---------- Agar foydalanuvchidan yangi fayl nomi kutilayotgan bo'lsa ----------
    if context.user_data.get('awaiting_rename'):
        context.user_data['awaiting_rename'] = False
        data = context.user_data.get('last_output_bytes')
        old_name = context.user_data.get('last_output_name', 'fayl.docx')
        if not data:
            await update.message.reply_text("⚠️ O'zgartiriladigan fayl topilmadi. Avval faylni konvertatsiya qiling.")
            return

        ext = old_name.rsplit('.', 1)[-1] if '.' in old_name else 'docx'
        new_name = re.sub(r'[\\/:*?"<>|]', '_', text.strip())
        if not new_name.lower().endswith('.' + ext):
            new_name = f"{new_name}.{ext}"

        context.user_data['last_output_name'] = new_name
        await update.message.reply_document(
            document=BytesIO(data),
            filename=new_name,
            caption=f"✅ Nomi o'zgartirildi:\n📄 {new_name}",
            reply_markup=result_keyboard()
        )
        return

    if len(text) < 10:
        await update.message.reply_text("⚠️ Kamida 10 ta belgi kerak.")
        return

    msg = await update.message.reply_text("⏳ Matn qayta ishlanmoqda...")

    try:
        output = await asyncio.to_thread(text_to_word, text)
        await msg.delete()
        await send_result_document(
            update, context, output, "matn.docx",
            caption=f"✅ Tayyor!\n📏 {len(text)} belgi"
        )
        log_conversion(user.id, "text", "matn.txt", success=True)
    except Exception as e:
        await msg.delete()
        logger.error(f"Matn xatolik: {e}", exc_info=True)
        log_conversion(user.id, "text", "matn.txt", success=False)
        await update.message.reply_text("❌ Xatolik yuz berdi.")


async def cancel_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['awaiting_contact_message'] = False
    context.user_data['awaiting_rename'] = False
    context.user_data['collecting_pdf'] = False
    await update.message.reply_text("❌ Bekor qilindi.", reply_markup=main_menu_keyboard())


async def reply_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Admin foydalanuvchi murojaatiga javob berish: /reply <user_id> <matn>"""
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("⛔ Faqat admin uchun.")
        return

    args = context.args
    if len(args) < 2:
        await update.message.reply_text("Foydalanish: /reply <user_id> <javob matni>")
        return

    try:
        target_id = int(args[0])
    except ValueError:
        await update.message.reply_text("⚠️ user_id raqam bo'lishi kerak.")
        return

    reply_text = " ".join(args[1:])
    try:
        await context.bot.send_message(
            target_id,
            f"📩 ADMIN JAVOBI:\n\n{reply_text}"
        )
        await update.message.reply_text("✅ Javob yuborildi.")
    except Exception as e:
        logger.error(f"Javob yuborishda xatolik: {e}", exc_info=True)
        await update.message.reply_text(f"❌ Yuborib bo'lmadi: {str(e)[:150]}")


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
    app.add_handler(CommandHandler("rasm_pdf", rasm_pdf_command))
    app.add_handler(CommandHandler("cancel", cancel_command))
    app.add_handler(CommandHandler("reply", reply_command))

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
