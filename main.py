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
from telegram.error import Forbidden, BadRequest
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
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
if not TOKEN:
    raise RuntimeError(
        "TELEGRAM_BOT_TOKEN environment variable topilmadi."
    )

ADMIN_ID = int(os.getenv("ADMIN_ID", "8805822154"))
ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "@yoldoshev_3")
BOT_NAME = "📄 Professional File Converter"
MAX_FILE_SIZE = 50 * 1024 * 1024
OCR_LANGS = os.getenv("OCR_LANGS", "uzb+rus+eng")
DB_PATH = os.getenv("DB_PATH", "bot_stats.db")

# =============== SQLITE ===============
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS users (user_id INTEGER PRIMARY KEY, username TEXT, first_seen TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS conversions (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, file_type TEXT, file_name TEXT, success INTEGER, created_at TEXT)")
    conn.commit()
    conn.close()

def log_user(user_id: int, username: str):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("INSERT OR IGNORE INTO users (user_id, username, first_seen) VALUES (?, ?, ?)", (user_id, username, datetime.now().isoformat()))
    conn.commit()
    conn.close()

def log_conversion(user_id: int, file_type: str, file_name: str, success: bool):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("INSERT INTO conversions (user_id, file_type, file_name, success, created_at) VALUES (?, ?, ?, ?, ?)", (user_id, file_type, file_name, int(success), datetime.now().isoformat()))
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
    cur.execute("SELECT file_type, COUNT(*) c FROM conversions GROUP BY file_type ORDER BY c DESC LIMIT 5")
    top_types = cur.fetchall()
    conn.close()
    return {"total_users": total_users, "total_conversions": total_conversions, "successful": successful, "top_types": top_types}

# =============== WORD HUJJAT ===============
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

# =============== PDF → WORD ===============
def pdf_to_word(pdf_content: bytes, original_name: str, progress_cb=None) -> BytesIO:
    doc = create_doc("📄 PDF Hujjat", original_name)
    pdf = fitz.open(stream=pdf_content, filetype="pdf")
    total_pages = len(pdf)
    for page_num in range(total_pages):
        page = pdf[page_num]
        if total_pages > 1:
            marker = doc.add_paragraph()
            marker.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            marker_run = marker.add_run(f"— {page_num + 1} —")
            marker_run.font.size = Pt(7)
            marker_run.font.color.rgb = RGBColor(190, 190, 190)
        table_rects = []
        tabs = None
        try:
            drawings_count = len(page.get_drawings())
            if drawings_count < 800:
                tabs = page.find_tables()
                for table_obj in tabs.tables:
                    table_rects.append(fitz.Rect(table_obj.bbox))
        except Exception as e:
            logger.warning(f"Jadval xatolik: {e}")
        def inside_table(bbox):
            r = fitz.Rect(bbox)
            return any(r.intersects(tr) for tr in table_rects)
        items = []
        try:
            blocks = page.get_text("dict")["blocks"]
            for block in blocks:
                if block.get("type") != 0:
                    continue
                bbox = block.get("bbox", (0, 0, 0, 0))
                if inside_table(bbox):
                    continue
                items.append((bbox[1], bbox[0], "text_block", block))
        except Exception as e:
            logger.warning(f"Matn xatolik: {e}")
        try:
            for img_index, img in enumerate(page.get_images(full=True), 1):
                xref = img[0]
                try:
                    img_rects = page.get_image_rects(xref)
                    bbox = img_rects[0] if img_rects else fitz.Rect(0, 0, 1, 1)
                    items.append((bbox.y0, bbox.x0, "image", (xref, img_index)))
                except:
                    pass
        except:
            pass
        if tabs is not None:
            for table_obj in tabs.tables:
                bbox = table_obj.bbox
                items.append((bbox[1], bbox[0], "table", table_obj))
        text_x_positions = sorted(it[1] for it in items if it[2] == "text_block")
        column_bounds = []
        if text_x_positions:
            clusters = [[text_x_positions[0]]]
            for x in text_x_positions[1:]:
                if x - clusters[-1][-1] > 40:
                    clusters.append([x])
                else:
                    clusters[-1].append(x)
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
        items.sort(key=lambda it: (column_index(it[1]), round(it[0] / 5)))
        for y, x, kind, payload in items:
            if kind == "text_block":
                block = payload
                if looks_like_heading(block):
                    text = "".join(s["text"] for line in block.get("lines", []) for s in line.get("spans", [])).strip()
                    text = re.sub(r'\s+', ' ', text)
                    if text:
                        doc.add_heading(text, level=2)
                    continue
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
                except:
                    pass
            elif kind == "table":
                table_obj = payload
                try:
                    table_data = table_obj.extract()
                    if not table_data or len(table_data) < 1:
                        continue
                    has_content = any(cell and str(cell).strip() for row in table_data for cell in row if cell)
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
                except:
                    pass
        if page_num < total_pages - 1:
            doc.add_page_break()
        if progress_cb:
            try:
                progress_cb(page_num + 1, total_pages)
            except:
                pass
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
    import base64
    for img_tag in soup.find_all('img'):
        src = img_tag.get('src', '')
        if src.startswith('data:image'):
            try:
                header, encoded = src.split(',', 1)
                img_bytes = base64.b64decode(encoded)
                doc.add_picture(BytesIO(img_bytes), width=Inches(5))
            except:
                pass
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
    finally:
        os.unlink(temp_epub.name)
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# =============== RASM(LAR) → PDF ===============
def images_to_pdf(images_content: list) -> BytesIO:
    """Bir nechta rasmni bitta PDF faylga birlashtirish"""
    if not images_content:
        raise ValueError("Rasmlar topilmadi")
    
    pil_images = []
    for content in images_content:
        img = Image.open(BytesIO(content))
        if img.mode == 'RGBA':
            background = Image.new('RGB', img.size, (255, 255, 255))
            background.paste(img, mask=img.split()[3])
            img = background
        elif img.mode != 'RGB':
            img = img.convert('RGB')
        pil_images.append(img)
    
    output = BytesIO()
    if len(pil_images) == 1:
        pil_images[0].save(output, format='PDF')
    else:
        pil_images[0].save(output, format='PDF', save_all=True, append_images=pil_images[1:])
    output.seek(0)
    return output

# =============== RASM OCR ===============
def preprocess_image_for_ocr(image: Image.Image) -> Image.Image:
    if image.mode != 'L':
        image = image.convert('L')
    image = ImageEnhance.Contrast(image).enhance(2.0)
    image = ImageEnhance.Sharpness(image).enhance(1.5)
    if image.width < 1000:
        scale = 1000 / image.width
        new_size = (int(image.width * scale), int(image.height * scale))
        image = image.resize(new_size, Image.LANCZOS)
    return image

def image_ocr(image_content: bytes) -> str:
    try:
        image = Image.open(BytesIO(image_content))
        image = preprocess_image_for_ocr(image)
        best_text = ""
        for psm in ('6', '3', '11'):
            try:
                text = pytesseract.image_to_string(image, lang=OCR_LANGS, config=f'--psm {psm}')
                if len(text.strip()) > len(best_text.strip()):
                    best_text = text
            except:
                pass
        return best_text.strip()
    except:
        return ""

def result_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[InlineKeyboardButton("✏️ Nomini o'zgartirish", callback_data="rename")]])

async def send_result_document(update: Update, context: ContextTypes.DEFAULT_TYPE, output: BytesIO, filename: str, caption: str):
    data = output.getvalue()
    context.user_data['last_output_bytes'] = data
    context.user_data['last_output_name'] = filename
    await update.message.reply_document(document=BytesIO(data), filename=filename, caption=caption, reply_markup=result_keyboard())

# =============== TUGMALAR ===============
def main_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📖 Yordam", callback_data="help"), InlineKeyboardButton("📋 Formatlar", callback_data="formats")],
        [InlineKeyboardButton("ℹ️ Bot haqida", callback_data="about"), InlineKeyboardButton("📩 Murojaat", callback_data="contact_admin")],
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
📩 9. Admin bilan bog'lanish — "📩 Murojaat" tugmasi

⚠️ Fayl hajmi 50 MB dan oshmasin
📞 Admin: {ADMIN_USERNAME}"""

FORMATS_TEXT = """📋 FORMATLAR

📄 .pdf — PDF hujjatlar
📊 .xlsx, .xls, .csv — Excel
🌐 .html, .htm — Web sahifalar
📚 .epub — Elektron kitoblar
📝 .txt, .md, .json, .xml — Matn
🖼 .jpg, .png, .bmp — Rasmlar

✅ Barchasi → .docx"""

ABOUT_TEXT = f"""ℹ️ BOT HAQIDA

🏷 {BOT_NAME}
💼 Fayllarni Microsoft Word (.docx) formatiga o'tkazish uchun professional bot.
👨‍💼 {ADMIN_USERNAME}
📞 {ADMIN_USERNAME}"""

# =============== HANDLERLAR ===============
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    log_user(user.id, user.username or user.first_name)
    message = f"""✨ Assalomu alaykum, {user.first_name}!

💼 {BOT_NAME} ga xush kelibsiz!

📌 Quyidagi formatlarni Word (.docx) ga o'tkazadi:
📄 PDF | 📊 Excel | 🌐 HTML | 📚 EPUB | 📝 Matn | 🖼 Rasmlar (OCR)

⚙️ Fayl yoki matn yuboring — avtomatik Word'ga o'tkazaman.
📌 Maksimal hajm: 50 MB | Bepul ✅

Quyidagi tugmalardan foydalaning 👇"""
    await update.message.reply_text(message, reply_markup=main_menu_keyboard())

async def menu_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
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
        await query.message.reply_text("📩 Xabaringizni yozing — adminga yuboriladi.\nBekor qilish uchun /cancel")
    
    elif query.data == "rename":
        if not context.user_data.get('last_output_bytes'):
            await query.message.reply_text("⚠️ O'zgartiriladigan fayl topilmadi.")
            return
        context.user_data['awaiting_rename'] = True
        await query.message.reply_text("✏️ Yangi fayl nomini yozing (kengaytmasiz):")
    
    elif query.data == "make_pdf":
        images = context.user_data.get('pdf_images', [])
        if not images:
            await query.message.reply_text("⚠️ Hech qanday rasm topilmadi. Avval rasm yuboring.")
            return
        
        await query.edit_message_reply_markup(reply_markup=None)
        progress = await query.message.reply_text(f"⚙️ {len(images)} ta rasmdan PDF yaratilmoqda...")
        
        try:
            output = images_to_pdf(images)
            await progress.delete()
            
            data = output.getvalue()
            context.user_data['last_output_bytes'] = data
            context.user_data['last_output_name'] = "rasmlar.pdf"
            
            await query.message.reply_document(
                document=BytesIO(data),
                filename="rasmlar.pdf",
                caption=f"✅ Tayyor!\n🖼 {len(images)} ta rasmdan PDF yaratildi.",
                reply_markup=result_keyboard()
            )
            log_conversion(query.from_user.id, "images_to_pdf", "rasmlar.pdf", success=True)
        except Exception as e:
            await progress.delete()
            logger.error(f"Rasm→PDF xatolik: {e}")
            await query.message.reply_text(f"❌ Xatolik: {str(e)[:150]}\n📞 {ADMIN_USERNAME}")
        finally:
            context.user_data['collecting_pdf'] = False
            context.user_data['pdf_images'] = []
    
    elif query.data == "cancel_pdf":
        context.user_data['collecting_pdf'] = False
        context.user_data['pdf_images'] = []
        await query.edit_message_reply_markup(reply_markup=None)
        await query.message.reply_text("❌ Bekor qilindi.")

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
🏆 Eng ko'p: 
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
        
        if file_ext == 'pdf':
            loop = asyncio.get_running_loop()
            last_edit = {"time": datetime.min, "text": ""}
            
            async def update_progress_msg(text: str):
                now = datetime.now()
                if text == last_edit["text"] or (now - last_edit["time"]).total_seconds() < 2:
                    return
                last_edit["time"] = now
                last_edit["text"] = text
                try:
                    await msg.edit_text(text)
                except:
                    pass
            
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
                await update.message.reply_text(f"⏱ {file_name} konvertatsiyasi juda ko'p vaqt oldi va to'xtatildi.\n💡 Faylni kichikroq qismlarga bo'lib yuboring.")
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
        await send_result_document(update, context, output, output_name, f"✅ Tayyor!\n📎 {file_name}\n📄 {output_name}")
        log_conversion(user.id, file_ext, file_name, success=True)
        
    except Exception as e:
        await msg.delete()
        logger.error(f"Xatolik ({file_name}): {e}")
        await update.message.reply_text(f"❌ Xatolik yuz berdi.\n📞 {ADMIN_USERNAME}")

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    log_user(user.id, user.username or user.first_name)
    photo = update.message.photo[-1]
    file = await photo.get_file()
    content = await file.download_as_bytearray()
    
    if context.user_data.get('collecting_pdf'):
        context.user_data.setdefault('pdf_images', []).append(bytes(content))
        count = len(context.user_data['pdf_images'])
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton(f"✅ PDF qilish ({count} ta rasm)", callback_data="make_pdf")],
            [InlineKeyboardButton("❌ Bekor qilish", callback_data="cancel_pdf")],
        ])
        await update.message.reply_text(f"📥 {count}-rasm qabul qilindi.\nYana rasm yuboring yoki tugmani bosing 👇", reply_markup=keyboard)
        return
    
    msg = await update.message.reply_text("🖼 Rasm qayta ishlanmoqda (OCR)...")
    try:
        text = await asyncio.to_thread(image_ocr, bytes(content))
        if text and len(text.strip()) > 10:
            output = await asyncio.to_thread(text_to_word, text)
            await msg.delete()
            await send_result_document(update, context, output, "rasmdagi_matn.docx", f"✅ Tayyor!\n📏 {len(text)} belgi")
            log_conversion(user.id, "photo_ocr", "photo.jpg", success=True)
        else:
            await msg.delete()
            await update.message.reply_text("⚠️ Rasmdan matn topilmadi.\n📄 Rasmni PDF qilish uchun /rasm_pdf")
    except Exception as e:
        await msg.delete()
        logger.error(f"Rasm xatolik: {e}")
        await update.message.reply_text(f"❌ Xatolik.\n📞 {ADMIN_USERNAME}")

async def rasm_pdf_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['collecting_pdf'] = True
    context.user_data['pdf_images'] = []
    await update.message.reply_text("📄 Rasm → PDF rejimi yoqildi.\n\nEndi bir nechta rasm yuboring, so'ng \"✅ PDF qilish\" tugmasini bosing.")

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    user = update.effective_user
    log_user(user.id, user.username or user.first_name)
    
    if context.user_data.get('awaiting_contact_message'):
        context.user_data['awaiting_contact_message'] = False
        username_display = f"@{user.username}" if user.username else user.first_name
        try:
            await context.bot.send_message(ADMIN_ID, f"📩 YANGI MUROJAAT\n\n👤 {username_display}\n🆔 ID: {user.id}\n\n💬 {text}\n\n↩️ Javob: /reply {user.id} <matn>")
            await update.message.reply_text("✅ Murojaatingiz adminga yuborildi.", reply_markup=back_keyboard())
        except:
            await update.message.reply_text(f"❌ Yuborib bo'lmadi. Admin: {ADMIN_USERNAME}")
        return
    
    if context.user_data.get('awaiting_rename'):
        context.user_data['awaiting_rename'] = False
        data = context.user_data.get('last_output_bytes')
        old_name = context.user_data.get('last_output_name', 'fayl.docx')
        if not data:
            await update.message.reply_text("⚠️ Fayl topilmadi.")
            return
        ext = old_name.rsplit('.', 1)[-1] if '.' in old_name else 'docx'
        new_name = re.sub(r'[\\/:*?"<>|]', '_', text.strip())
        if not new_name.lower().endswith('.' + ext):
            new_name = f"{new_name}.{ext}"
        context.user_data['last_output_name'] = new_name
        await update.message.reply_document(document=BytesIO(data), filename=new_name, caption=f"✅ Nomi o'zgartirildi:\n📄 {new_name}", reply_markup=result_keyboard())
        return
    
    if len(text) < 10:
        await update.message.reply_text("⚠️ Kamida 10 ta belgi kerak.")
        return
    
    msg = await update.message.reply_text("⏳ Matn qayta ishlanmoqda...")
    try:
        output = await asyncio.to_thread(text_to_word, text)
        await msg.delete()
        await send_result_document(update, context, output, "matn.docx", f"✅ Tayyor!\n📏 {len(text)} belgi")
        log_conversion(user.id, "text", "matn.txt", success=True)
    except Exception as e:
        await msg.delete()
        logger.error(f"Matn xatolik: {e}")
        await update.message.reply_text("❌ Xatolik yuz berdi.")

async def cancel_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['awaiting_contact_message'] = False
    context.user_data['awaiting_rename'] = False
    context.user_data['collecting_pdf'] = False
    await update.message.reply_text("❌ Bekor qilindi.", reply_markup=main_menu_keyboard())

async def reply_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("⛔ Faqat admin uchun.")
        return
    args = context.args
    if len(args) < 2:
        await update.message.reply_text("/reply <user_id> <matn>")
        return
    try:
        target_id = int(args[0])
    except ValueError:
        await update.message.reply_text("⚠️ user_id raqam bo'lishi kerak.")
        return
    reply_text = " ".join(args[1:])
    try:
        await context.bot.send_message(target_id, f"📩 ADMIN JAVOBI:\n\n{reply_text}")
        await update.message.reply_text("✅ Javob yuborildi.")
    except Exception as e:
        await update.message.reply_text(f"❌ Yuborib bo'lmadi: {str(e)[:150]}")

async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    logger.error(f"Xatolik: {context.error}")

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
