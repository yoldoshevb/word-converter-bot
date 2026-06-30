import os
import logging
import tempfile
from datetime import datetime
from io import BytesIO
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import pdfplumber
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

# =============== LOGGING ===============
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# =============== KONFIGURATSIYA ===============
TOKEN = "8805822154:AAHdDCps2cumRbpQWm3iw5bEzg0vTjmMOdQ"
ADMIN_ID = 8805822154
ADMIN_USERNAME = "@yoldoshev_3"
BOT_NAME = "📄 Professional File Converter"
MAX_FILE_SIZE = 50 * 1024 * 1024

# =============== SIFATLI WORD HUJJAT ===============
def create_quality_document(title: str, original_name: str = "") -> Document:
    """Yuqori sifatli Word hujjat yaratish"""
    doc = Document()
    
    # Sahifa sozlamalari
    section = doc.sections[0]
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    
    # Default style
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    style.paragraph_format.space_after = Pt(6)
    style.paragraph_format.line_spacing = 1.15
    
    # Sarlavha
    title_heading = doc.add_heading(title, level=0)
    title_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title_heading.runs:
        run.font.size = Pt(18)
        run.font.color.rgb = RGBColor(0, 51, 102)
    
    # Sana
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
    
    # Ajratuvchi chiziq
    line_para = doc.add_paragraph()
    line_para.paragraph_format.space_before = Pt(12)
    line_para.paragraph_format.space_after = Pt(12)
    line_run = line_para.add_run("─" * 70)
    line_run.font.size = Pt(6)
    line_run.font.color.rgb = RGBColor(200, 200, 200)
    
    return doc

def add_formatted_text(doc: Document, text: str, style_name: str = 'Normal'):
    """Formatlangan matn qo'shish"""
    paragraphs = text.split('\n')
    for para_text in paragraphs:
        if para_text.strip():
            # Suv belgilarini filtrash
            cleaned = clean_text(para_text.strip())
            if cleaned:
                p = doc.add_paragraph(cleaned, style=style_name)

def clean_text(text: str) -> str:
    """Matnni tozalash — suv belgilari va keraksiz belgilarni olib tashlash"""
    # Juda qisqa satrlarni o'tkazib yuborish (suv belgilari odatda qisqa)
    if len(text) < 3 and not any(c.isalnum() for c in text):
        return ""
    
    # Keraksiz belgilarni tozalash
    import re
    # Ko'p takrorlanadigan nuqta yoki chiziqchalar
    if re.match(r'^[\.\-\s\=\_\~]{5,}$', text):
        return ""
    
    # Faqat maxsus belgilar
    if not any(c.isalnum() for c in text) and len(text) < 10:
        return ""
    
    return text

# =============== PDF → WORD (SIFATLI) ===============
def pdf_to_word_quality(pdf_content: bytes, original_name: str) -> BytesIO:
    """Yuqori sifatli PDF to Word konvertatsiyasi"""
    doc = create_quality_document("📄 PDF Hujjat", original_name)
    
    temp_pdf = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
    temp_pdf.write(pdf_content)
    temp_pdf.close()
    
    try:
        with pdfplumber.open(temp_pdf.name) as pdf:
            total_pages = len(pdf.pages)
            
            for page_num, page in enumerate(pdf.pages, 1):
                if total_pages > 1:
                    doc.add_heading(f'Sahifa {page_num}', level=1)
                
                # Matn olish (layout saqlangan holda)
                text = page.extract_text(
                    x_tolerance=2,
                    y_tolerance=2,
                    keep_blank_chars=False,
                    use_text_flow=True
                )
                
                if text:
                    add_formatted_text(doc, text)
                
                # Jadvallar
                tables = page.extract_tables()
                if tables:
                    for table_data in tables:
                        if table_data and len(table_data) > 1:
                            # Bo'sh jadvallarni o'tkazib yuborish
                            has_content = any(
                                cell and str(cell).strip() 
                                for row in table_data 
                                for cell in row
                            )
                            
                            if has_content:
                                rows = len(table_data)
                                cols = max(len(row) for row in table_data if row)
                                
                                if cols > 0:
                                    doc.add_paragraph()  # Bo'sh joy
                                    table = doc.add_table(rows=rows, cols=cols)
                                    table.style = 'Light Grid Accent 1'
                                    
                                    for i, row in enumerate(table_data):
                                        for j, cell in enumerate(row):
                                            if j < cols and cell:
                                                cell_text = clean_text(str(cell))
                                                if cell_text:
                                                    table.cell(i, j).text = cell_text
                                    
                                    doc.add_paragraph()
                
                if page_num < total_pages:
                    doc.add_page_break()
    finally:
        os.unlink(temp_pdf.name)
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# =============== EXCEL → WORD (SIFATLI) ===============
def excel_to_word_quality(file_content: bytes, file_ext: str, original_name: str) -> BytesIO:
    """Yuqori sifatli Excel to Word"""
    import pandas as pd
    
    doc = create_quality_document("📊 Excel Ma'lumotlari", original_name)
    
    if file_ext == 'csv':
        df = pd.read_csv(BytesIO(file_content), encoding='utf-8')
    else:
        df = pd.read_excel(BytesIO(file_content))
    
    # Bo'sh ustunlarni olib tashlash
    df = df.dropna(how='all', axis=1)
    df = df.dropna(how='all', axis=0)
    
    if len(df) == 0:
        doc.add_paragraph("Ma'lumot topilmadi.")
    else:
        info_para = doc.add_paragraph()
        info_run = info_para.add_run(f"Qatorlar: {len(df)} | Ustunlar: {len(df.columns)}")
        info_run.font.size = Pt(10)
        info_run.font.color.rgb = RGBColor(100, 100, 100)
        
        doc.add_paragraph()
        
        table = doc.add_table(rows=len(df) + 1, cols=len(df.columns))
        table.style = 'Light Grid Accent 1'
        
        # Header
        for j, col in enumerate(df.columns):
            cell = table.cell(0, j)
            cell.text = str(col)
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.bold = True
                    run.font.size = Pt(10)
                    run.font.color.rgb = RGBColor(255, 255, 255)
            # Header background
            shading = cell._element.get_or_add_tcPr()
            shading_elm = shading.makeelement(qn('w:shd'), {
                qn('w:fill'): '003366',
                qn('w:val'): 'clear'
            })
            shading.append(shading_elm)
        
        # Ma'lumotlar
        for i, (_, row) in enumerate(df.iterrows()):
            for j, value in enumerate(row):
                cell = table.cell(i + 1, j)
                if pd.notna(value):
                    cell.text = str(value)
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(10)
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# =============== MATN → WORD (SIFATLI) ===============
def text_to_word_quality(text: str) -> BytesIO:
    """Yuqori sifatli matn to Word"""
    doc = create_quality_document("📝 Matn Hujjati")
    
    info_para = doc.add_paragraph()
    info_run = info_para.add_run(f"Belgilar soni: {len(text)}")
    info_run.font.size = Pt(9)
    info_run.font.color.rgb = RGBColor(128, 128, 128)
    
    doc.add_paragraph()
    
    add_formatted_text(doc, text)
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# =============== RASM OCR (SIFATLI) ===============
def image_to_text_quality(image_content: bytes) -> str:
    """Yuqori sifatli OCR"""
    try:
        from PIL import Image, ImageEnhance, ImageFilter
        import pytesseract
        
        image = Image.open(BytesIO(image_content))
        
        # Katta rasmni o'lchamini sozlash
        if image.width > 2000 or image.height > 2000:
            image.thumbnail((2000, 2000), Image.LANCZOS)
        
        # Kontrastni oshirish
        enhancer = ImageEnhance.Contrast(image)
        image = enhancer.enhance(2.0)
        
        # Keskinlikni oshirish
        enhancer = ImageEnhance.Sharpness(image)
        image = enhancer.enhance(2.0)
        
        # Oq-qora qilish
        image = image.convert('L')
        
        # Noise olib tashlash
        image = image.filter(ImageFilter.MedianFilter(size=3))
        
        # OCR
        text = pytesseract.image_to_string(
            image,
            lang='eng+rus+uzb',
            config='--psm 6 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789.,!?\'\"-:;()[]{}/\n '
        )
        
        return text.strip()
    except Exception as e:
        logger.error(f"OCR xatolik: {e}")
        return ""

# =============== HTML → WORD ===============
def html_to_word_quality(file_content: bytes, original_name: str) -> BytesIO:
    from bs4 import BeautifulSoup
    
    doc = create_quality_document("🌐 HTML Hujjat", original_name)
    
    html = file_content.decode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')
    
    # Script va style larni olib tashlash
    for script in soup(["script", "style"]):
        script.decompose()
    
    text = soup.get_text()
    add_formatted_text(doc, text)
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# =============== EPUB → WORD ===============
def epub_to_word_quality(file_content: bytes, original_name: str) -> BytesIO:
    from ebooklib import epub
    from bs4 import BeautifulSoup
    
    doc = create_quality_document("📚 Elektron Kitob", original_name)
    
    temp_epub = tempfile.NamedTemporaryFile(suffix='.epub', delete=False)
    temp_epub.write(file_content)
    temp_epub.close()
    
    try:
        book = epub.read_epub(temp_epub.name)
        chapter_num = 0
        
        for item in book.get_items():
            if item.get_type() == 9:
                soup = BeautifulSoup(item.get_content(), 'html.parser')
                for script in soup(["script", "style"]):
                    script.decompose()
                
                text = soup.get_text()
                if text.strip():
                    chapter_num += 1
                    doc.add_heading(f'Bo\'lim {chapter_num}', level=1)
                    add_formatted_text(doc, text)
                    doc.add_page_break()
    finally:
        os.unlink(temp_epub.name)
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# =============== BOT HANDLERLARI ===============

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    
    message = f"""✨ Assalomu alaykum, {user.first_name}!

💼 {BOT_NAME} ga xush kelibsiz!

📌 Quyidagi formatlarni Word (.docx) ga o'tkazadi:

📄 PDF hujjatlar
📊 Excel jadvallari (.xlsx, .xls, .csv)
🌐 HTML sahifalar
📚 EPUB elektron kitoblar
📝 Matn fayllari (.txt, .md, .json, .xml)
🖼 Rasmlardagi matnlar (OCR)

⚙️ Foydalanish:
1️⃣ Fayl yoki matn yuboring
2️⃣ Konvertatsiya avtomatik
3️⃣ Word faylni yuklab oling

📌 Cheklovlar:
🔹 Maksimal hajm: 50 MB
🔹 24/7 ishlaydi
🔹 Bepul ✅

📋 Buyruqlar:
/start — Bosh menyu
/help — Yordam
/formats — Formatlar
/about — Bot haqida

👨‍💼 Admin: {ADMIN_USERNAME}"""

    await update.message.reply_text(message)

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = f"""📖 YORDAM

📝 1. MATN
Oddiy matn yozing — Word faylga o'tkaziladi.

📄 2. PDF
PDF fayl yuboring. Matn va jadvallar saqlanadi.

📊 3. EXCEL
.xlsx, .xls, .csv fayllarni yuboring.

🌐 4. HTML
.html fayllarni yuboring.

📚 5. EPUB
Elektron kitoblarni yuboring.

🖼 6. RASM
Matnli rasm yuboring. OCR orqali matn ajratiladi.

⚠️ Fayl hajmi 50 MB dan oshmasin.

📞 Admin: {ADMIN_USERNAME}"""

    await update.message.reply_text(message)

async def formats_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = f"""📋 FORMATLAR

📄 Hujjatlar:
• .pdf
• .xlsx, .xls, .csv
• .html, .htm

📚 Kitoblar:
• .epub

📝 Matn:
• Oddiy matn
• .txt, .md, .json, .xml

🖼 Rasmlar (OCR):
• .jpg, .png, .bmp

✅ Barchasi → .docx"""

    await update.message.reply_text(message)

async def admin_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("⛔ Faqat admin uchun.")
        return
    
    message = f"""👑 ADMIN PANEL

🟢 Holat: Aktiv
👨‍💼 Admin: {ADMIN_USERNAME}"""

    await update.message.reply_text(message)

async def about_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = f"""ℹ️ BOT HAQIDA

🏷 {BOT_NAME}

💼 Fayllarni Microsoft Word (.docx) formatiga professional o'tkazish uchun yaratilgan.

👨‍💼 {ADMIN_USERNAME}
📞 {ADMIN_USERNAME}"""

    await update.message.reply_text(message)

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    file_name = document.file_name
    
    if document.file_size > MAX_FILE_SIZE:
        await update.message.reply_text("❌ Fayl hajmi 50 MB dan oshmasligi kerak.")
        return
    
    file_ext = file_name.lower().split('.')[-1] if '.' in file_name else ''
    supported = ['pdf', 'xlsx', 'xls', 'csv', 'html', 'htm', 'epub', 'txt', 'md', 'json', 'xml']
    
    if file_ext not in supported:
        await update.message.reply_text(f"❌ .{file_ext} qo'llab-quvvatlanmaydi.\n📋 /formats")
        return
    
    msg = await update.message.reply_text(f"⏳ {file_name} qayta ishlanmoqda...")
    
    try:
        file = await document.get_file()
        content = await file.download_as_bytearray()
        output_name = file_name.rsplit('.', 1)[0] + '.docx'
        
        if file_ext == 'pdf':
            output = pdf_to_word_quality(content, file_name)
        elif file_ext in ['xlsx', 'xls', 'csv']:
            output = excel_to_word_quality(content, file_ext, file_name)
        elif file_ext in ['html', 'htm']:
            output = html_to_word_quality(content, file_name)
        elif file_ext == 'epub':
            output = epub_to_word_quality(content, file_name)
        else:
            text = content.decode('utf-8')
            output = text_to_word_quality(text)
        
        await msg.delete()
        await update.message.reply_document(
            document=output,
            filename=output_name,
            caption=f"✅ Tayyor!\n📎 {file_name}\n📄 {output_name}"
        )
        
    except Exception as e:
        await msg.delete()
        logger.error(f"Xatolik: {e}")
        await update.message.reply_text(f"❌ Xatolik yuz berdi.\n📞 {ADMIN_USERNAME}")

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = await update.message.reply_text("🖼 Rasm qayta ishlanmoqda...")
    
    try:
        photo = update.message.photo[-1]
        file = await photo.get_file()
        content = await file.download_as_bytearray()
        
        text = image_to_text_quality(content)
        
        if text and len(text.strip()) > 10:
            output = text_to_word_quality(text)
            await msg.delete()
            await update.message.reply_document(
                document=output,
                filename="rasmdagi_matn.docx",
                caption=f"✅ Tayyor!\n📏 {len(text)} belgi"
            )
        else:
            await msg.delete()
            await update.message.reply_text(
                "⚠️ Rasmdan matn topilmadi.\n\n"
                "📌 Sabablar:\n"
                "• Rasmda matn yo'q\n"
                "• Sifatsiz rasm\n"
                "• Qo'l yozuvi\n\n"
                "💡 Aniqroq rasm yuboring."
            )
    except Exception as e:
        await msg.delete()
        logger.error(f"Xatolik: {e}")
        await update.message.reply_text(f"❌ Xatolik.\n📞 {ADMIN_USERNAME}")

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    
    if len(text) < 10:
        await update.message.reply_text("⚠️ Kamida 10 ta belgi kerak.\n📖 /help")
        return
    
    msg = await update.message.reply_text("⏳ Matn qayta ishlanmoqda...")
    
    try:
        output = text_to_word_quality(text)
        await msg.delete()
        await update.message.reply_document(
            document=output,
            filename="matn.docx",
            caption=f"✅ Tayyor!\n📏 {len(text)} belgi"
        )
    except Exception as e:
        await msg.delete()
        logger.error(f"Xatolik: {e}")
        await update.message.reply_text("❌ Xatolik yuz berdi.")

async def error_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    logger.error(f"Xatolik: {context.error}")
    try:
        if update and update.effective_message:
            await update.effective_message.reply_text(f"❌ Xatolik.\n📞 {ADMIN_USERNAME}")
    except:
        pass

# =============== MAIN ===============
def main():
    if not TOKEN:
        logger.error("❌ TOKEN yo'q!")
        return
    
    logger.info(f"🤖 {BOT_NAME} ishga tushmoqda...")
    
    app = Application.builder().token(TOKEN).build()
    
    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("formats", formats_command))
    app.add_handler(CommandHandler("admin", admin_command))
    app.add_handler(CommandHandler("about", about_command))
    
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    
    app.add_error_handler(error_handler)
    
    WEBHOOK_URL = os.getenv("WEBHOOK_URL", "")
    PORT = int(os.getenv("PORT", "8080"))
    
    if WEBHOOK_URL:
        app.run_webhook(listen="0.0.0.0", port=PORT, webhook_url=f"{WEBHOOK_URL}/webhook", drop_pending_updates=True)
    else:
        app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
