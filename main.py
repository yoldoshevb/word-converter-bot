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
from PyPDF2 import PdfReader
from PIL import Image
import pytesseract
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

# =============== LOGGING ===============
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# =============== KONFIGURATSIYA ===============
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "8805822154:AAHdDCps2cumRbpQWm3iw5bEzg0vTjmMOdQ")
ADMIN_ID = 8805822154
ADMIN_USERNAME = "@yoldoshev_3"
BOT_NAME = "📄 Professional File Converter"
MAX_FILE_SIZE = 50 * 1024 * 1024

# =============== SIFATLI WORD HUJJAT ===============
def create_quality_document(title: str, original_name: str = "") -> Document:
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
    style.paragraph_format.space_after = Pt(6)
    style.paragraph_format.line_spacing = 1.15
    
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
    line_para.paragraph_format.space_before = Pt(12)
    line_para.paragraph_format.space_after = Pt(12)
    line_run = line_para.add_run("─" * 70)
    line_run.font.size = Pt(6)
    line_run.font.color.rgb = RGBColor(200, 200, 200)
    
    return doc

def clean_text(text: str) -> str:
    """Suv belgilari va keraksiz belgilarni tozalash"""
    import re
    
    if not text or len(text.strip()) < 2:
        return ""
    
    text = text.strip()
    
    # Faqat maxsus belgilar bo'lsa
    if not any(c.isalnum() for c in text) and len(text) < 15:
        return ""
    
    # Takrorlanuvchi belgilar
    if re.match(r'^[\.\-\s\=\_\~\|]{5,}$', text):
        return ""
    
    # URL yoki email formatidagi suv belgilari
    if re.match(r'^(www\.|http|@)', text) and len(text) < 20:
        return ""
    
    # Ko'p sonli bo'sh joylar
    text = re.sub(r'\s+', ' ', text)
    
    return text

def add_formatted_text(doc: Document, text: str):
    """Formatlangan matn qo'shish"""
    paragraphs = text.split('\n')
    for para_text in paragraphs:
        cleaned = clean_text(para_text)
        if cleaned:
            # Juda uzun satrlarni paragraf qilish
            if len(cleaned) > 200:
                # Gap bo'yicha bo'lish
                sentences = cleaned.replace('. ', '.\n').split('\n')
                for sentence in sentences:
                    if sentence.strip():
                        doc.add_paragraph(sentence.strip())
            else:
                doc.add_paragraph(cleaned)

# =============== RASMDAN MATN (SIFATLI OCR) ===============
def image_ocr_quality(image: Image.Image) -> str:
    """Sifatli OCR"""
    try:
        # Katta rasmlarni kichiklashtirish
        max_size = 2500
        if image.width > max_size or image.height > max_size:
            ratio = min(max_size / image.width, max_size / image.height)
            new_size = (int(image.width * ratio), int(image.height * ratio))
            image = image.resize(new_size, Image.LANCZOS)
        
        # Kulrang formatga o'tkazish
        if image.mode != 'L':
            image = image.convert('L')
        
        # Kontrast oshirish
        from PIL import ImageEnhance
        enhancer = ImageEnhance.Contrast(image)
        image = enhancer.enhance(2.5)
        
        # Keskinlik oshirish
        enhancer = ImageEnhance.Sharpness(image)
        image = enhancer.enhance(2.0)
        
        # OCR — avval inglizcha, topilmasa ruscha
        text = pytesseract.image_to_string(image, lang='eng+rus+uzb', config='--psm 6')
        
        if not text.strip():
            # Boshqa konfiguratsiya bilan qayta urinish
            text = pytesseract.image_to_string(image, lang='eng+rus+uzb', config='--psm 3')
        
        return text.strip()
    except Exception as e:
        logger.error(f"OCR xatolik: {e}")
        return ""

# =============== PDF ICHIDAGI RASMLARNI AJRATISH ===============
def extract_images_from_pdf(pdf_content: bytes) -> list:
    """PDF ichidagi rasmlarni ajratib olish va OCR qilish"""
    try:
        from pdf2image import convert_from_bytes
        from PyPDF2 import PdfReader
        
        images_text = []
        
        # pdf2image orqali har bir sahifani rasmga aylantirish
        pdf_images = convert_from_bytes(
            pdf_content,
            dpi=200,
            fmt='jpeg',
            thread_count=2
        )
        
        for page_num, image in enumerate(pdf_images, 1):
            text = image_ocr_quality(image)
            if text:
                images_text.append((page_num, text))
        
        return images_text
    except Exception as e:
        logger.error(f"PDF rasm ajratish xatolik: {e}")
        return []

# =============== PDF → WORD (KENGAYTIRILGAN) ===============
def pdf_to_word_pro(pdf_content: bytes, original_name: str) -> BytesIO:
    """Professional PDF to Word — matn + jadvallar + rasmlar"""
    doc = create_quality_document("📄 PDF Hujjat", original_name)
    
    temp_pdf = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
    temp_pdf.write(pdf_content)
    temp_pdf.close()
    
    try:
        # 1. Avval pdfplumber orqali matn va jadvallarni olish
        with pdfplumber.open(temp_pdf.name) as pdf:
            total_pages = len(pdf.pages)
            
            for page_num, page in enumerate(pdf.pages, 1):
                if total_pages > 1:
                    doc.add_heading(f'Sahifa {page_num}', level=1)
                
                # Matn olish — yaxshilangan sozlamalar
                text = page.extract_text(
                    x_tolerance=3,
                    y_tolerance=3,
                    keep_blank_chars=False,
                    use_text_flow=True,
                    extra_attrs=[]
                )
                
                if text and len(text.strip()) > 20:
                    add_formatted_text(doc, text)
                
                # Jadvallar
                tables = page.extract_tables()
                if tables:
                    for table_data in tables:
                        if table_data:
                            has_content = any(
                                cell and str(cell).strip()
                                for row in table_data
                                for cell in row if cell
                            )
                            
                            if has_content and len(table_data) > 1:
                                rows = len(table_data)
                                cols = max(len(row) for row in table_data if row)
                                
                                if cols > 0:
                                    doc.add_paragraph()
                                    table = doc.add_table(rows=rows, cols=cols)
                                    table.style = 'Light Grid Accent 1'
                                    
                                    for i, row in enumerate(table_data):
                                        for j in range(cols):
                                            if j < len(row) and row[j]:
                                                cell_text = clean_text(str(row[j]))
                                                if cell_text:
                                                    table.cell(i, j).text = cell_text
                                    
                                    doc.add_paragraph()
                
                if page_num < total_pages:
                    doc.add_page_break()
        
        # 2. PDF dan rasmlarni ajratib olish va OCR qilish
        try:
            from pdf2image import convert_from_bytes
            
            pdf_images = convert_from_bytes(
                pdf_content,
                dpi=200,
                fmt='jpeg',
                thread_count=1
            )
            
            # Har bir sahifani rasm sifatida tekshirish
            for page_num, image in enumerate(pdf_images, 1):
                ocr_text = image_ocr_quality(image)
                
                if ocr_text and len(ocr_text.strip()) > 30:
                    # Agar pdfplumber matn topmagan bo'lsa yoki kam matn topgan bo'lsa
                    doc.add_heading(f'Sahifa {page_num} (Rasm skan)', level=1)
                    add_formatted_text(doc, ocr_text)
                    
                    if page_num < len(pdf_images):
                        doc.add_page_break()
        except:
            logger.warning("pdf2image ishlamadi — faqat matn olinadi")
        
    finally:
        os.unlink(temp_pdf.name)
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# =============== EXCEL → WORD ===============
def excel_to_word_quality(file_content: bytes, file_ext: str, original_name: str) -> BytesIO:
    import pandas as pd
    
    doc = create_quality_document("📊 Excel Ma'lumotlari", original_name)
    
    try:
        if file_ext == 'csv':
            df = pd.read_csv(BytesIO(file_content), encoding='utf-8')
        else:
            df = pd.read_excel(BytesIO(file_content))
        
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
    except Exception as e:
        doc.add_paragraph(f"Xatolik: {str(e)[:100]}")
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# =============== MATN → WORD ===============
def text_to_word_quality(text: str) -> BytesIO:
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

# =============== HTML → WORD ===============
def html_to_word_quality(file_content: bytes, original_name: str) -> BytesIO:
    from bs4 import BeautifulSoup
    
    doc = create_quality_document("🌐 HTML Hujjat", original_name)
    
    html = file_content.decode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')
    
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

📄 PDF hujjatlar (matn + rasm + jadvallar)
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
PDF fayl yuboring. Matn, jadvallar va rasmlar qayta ishlanadi.
• Oddiy PDF: matn va jadvallar olinadi
• Skanerlangan PDF: OCR orqali matn ajratiladi

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
• .pdf — PDF (matn + rasm + jadvallar)
• .xlsx, .xls, .csv — Excel jadvallari
• .html, .htm — Web sahifalar

📚 Kitoblar:
• .epub — Elektron kitoblar

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
    
    msg = await update.message.reply_text(f"⏳ {file_name} qayta ishlanmoqda...\nBu biroz vaqt olishi mumkin.")
    
    try:
        file = await document.get_file()
        content = await file.download_as_bytearray()
        output_name = file_name.rsplit('.', 1)[0] + '.docx'
        
        if file_ext == 'pdf':
            msg = await msg.edit_text("📄 PDF qayta ishlanmoqda: matn va rasmlar olinmoqda...")
            output = pdf_to_word_pro(content, file_name)
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
        
        image = Image.open(BytesIO(content))
        text = image_ocr_quality(image)
        
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
                "💡 Aniqroq, bosma matnli rasm yuboring."
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
