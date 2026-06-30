import os
import logging
import tempfile
from datetime import datetime
from io import BytesIO
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
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
ADMIN_ID = 8805822154  # ⚠️ BU TOKEN EMAS, BU SIZNING TELEGRAM ID'INGIZ BO'LISHI KERAK!
ADMIN_USERNAME = "@yoldoshev_3"
BOT_NAME = "📄 Professional File Converter"
BOT_VERSION = "3.0.0"
MAX_FILE_SIZE = 50 * 1024 * 1024

# =============== STATISTIKA ===============
_stats = {
    "total_conversions": 0,
    "pdf_count": 0,
    "excel_count": 0,
    "text_count": 0,
    "image_count": 0,
    "html_count": 0,
    "epub_count": 0,
    "other_count": 0,
    "users": set(),
    "start_date": datetime.now()
}

def update_stats(file_type: str, user_id: int):
    """Statistikani yangilash"""
    _stats["total_conversions"] += 1
    _stats["users"].add(user_id)
    
    stats_map = {
        "pdf": "pdf_count", "xlsx": "excel_count", "xls": "excel_count",
        "csv": "excel_count", "html": "html_count", "htm": "html_count",
        "epub": "epub_count", "image": "image_count", "text": "text_count"
    }
    
    key = stats_map.get(file_type, "other_count")
    _stats[key] += 1

# =============== PROFESSIONAL WORD HUJJAT ===============
def create_professional_document(title: str, content_type: str, original_name: str = "") -> Document:
    """Professional formatdagi Word hujjat yaratish"""
    doc = Document()
    
    # Sahifa sozlamalari
    section = doc.sections[0]
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(3)
    section.right_margin = Cm(3)
    
    # Header
    header = section.header
    header_para = header.paragraphs[0]
    header_para.text = f"📄 {BOT_NAME} | Rasmiy hujjat"
    header_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for run in header_para.runs:
        run.font.size = Pt(8)
    
    # Footer
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.text = f"📅 Yaratilgan sana: {datetime.now().strftime('%d.%m.%Y %H:%M')} | © {BOT_NAME}"
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in footer_para.runs:
        run.font.size = Pt(8)
    
    # Asosiy sarlavha
    main_title = doc.add_heading(f'{title}', level=0)
    main_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Metadata
    doc.add_paragraph()
    meta_para = doc.add_paragraph()
    meta_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta_run = meta_para.add_run(f"📋 Hujjat turi: {content_type} | 📅 Sana: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    meta_run.font.size = Pt(10)
    
    if original_name:
        meta_para2 = doc.add_paragraph()
        meta_para2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        meta_run2 = meta_para2.add_run(f"📎 Asl fayl: {original_name}")
        meta_run2.font.size = Pt(9)
    
    # Ajratuvchi chiziq
    doc.add_paragraph("━" * 60)
    doc.add_paragraph()
    
    return doc

def add_content_to_doc(doc: Document, text: str):
    """Matnni hujjatga qo'shish"""
    paragraphs = text.split('\n')
    for para in paragraphs:
        if para.strip():
            p = doc.add_paragraph(para.strip())
            p.style.font.size = Pt(12)

# =============== PDF → WORD ===============
def pdf_to_word(pdf_content: bytes, original_name: str) -> BytesIO:
    """PDF faylni Word formatiga o'tkazish"""
    doc = create_professional_document("📄 PDF Hujjat Konvertatsiyasi", "PDF → Word", original_name)
    
    temp_pdf = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
    temp_pdf.write(pdf_content)
    temp_pdf.close()
    
    try:
        with pdfplumber.open(temp_pdf.name) as pdf:
            total_pages = len(pdf.pages)
            doc.add_paragraph(f"📑 Jami sahifalar soni: {total_pages}")
            doc.add_paragraph("━" * 60)
            doc.add_paragraph()
            
            for page_num, page in enumerate(pdf.pages, 1):
                text = page.extract_text()
                if text:
                    doc.add_heading(f'📖 SAHIFA {page_num} / {total_pages}', level=1)
                    add_content_to_doc(doc, text)
                
                tables = page.extract_tables()
                if tables:
                    for table_num, table_data in enumerate(tables, 1):
                        if table_data and len(table_data) > 0:
                            doc.add_heading(f'📊 Jadval {table_num}', level=2)
                            rows = len(table_data)
                            cols = len(table_data[0]) if table_data[0] else 1
                            table = doc.add_table(rows=rows, cols=cols)
                            table.style = 'Table Grid'
                            
                            for i, row in enumerate(table_data):
                                for j, cell in enumerate(row):
                                    if cell and j < cols:
                                        table.cell(i, j).text = str(cell)
                            
                            doc.add_paragraph()
                
                if page_num < total_pages:
                    doc.add_page_break()
    finally:
        os.unlink(temp_pdf.name)
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# =============== EXCEL → WORD ===============
def excel_to_word(file_content: bytes, file_ext: str, original_name: str) -> BytesIO:
    """Excel faylni Word formatiga o'tkazish"""
    import pandas as pd
    
    doc = create_professional_document("📊 Excel Ma'lumotlari", f"{file_ext.upper()} → Word", original_name)
    
    if file_ext == 'csv':
        df = pd.read_csv(BytesIO(file_content))
    else:
        df = pd.read_excel(BytesIO(file_content))
    
    doc.add_paragraph(f"📈 Qatorlar soni: {len(df)}")
    doc.add_paragraph(f"📊 Ustunlar soni: {len(df.columns)}")
    doc.add_paragraph("━" * 60)
    doc.add_paragraph()
    
    table = doc.add_table(rows=len(df) + 1, cols=len(df.columns))
    table.style = 'Table Grid'
    
    for j, col in enumerate(df.columns):
        cell = table.cell(0, j)
        cell.text = str(col)
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(10)
    
    for i, (_, row) in enumerate(df.iterrows()):
        for j, value in enumerate(row):
            cell = table.cell(i + 1, j)
            cell.text = str(value) if pd.notna(value) else ""
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# =============== HTML → WORD ===============
def html_to_word(file_content: bytes, original_name: str) -> BytesIO:
    """HTML faylni Word formatiga o'tkazish"""
    from bs4 import BeautifulSoup
    
    doc = create_professional_document("🌐 HTML Hujjat Konvertatsiyasi", "HTML → Word", original_name)
    
    html = file_content.decode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')
    text = soup.get_text()
    
    add_content_to_doc(doc, text)
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# =============== EPUB → WORD ===============
def epub_to_word(file_content: bytes, original_name: str) -> BytesIO:
    """EPUB faylni Word formatiga o'tkazish"""
    from ebooklib import epub
    from bs4 import BeautifulSoup
    
    doc = create_professional_document("📚 Elektron Kitob Konvertatsiyasi", "EPUB → Word", original_name)
    
    temp_epub = tempfile.NamedTemporaryFile(suffix='.epub', delete=False)
    temp_epub.write(file_content)
    temp_epub.close()
    
    try:
        book = epub.read_epub(temp_epub.name)
        
        chapter_num = 0
        for item in book.get_items():
            if item.get_type() == 9:
                chapter_num += 1
                soup = BeautifulSoup(item.get_content(), 'html.parser')
                text = soup.get_text()
                
                if text.strip():
                    doc.add_heading(f'📖 BO\'LIM {chapter_num}', level=1)
                    add_content_to_doc(doc, text)
                    doc.add_page_break()
    finally:
        os.unlink(temp_epub.name)
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# =============== MATN → WORD ===============
def text_to_word(text: str) -> BytesIO:
    """Matnni Word formatiga o'tkazish"""
    doc = create_professional_document("📝 Matn Konvertatsiyasi", "Matn → Word")
    doc.add_paragraph(f"📏 Belgilar soni: {len(text)}")
    doc.add_paragraph("━" * 60)
    doc.add_paragraph()
    add_content_to_doc(doc, text)
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# =============== RASM OCR ===============
def image_to_text(image_content: bytes) -> str:
    """Rasmdan matn olish"""
    try:
        from PIL import Image
        import pytesseract
        
        image = Image.open(BytesIO(image_content))
        text = pytesseract.image_to_string(image, lang='eng+rus+uzb')
        return text
    except:
        return ""

# =============== BOT HANDLERLARI ===============

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Start komandasi"""
    user = update.effective_user
    
    message = f"""✨ *Assalomu alaykum, {user.first_name}!*

💼 *{BOT_NAME}* ga xush kelibsiz!

┏━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
┃   📋 PROFESSIONAL FILE    ┃
┃   CONVERTER BOT           ┃
┗━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

📌 *Quyidagi formatlarni Word (.docx) ga o'tkazadi:*

📄 • PDF hujjatlar
📊 • Excel jadvallari (.xlsx, .xls, .csv)
🌐 • HTML sahifalar
📚 • EPUB elektron kitoblar
📝 • Matn fayllari (.txt, .md, .json, .xml)
🖼 • Rasmlardagi matnlar (OCR)

⚙️ *Foydalanish tartibi:*
1️⃣ Botga fayl yoki matn yuboring
2️⃣ Konvertatsiya avtomatik amalga oshiriladi
3️⃣ Tayyor Word faylni yuklab oling

📌 *Cheklovlar:*
🔹 Maksimal fayl hajmi: 50 MB
🔹 Bot 24/7 rejimida ishlaydi
🔹 Xizmat mutlaqo *bepul* ✅

📋 *Buyruqlar:*
/start - Bosh menyu
/help - Yordam
/formats - Formatlar ro'yxati
/about - Bot haqida

👨‍💼 *Admin:* {ADMIN_USERNAME}"""

    await update.message.reply_text(message, parse_mode='Markdown')

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Yordam komandasi"""
    message = f"""📖 *YORDAM BO'LIMI*

┏━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
┃   📚 FOYDALANISH BO'YICHA ┃
┃   QO'LLANMA               ┃
┗━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

📝 *1. MATN YUBORISH*
Oddiy matn yozib yuboring — avtomatik Word faylga o'tkaziladi.

📄 *2. PDF FAYL YUBORISH*
PDF faylni yuboring. Matn va jadvallar to'liq saqlanadi.

📊 *3. EXCEL FAYL YUBORISH*
.xlsx, .xls yoki .csv formatdagi fayllarni yuboring.

🌐 *4. HTML FAYL YUBORISH*
.html fayllarni yuboring.

📚 *5. EPUB FAYL YUBORISH*
Elektron kitoblarni yuboring.

🖼 *6. RASM YUBORISH*
Matnli rasm yuboring. OCR texnologiyasi orqali matn ajratib olinadi.

⚠️ *Muhim eslatmalar:*
• Fayl hajmi 50 MB dan oshmasligi kerak
• Konvertatsiya vaqti fayl hajmiga bog'liq
• Barcha fayllar .docx formatiga o'tkaziladi

📞 *Texnik yordam:* {ADMIN_USERNAME}"""

    await update.message.reply_text(message, parse_mode='Markdown')

async def formats_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Formatlar ro'yxati"""
    message = f"""📋 *QO'LLAB-QUVVATLANADIGAN FORMATLAR*

┏━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
┃   📎 FORMATLAR RO'YXATI   ┃
┗━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

📄 *Hujjat formatlari:*
• .pdf - PDF hujjatlar
• .xlsx - Microsoft Excel (2007+)
• .xls - Microsoft Excel (97-2003)
• .csv - CSV jadvallar
• .html, .htm - Web sahifalar

📚 *Elektron kitoblar:*
• .epub - EPUB format

📝 *Matn formatlari:*
• Oddiy matn xabarlar
• .txt - Matn fayllari
• .md - Markdown
• .json - JSON ma'lumotlar
• .xml - XML ma'lumotlar

🖼 *Rasm formatlari (OCR):*
• .jpg, .jpeg
• .png
• .bmp

✅ *Barcha formatlar → .docx (Microsoft Word)*"""

    await update.message.reply_text(message, parse_mode='Markdown')

async def admin_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Admin panel"""
    user_id = update.effective_user.id
    
    if user_id != ADMIN_ID:
        await update.message.reply_text("⛔ Ushbu buyruq faqat admin uchun mavjud.")
        return
    
    uptime = datetime.now() - _stats["start_date"]
    days = uptime.days
    hours = uptime.seconds // 3600
    minutes = (uptime.seconds % 3600) // 60
    
    message = f"""👑 *ADMIN PANEL*

┏━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
┃   ⚙️ TIZIM BOSHQARUVI     ┃
┗━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

🟢 *Tizim holati:* Aktiv
📌 *Bot versiyasi:* {BOT_VERSION}
⏱ *Ishlash vaqti:* {days} kun, {hours} soat, {minutes} daqiqa

📊 *Statistika:*
━━━━━━━━━━━━━━━━━━━━━
🔄 Jami konvertatsiyalar: *{_stats['total_conversions']}*
👥 Foydalanuvchilar soni: *{len(_stats['users'])}*

📈 *Formatlar bo'yicha:*
📄 PDF: *{_stats['pdf_count']}*
📊 Excel: *{_stats['excel_count']}*
🌐 HTML: *{_stats['html_count']}*
📚 EPUB: *{_stats['epub_count']}*
📝 Matn: *{_stats['text_count']}*
🖼 Rasm: *{_stats['image_count']}*
📎 Boshqa: *{_stats['other_count']}*

👨‍💼 *Admin:* {ADMIN_USERNAME}"""

    await update.message.reply_text(message, parse_mode='Markdown')

async def about_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Bot haqida"""
    message = f"""ℹ️ *BOT HAQIDA*

┏━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
┃   📌 MA'LUMOT             ┃
┗━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

🏷 *Nomi:* {BOT_NAME}
📌 *Versiya:* {BOT_VERSION}


👨‍💼 *Admin:* {ADMIN_USERNAME}
📞 *Aloqa:* {ADMIN_USERNAME}

© 2026 {BOT_NAME}. Barcha huquqlar himoyalangan."""

    await update.message.reply_text(message, parse_mode='Markdown')

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Fayllarni qabul qilish"""
    document = update.message.document
    file_name = document.file_name
    user_id = update.effective_user.id
    
    # Fayl hajmi tekshiruvi
    if document.file_size > MAX_FILE_SIZE:
        await update.message.reply_text(
            "❌ *Xatolik:* Fayl hajmi 50 MB dan oshmasligi kerak.",
            parse_mode='Markdown'
        )
        return
    
    # Fayl formatini aniqlash
    file_ext = file_name.lower().split('.')[-1] if '.' in file_name else ''
    supported_extensions = ['pdf', 'xlsx', 'xls', 'csv', 'html', 'htm', 'epub', 'txt', 'md', 'json', 'xml']
    
    if file_ext not in supported_extensions:
        await update.message.reply_text(
            f"❌ *Xatolik:* .{file_ext} formati qo'llab-quvvatlanmaydi.\n"
            f"📋 Formatlar ro'yxati: /formats",
            parse_mode='Markdown'
        )
        return
    
    processing_msg = await update.message.reply_text(
        f"⏳ *Fayl qayta ishlanmoqda...*\n\n"
        f"📎 Nomi: *{file_name}*\n"
        f"📦 Hajmi: *{document.file_size // 1024} KB*\n"
        f"⏱ Iltimos, kuting...",
        parse_mode='Markdown'
    )
    
    try:
        file = await document.get_file()
        file_content = await file.download_as_bytearray()
        
        output_file = None
        output_name = file_name.rsplit('.', 1)[0] + '.docx'
        
        if file_ext == 'pdf':
            output_file = pdf_to_word(file_content, file_name)
            update_stats("pdf", user_id)
        elif file_ext in ['xlsx', 'xls', 'csv']:
            output_file = excel_to_word(file_content, file_ext, file_name)
            update_stats(file_ext, user_id)
        elif file_ext in ['html', 'htm']:
            output_file = html_to_word(file_content, file_name)
            update_stats("html", user_id)
        elif file_ext == 'epub':
            output_file = epub_to_word(file_content, file_name)
            update_stats("epub", user_id)
        else:
            text = file_content.decode('utf-8')
            output_file = text_to_word(text)
            update_stats("text", user_id)
        
        await processing_msg.delete()
        
        await update.message.reply_document(
            document=output_file,
            filename=output_name,
            caption=f"✅ *Konvertatsiya muvaffaqiyatli yakunlandi!*\n\n"
                   f"📎 Asl fayl: *{file_name}*\n"
                   f"📄 Yangi fayl: *{output_name}*\n"
                   f"🔤 Format: Word (.docx)",
            parse_mode='Markdown'
        )
        
    except Exception as e:
        await processing_msg.delete()
        logger.error(f"Xatolik: {e}")
        await update.message.reply_text(
            f"❌ *Konvertatsiya jarayonida xatolik!*\n\n"
            f"📎 Fayl: *{file_name}*\n"
            f"🔍 Xato: {str(e)[:100]}\n\n"
            f"📞 Adminga murojaat qiling: {ADMIN_USERNAME}",
            parse_mode='Markdown'
        )

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Rasmlarni qabul qilish"""
    user_id = update.effective_user.id
    
    processing_msg = await update.message.reply_text(
        "🖼 *Rasm qayta ishlanmoqda...*\n\n"
        "🔍 OCR texnologiyasi orqali matn ajratib olinmoqda.\n"
        "⏱ Iltimos, kuting...",
        parse_mode='Markdown'
    )
    
    try:
        photo = update.message.photo[-1]
        file = await photo.get_file()
        image_content = await file.download_as_bytearray()
        
        text = image_to_text(image_content)
        
        if text and len(text.strip()) > 10:
            output_file = text_to_word(text)
            update_stats("image", user_id)
            
            await processing_msg.delete()
            
            await update.message.reply_document(
                document=output_file,
                filename="rasmdagi_matn.docx",
                caption=f"✅ *Rasmdagi matn muvaffaqiyatli ajratib olindi!*\n\n"
                       f"📏 Matn hajmi: *{len(text)}* belgi\n"
                       f"📄 Format: Word (.docx)\n"
                       f"📁 Fayl: rasmdagi_matn.docx",
                parse_mode='Markdown'
            )
        else:
            await processing_msg.delete()
            await update.message.reply_text(
                "⚠️ *Rasmdan matn topilmadi!*\n\n"
                "📌 *Mumkin bo'lgan sabablar:*\n"
                "• Rasmda matn mavjud emas\n"
                "• Rasm sifati past\n"
                "• Qo'l yozuvi (faqat bosma matn)\n\n"
                "💡 Iltimos, aniqroq rasm yuboring.",
                parse_mode='Markdown'
            )
    
    except Exception as e:
        await processing_msg.delete()
        logger.error(f"OCR xatolik: {e}")
        await update.message.reply_text(
            f"❌ *Rasmni qayta ishlashda xatolik!*\n"
            f"📞 Adminga murojaat qiling: {ADMIN_USERNAME}",
            parse_mode='Markdown'
        )

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Matn xabarlarni qabul qilish"""
    text = update.message.text
    user_id = update.effective_user.id
    
    if len(text) < 10:
        await update.message.reply_text(
            "⚠️ *Matn juda qisqa!*\n\n"
            "📝 Kamida 10 ta belgidan iborat matn yuboring.\n"
            "📖 Yordam uchun: /help",
            parse_mode='Markdown'
        )
        return
    
    processing_msg = await update.message.reply_text(
        "⏳ *Matn qayta ishlanmoqda...*",
        parse_mode='Markdown'
    )
    
    try:
        output_file = text_to_word(text)
        update_stats("text", user_id)
        
        await processing_msg.delete()
        
        await update.message.reply_document(
            document=output_file,
            filename="matn.docx",
            caption=f"✅ *Matn muvaffaqiyatli Word formatiga o'tkazildi!*\n\n"
                   f"📏 Belgilar soni: *{len(text)}*\n"
                   f"📄 Fayl nomi: *matn.docx*\n"
                   f"🔤 Format: Word (.docx)",
            parse_mode='Markdown'
        )
        
    except Exception as e:
        await processing_msg.delete()
        logger.error(f"Matn xatolik: {e}")
        await update.message.reply_text(
            "❌ *Matnni qayta ishlashda xatolik yuz berdi.*",
            parse_mode='Markdown'
        )

async def error_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Xatoliklarni boshqarish"""
    logger.error(f"Xatolik: {context.error}")
    
    try:
        if update and update.effective_message:
            await update.effective_message.reply_text(
                "❌ *Kutilmagan xatolik yuz berdi.*\n\n"
                "🔄 Iltimos, qaytadan urinib ko'ring.\n"
                f"📞 Yordam uchun: {ADMIN_USERNAME}",
                parse_mode='Markdown'
            )
    except:
        pass

# =============== ASOSIY FUNKSIYA ===============
def main():
    """Botni ishga tushirish"""
    
    if not TOKEN:
        logger.error("❌ TOKEN o'rnatilmagan!")
        return
    
    logger.info(f"🤖 {BOT_NAME} ishga tushirilmoqda...")
    logger.info(f"📌 Versiya: {BOT_VERSION}")
    logger.info(f"👨‍💼 Admin: {ADMIN_USERNAME}")
    
    app = Application.builder().token(TOKEN).build()
    
    # Command handlerlar
    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("formats", formats_command))
    app.add_handler(CommandHandler("admin", admin_command))
    app.add_handler(CommandHandler("about", about_command))
    
    # Message handlerlar
    app.add_handler(MessageHandler(filters.DOCUMENT, handle_document))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    
    # Error handler
    app.add_error_handler(error_handler)
    
    # Webhook yoki polling
    WEBHOOK_URL = os.getenv("WEBHOOK_URL", "")
    PORT = int(os.getenv("PORT", "8080"))
    
    if WEBHOOK_URL:
        logger.info(f"🌐 Webhook rejimi: {WEBHOOK_URL}")
        app.run_webhook(
            listen="0.0.0.0",
            port=PORT,
            webhook_url=f"{WEBHOOK_URL}/webhook",
            drop_pending_updates=True
        )
    else:
        logger.info("📡 Polling rejimi")
        app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
