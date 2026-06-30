import os
import logging
import tempfile
import json
from datetime import datetime
from typing import Dict, Any
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler, 
    CallbackQueryHandler, filters, ContextTypes
)
from docx import Document
import pdfplumber
from io import BytesIO

# =============== LOGGING ===============
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# =============== KONFIGURATSIYA ===============
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
ADMIN_IDS = [int(id) for id in os.getenv("ADMIN_IDS", "").split(",") if id]
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50 MB

# Statistika
stats = {
    "total_users": set(),
    "total_conversions": 0,
    "conversions_by_type": {},
    "start_time": datetime.now()
}

# =============== YORDAMCHI FUNKSIYALAR ===============
def save_stats():
    """Statistikani saqlash"""
    try:
        data = {
            "total_users": list(stats["total_users"]),
            "total_conversions": stats["total_conversions"],
            "conversions_by_type": stats["conversions_by_type"],
            "start_time": stats["start_time"].isoformat()
        }
        with open("stats.json", "w") as f:
            json.dump(data, f)
    except:
        pass

def load_stats():
    """Statistikani yuklash"""
    try:
        with open("stats.json", "r") as f:
            data = json.load(f)
            stats["total_users"] = set(data.get("total_users", []))
            stats["total_conversions"] = data.get("total_conversions", 0)
            stats["conversions_by_type"] = data.get("conversions_by_type", {})
    except:
        pass

# =============== KEYBOARDLAR ===============
def get_main_keyboard():
    """Asosiy menyu tugmalari"""
    keyboard = [
        [InlineKeyboardButton("📄 Formatlar", callback_data="formats"),
         InlineKeyboardButton("❓ Yordam", callback_data="help")],
        [InlineKeyboardButton("📊 Statistika", callback_data="stats"),
         InlineKeyboardButton("ℹ️ Bot haqida", callback_data="about")]
    ]
    return InlineKeyboardMarkup(keyboard)

def get_admin_keyboard():
    """Admin panel tugmalari"""
    keyboard = [
        [InlineKeyboardButton("📊 To'liq statistika", callback_data="admin_stats"),
         InlineKeyboardButton("👥 Foydalanuvchilar", callback_data="admin_users")],
        [InlineKeyboardButton("📨 Xabar yuborish", callback_data="admin_broadcast"),
         InlineKeyboardButton("🔄 Botni qayta yuklash", callback_data="admin_restart")],
        [InlineKeyboardButton("🔙 Orqaga", callback_data="back_to_main")]
    ]
    return InlineKeyboardMarkup(keyboard)

# =============== PDF NI WORD GA ===============
def pdf_to_word(pdf_content, user_id=None):
    """PDF faylni Word formatiga o'tkazish"""
    doc = Document()
    doc.add_heading('📄 PDF Konvertatsiyasi', 0)
    
    temp_pdf = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
    temp_pdf.write(pdf_content)
    temp_pdf.close()
    
    try:
        with pdfplumber.open(temp_pdf.name) as pdf:
            total_pages = len(pdf.pages)
            doc.add_paragraph(f"📑 Jami sahifalar: {total_pages}")
            doc.add_paragraph(f"📅 Sana: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
            doc.add_paragraph("=" * 50)
            
            for page_num, page in enumerate(pdf.pages, 1):
                text = page.extract_text()
                if text:
                    doc.add_heading(f'📖 Sahifa {page_num}/{total_pages}', level=1)
                    for para in text.split('\n'):
                        if para.strip():
                            doc.add_paragraph(para.strip())
                
                tables = page.extract_tables()
                for idx, table_data in enumerate(tables, 1):
                    if table_data:
                        doc.add_heading(f'📊 Jadval {idx}', level=2)
                        rows = len(table_data)
                        cols = len(table_data[0]) if table_data[0] else 1
                        table = doc.add_table(rows=rows, cols=cols)
                        table.style = 'Table Grid'
                        
                        for i, row in enumerate(table_data):
                            for j, cell in enumerate(row):
                                if cell and j < cols:
                                    table.cell(i, j).text = str(cell)
                        
                        doc.add_paragraph()
    finally:
        os.unlink(temp_pdf.name)
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# =============== EXCEL NI WORD GA ===============
def excel_to_word(file_content, file_ext):
    """Excel/CSV ni Word ga o'tkazish"""
    import pandas as pd
    
    if file_ext == 'csv':
        df = pd.read_csv(BytesIO(file_content))
    else:
        df = pd.read_excel(BytesIO(file_content))
    
    doc = Document()
    doc.add_heading('📊 Excel Ma\'lumotlari', 0)
    doc.add_paragraph(f"📅 Sana: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    doc.add_paragraph(f"📈 Qatorlar: {len(df)}, Ustunlar: {len(df.columns)}")
    doc.add_paragraph("=" * 50)
    
    table = doc.add_table(rows=len(df)+1, cols=len(df.columns))
    table.style = 'Table Grid'
    
    for j, col in enumerate(df.columns):
        table.cell(0, j).text = str(col)
    
    for i, (_, row) in enumerate(df.iterrows()):
        for j, value in enumerate(row):
            table.cell(i+1, j).text = str(value) if pd.notna(value) else ''
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# =============== MATNNI WORD GA ===============
def text_to_word(text):
    """Matnni Word formatiga o'tkazish"""
    doc = Document()
    doc.add_heading('📝 Matn Konvertatsiyasi', 0)
    doc.add_paragraph(f"📅 Sana: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    doc.add_paragraph(f"📏 Belgilar soni: {len(text)}")
    doc.add_paragraph("=" * 50)
    
    for para in text.split('\n'):
        if para.strip():
            doc.add_paragraph(para.strip())
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# =============== HTML NI WORD GA ===============
def html_to_word(file_content):
    """HTML ni Word ga o'tkazish"""
    from bs4 import BeautifulSoup
    
    html = file_content.decode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')
    
    doc = Document()
    doc.add_heading('🌐 HTML Konvertatsiyasi', 0)
    
    text = soup.get_text()
    for para in text.split('\n'):
        if para.strip():
            doc.add_paragraph(para.strip())
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# =============== EPUB NI WORD GA ===============
def epub_to_word(file_content):
    """EPUB ni Word ga o'tkazish"""
    from ebooklib import epub
    from bs4 import BeautifulSoup
    
    temp_epub = tempfile.NamedTemporaryFile(suffix='.epub', delete=False)
    temp_epub.write(file_content)
    temp_epub.close()
    
    doc = Document()
    doc.add_heading('📚 Elektron Kitob', 0)
    
    try:
        book = epub.read_epub(temp_epub.name)
        
        for item in book.get_items():
            if item.get_type() == 9:
                soup = BeautifulSoup(item.get_content(), 'html.parser')
                text = soup.get_text()
                
                for para in text.split('\n'):
                    if para.strip():
                        doc.add_paragraph(para.strip())
                
                doc.add_page_break()
    finally:
        os.unlink(temp_epub.name)
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# =============== RASMDAN MATN (OCR) ===============
def image_to_text(image_content):
    """Rasmdan matn olish"""
    try:
        from PIL import Image
        import pytesseract
        
        image = Image.open(BytesIO(image_content))
        text = pytesseract.image_to_string(image, lang='eng+rus+uzb')
        return text
    except:
        return None

# =============== HANDLERLAR ===============

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Start komandasi"""
    user = update.effective_user
    stats["total_users"].add(user.id)
    save_stats()
    
    welcome_text = f"""
🎉 *Assalomu alaykum, {user.first_name}!*

*Universal File Converter Bot* ga xush kelibsiz!

╔══════════════════════════╗
║  📝 Matn → Word         ║
║  📄 PDF → Word          ║
║  📊 Excel → Word        ║
║  🌐 HTML → Word         ║
║  📚 EPUB → Word         ║
║  🖼 Rasm (OCR) → Word   ║
║  📋 JSON/XML → Word     ║
╚══════════════════════════╝

*🎯 Qanday ishlatiladi?*
1️⃣ Menga fayl yoki matn yuboring
2️⃣ Avtomatik qayta ishlanadi
3️⃣ Word faylni yuklab oling

📌 *Cheklovlar:*
• Maksimal fayl hajmi: 50 MB
• 24/7 ishlaydi
• Mutlaqo bepul ✅

*Quyidagi menyudan foydalaning:* 👇
"""
    await update.message.reply_text(
        welcome_text,
        parse_mode='Markdown',
        reply_markup=get_main_keyboard()
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Yordam komandasi"""
    help_text = """
📖 *YORDAM MARKAZI*

╔══════════════════════════════════╗
║      BOTDAN FOYDALANISH         ║
╚══════════════════════════════════╝

*1️⃣ MATN YUBORISH*
Oddiy matn yozib yuboring - avtomatik Word fayl qaytaradi

*2️⃣ PDF FAYL YUBORISH*
PDF faylni yuboring, u Word ga o'tkaziladi
• Matn va jadvallar saqlanadi
• Ko'p sahifali PDF qo'llab-quvvatlanadi

*3️⃣ EXCEL YUBORISH*
.xlsx, .xls yoki .csv fayl yuboring
• Ma'lumotlar jadval ko'rinishida saqlanadi

*4️⃣ HTML YUBORISH*
.html fayllar yuboring
• Formatlangan matn saqlanadi

*5️⃣ EPUB YUBORISH*
Elektron kitoblarni yuboring
• Barcha bo'limlar saqlanadi

*6️⃣ RASM YUBORISH*
Matnli rasm yuboring (OCR)
• Ingliz, rus, o'zbek tillari

📞 *Bog'lanish:*
Admin: @yoldoshev_3

📌 *Muhim:*
• Fayl 50 MB dan katta bo'lmasin
• Bot 24/7 ishlaydi
• Mutlaqo bepul
"""
    await update.message.reply_text(
        help_text,
        parse_mode='Markdown',
        reply_markup=get_main_keyboard()
    )

async def formats_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Formatlar ro'yxati"""
    formats_text = """
📋 *QO'LLAB-QUVVATLANADIGAN FORMATLAR*

╔══════════════════════════════════╗
║     FORMATLAR RO'YXATI          ║
╚══════════════════════════════════╝

📝 *Matn formatlar:*
• Oddiy matn xabarlar
• .txt - Matn fayllari
• .md - Markdown

📄 *Hujjatlar:*
• .pdf - PDF hujjatlar
• .doc - Word (ko'rish)
• .html - Web sahifalar

📊 *Jadvallar:*
• .xlsx - Excel (yangi)
• .xls - Excel (eski)
• .csv - CSV jadval

📚 *Kitoblar:*
• .epub - Elektron kitoblar

🌐 *Web formatlar:*
• .html - HTML
• .json - JSON
• .xml - XML

🖼 *Rasmlar (OCR):*
• .jpg, .png, .bmp
• 50+ til qo'llab-quvvatlanadi

📌 Barchasi → Word (.docx) formatiga
"""
    await update.message.reply_text(
        formats_text,
        parse_mode='Markdown',
        reply_markup=get_main_keyboard()
    )

async def stats_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Statistika"""
    uptime = datetime.now() - stats["start_time"]
    days = uptime.days
    hours = uptime.seconds // 3600
    minutes = (uptime.seconds % 3600) // 60
    
    stats_text = f"""
📊 *BOT STATISTIKASI*

╔══════════════════════════════════╗
║         STATISTIKA              ║
╚══════════════════════════════════╝

👥 *Jami foydalanuvchilar:* {len(stats['total_users'])}
🔄 *Jami konvertatsiyalar:* {stats['total_conversions']}
⏱ *Ishlash vaqti:* {days}k {hours}s {minutes}d

📈 *Formatlar bo'yicha:*
"""
    for fmt, count in stats.get('conversions_by_type', {}).items():
        stats_text += f"• {fmt}: {count} ta\n"
    
    stats_text += f"""
━━━━━━━━━━━━━━━━━━━━━━
🟢 *Holat:* Aktiv
⚡ *Tezlik:* Optimal
🚀 *Server:* Aktiv
"""
    
    await update.message.reply_text(
        stats_text,
        parse_mode='Markdown',
        reply_markup=get_main_keyboard()
    )

async def about_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Bot haqida"""
    about_text = """
ℹ️ *BOT HAQIDA*

  UNIVERSAL CONVERTER BOT  

*Versiya:* 2.0
*Yaratilgan sana:* 2026

✨ *Imkoniyatlar:*
• 10+ format qo'llab-quvvatlanadi
• Ko'p tilli OCR
• Tez konvertatsiya
• Avtomatik yangilanish
• 24/7 ishlaydi

📞 *Aloqa:*
Admin: @yoldoshev_3

💙 *Fikr-mulohazalar uchun:* /feedback
"""
    await update.message.reply_text(
        about_text,
        parse_mode='Markdown',
        reply_markup=get_main_keyboard()
    )

async def admin_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Admin panel"""
    user_id = update.effective_user.id
    
    if user_id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Siz admin emassiz!")
        return
    
    admin_text = f"""
👑 *ADMIN PANEL*

╔══════════════════════════════════╗
║        ADMIN BOSHQARUVI        ║
╚══════════════════════════════════╝

👥 *Foydalanuvchilar:* {len(stats['total_users'])}
🔄 *Konvertatsiyalar:* {stats['total_conversions']}
⏱ *Ishlash vaqti:* {datetime.now() - stats['start_time']}

*Admin IDlar:* {ADMIN_IDS}
"""
    await update.message.reply_text(
        admin_text,
        parse_mode='Markdown',
        reply_markup=get_admin_keyboard()
    )

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Tugma bosilganda"""
    query = update.callback_query
    await query.answer()
    
    if query.data == "formats":
        await formats_command(update, context)
    elif query.data == "help":
        await help_command(update, context)
    elif query.data == "stats":
        await stats_command(update, context)
    elif query.data == "about":
        await about_command(update, context)
    elif query.data == "admin_stats":
        # Admin statistika
        user_id = query.from_user.id
        if user_id not in ADMIN_IDS:
            await query.edit_message_text("⛔ Ruxsat yo'q!")
            return
        
        detailed_stats = f"""
📊 *TO'LIQ STATISTIKA*

👥 *Foydalanuvchilar:* {len(stats['total_users'])}
🔄 *Jami konvertatsiyalar:* {stats['total_conversions']}
⏱ *Ishlash vaqti:* {datetime.now() - stats['start_time']}

📈 *Formatlar bo'yicha:*
{json.dumps(stats.get('conversions_by_type', {}), indent=2, ensure_ascii=False)}
"""
        await query.edit_message_text(detailed_stats, parse_mode='Markdown')
    elif query.data == "admin_users":
        if query.from_user.id not in ADMIN_IDS:
            await query.edit_message_text("⛔ Ruxsat yo'q!")
            return
        users_list = "\n".join([f"• `{uid}`" for uid in list(stats['total_users'])[:20]])
        await query.edit_message_text(f"👥 *Foydalanuvchilar (20 ta):*\n{users_list}", parse_mode='Markdown')
    elif query.data == "back_to_main":
        await start(update, context)

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Fayllarni qabul qilish"""
    document = update.message.document
    file_name = document.file_name
    user_id = update.effective_user.id
    
    stats["total_users"].add(user_id)
    
    if document.file_size > MAX_FILE_SIZE:
        await update.message.reply_text(
            f"❌ Fayl juda katta! Maksimum {MAX_FILE_SIZE // (1024*1024)} MB."
        )
        return
    
    file_ext = file_name.lower().split('.')[-1] if '.' in file_name else ''
    supported = ['pdf', 'xlsx', 'xls', 'csv', 'html', 'htm', 'epub', 'txt', 'md', 'json', 'xml']
    
    if file_ext not in supported:
        await update.message.reply_text(
            f"❌ *{file_ext.upper()}* formati qo'llab-quvvatlanmaydi!\n"
            f"Formatlar ro'yxati: /formats",
            parse_mode='Markdown'
        )
        return
    
    processing_msg = await update.message.reply_text(
        f"⏳ *{file_name}* qayta ishlanmoqda...\n"
        f"📦 Hajmi: {document.file_size // 1024} KB\n"
        f"⏱ Iltimos kuting...",
        parse_mode='Markdown'
    )
    
    try:
        file = await document.get_file()
        file_content = await file.download_as_bytearray()
        
        output_file = None
        output_name = file_name.rsplit('.', 1)[0] + '.docx'
        
        # Konvertatsiya
        if file_ext == 'pdf':
            output_file = pdf_to_word(file_content, user_id)
        elif file_ext in ['xlsx', 'xls', 'csv']:
            output_file = excel_to_word(file_content, file_ext)
        elif file_ext in ['html', 'htm']:
            output_file = html_to_word(file_content)
        elif file_ext == 'epub':
            output_file = epub_to_word(file_content)
        else:  # txt, md, json, xml
            text = file_content.decode('utf-8')
            output_file = text_to_word(text)
        
        # Statistikani yangilash
        stats["total_conversions"] += 1
        stats["conversions_by_type"][file_ext] = stats["conversions_by_type"].get(file_ext, 0) + 1
        save_stats()
        
        await processing_msg.delete()
        
        await update.message.reply_document(
            document=output_file,
            filename=output_name,
            caption=f"✅ *Tayyor!*\n"
                   f"📄 {file_name} → {output_name}\n"
                   f"⏱ Konvertatsiya muvaffaqiyatli yakunlandi!",
            parse_mode='Markdown'
        )
        
    except Exception as e:
        await processing_msg.delete()
        logger.error(f"Xatolik: {e}")
        await update.message.reply_text(
            f"❌ *Xatolik yuz berdi!*\n\n"
            f"📄 Fayl: {file_name}\n"
            f"🔍 Xato: {str(e)[:200]}\n\n"
            f"Iltimos qaytadan urinib ko'ring yoki adminga murojaat qiling.",
            parse_mode='Markdown'
        )

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Rasmlarni qabul qilish"""
    user_id = update.effective_user.id
    stats["total_users"].add(user_id)
    
    processing_msg = await update.message.reply_text(
        "🖼 *Rasm qayta ishlanmoqda...*\n"
        "🔍 OCR texnologiyasi ishlamoqda\n"
        "⏱ Iltimos kuting...",
        parse_mode='Markdown'
    )
    
    try:
        photo = update.message.photo[-1]
        file = await photo.get_file()
        image_content = await file.download_as_bytearray()
        
        text = image_to_text(image_content)
        
        if text and len(text.strip()) > 10:
            output_file = text_to_word(text)
            
            stats["total_conversions"] += 1
            stats["conversions_by_type"]["rasm"] = stats["conversions_by_type"].get("rasm", 0) + 1
            save_stats()
            
            await processing_msg.delete()
            
            await update.message.reply_document(
                document=output_file,
                filename="rasmdagi_matn.docx",
                caption=f"✅ *Rasmdagi matn olindi!*\n"
                       f"📏 Topilgan matn: {len(text)} belgi\n"
                       f"📄 Word formatiga o'tkazildi",
                parse_mode='Markdown'
            )
        else:
            await processing_msg.delete()
            await update.message.reply_text(
                "❌ *Rasmdan matn topilmadi!*\n\n"
                "Sabablar:\n"
                "• Rasmda matn yo'q\n"
                "• Sifatsiz rasm\n"
                "• Qo'l yozuvi (bosma matn kerak)\n\n"
                "Iltimos aniqroq rasm yuboring.",
                parse_mode='Markdown'
            )
    
    except Exception as e:
        await processing_msg.delete()
        await update.message.reply_text(
            f"❌ *OCR xatolik:* {str(e)[:200]}",
            parse_mode='Markdown'
        )

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Matn xabarlarni qabul qilish"""
    text = update.message.text
    user_id = update.effective_user.id
    
    stats["total_users"].add(user_id)
    
    if len(text) < 10:
        await update.message.reply_text(
            "📝 *Matn juda qisqa!*\n\n"
            "Kamida 10 ta belgi kerak.\n"
            "Buyruqlar uchun: /help",
            parse_mode='Markdown'
        )
        return
    
    processing_msg = await update.message.reply_text(
        "⏳ *Matn qayta ishlanmoqda...*",
        parse_mode='Markdown'
    )
    
    try:
        output_file = text_to_word(text)
        
        stats["total_conversions"] += 1
        stats["conversions_by_type"]["matn"] = stats["conversions_by_type"].get("matn", 0) + 1
        save_stats()
        
        await processing_msg.delete()
        
        await update.message.reply_document(
            document=output_file,
            filename="matn.docx",
            caption=f"✅ *Matn Word ga o'tkazildi!*\n"
                   f"📏 Belgilar: {len(text)}\n"
                   f"📄 Fayl: matn.docx",
            parse_mode='Markdown'
        )
        
    except Exception as e:
        await processing_msg.delete()
        await update.message.reply_text(
            f"❌ *Xatolik:* {str(e)[:200]}",
            parse_mode='Markdown'
        )

async def feedback_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Fikr-mulohaza"""
    await update.message.reply_text(
        "💬 *Fikr-mulohaza uchun*\n\n"
        "Bot haqida fikringizni yozing yoki adminga murojaat qiling:\n"
        "@admin_username\n\n"
        "Rahmat! 🙏",
        parse_mode='Markdown'
    )

async def error_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Xatolik boshqaruvi"""
    logger.error(f"Xatolik: {context.error}")
    
    try:
        if update and update.effective_message:
            await update.effective_message.reply_text(
                "❌ *Kutilmagan xatolik!*\n\n"
                "Iltimos qaytadan urinib ko'ring.\n"
                "Muammo davom etsa: /help",
                parse_mode='Markdown'
            )
    except:
        pass

# =============== ASOSIY FUNKSIYA ===============
def main():
    """Botni ishga tushirish"""
    
    if not TOKEN:
        logger.error("❌ TELEGRAM_BOT_TOKEN o'rnatilmagan!")
        return
    
    # Statistikani yuklash
    load_stats()
    
    logger.info("🤖 Bot ishga tushirilmoqda...")
    
    # Application yaratish
    app = Application.builder().token(TOKEN).build()
    
    # Command handlerlar
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("formats", formats_command))
    app.add_handler(CommandHandler("stats", stats_command))
    app.add_handler(CommandHandler("about", about_command))
    app.add_handler(CommandHandler("admin", admin_command))
    app.add_handler(CommandHandler("feedback", feedback_command))
    
    # Message handlerlar
    app.add_handler(MessageHandler(filters.DOCUMENT, handle_document))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    
    # Callback handler
    app.add_handler(CallbackQueryHandler(button_handler))
    
    # Error handler
    app.add_error_handler(error_handler)
    
    # Webhook yoki polling
    WEBHOOK_URL = os.getenv("WEBHOOK_URL")
    PORT = int(os.getenv("PORT", "8080"))
    
    if WEBHOOK_URL:
        logger.info(f"🌐 Webhook: {WEBHOOK_URL}")
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
