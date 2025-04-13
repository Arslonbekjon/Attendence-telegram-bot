import os
import io
import asyncio
import pandas as pd
import matplotlib.pyplot as plt
from dotenv import load_dotenv
from telegram import Update, InputFile, InlineKeyboardMarkup, InlineKeyboardButton
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

# .env dan tokenni yuklash
load_dotenv()
TOKEN = os.getenv("BOT_TOKEN")

# /start komandasi
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Assalomu alaykum! Excel (.xlsx) formatdagi faylingizni yuboring, men sizga tahlil qilib, hisobot va grafik joʻnataman."
    )

# Excel faylni qabul qilish va tahlil qilish
async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    # Fayl nomining oxiri .xlsx ekanligini tekshiramiz
    if not document.file_name.endswith(".xlsx"):
        await update.message.reply_text("❌ Iltimos, faqat .xlsx formatdagi fayl yuboring.")
        return

    # Faylni yuklab olish
    file = await update.message.document.get_file()
    file_bytes = await file.download_as_bytearray()
    bio = io.BytesIO(file_bytes)

    try:
        df = pd.read_excel(bio)
    except Exception as e:
        await update.message.reply_text(f"❌ Faylni o'qishda xatolik: {e}")
        return

    # Kerakli ustunlar mavjudligini tekshirish
    required_columns = {"Employee", "Start", "End"}
    if not required_columns.issubset(df.columns):
        await update.message.reply_text(
            "❌ Faylda quyidagi ustunlar bo'lishi kerak: Employee, Start, End"
        )
        return

    # Vaqt ustunlarini datetime formatiga o'tkazish
    try:
        df["Start"] = pd.to_datetime(df["Start"])
        df["End"] = pd.to_datetime(df["End"])
    except Exception as e:
        await update.message.reply_text(f"❌ Vaqt formatida xatolik: {e}")
        return

    # Ish vaqtini (duration) hisoblash (daqiqalarda)
    df["Duration"] = (df["End"] - df["Start"]).dt.total_seconds() / 60
    df["Duration"] = df["Duration"].round(2)

    # Statistikani tayyorlash
    stats = df.groupby("Employee")["Duration"].sum().reset_index()
    stats["Duration"] = stats["Duration"].round(2)
    stats = stats.sort_values(by="Duration", ascending=False)

    # Hisobot matni
    summary_text = "*📊 Ish vaqti hisobot:*\n\n"
    for _, row in stats.iterrows():
        summary_text += f"👤 {row['Employee']}: {row['Duration']} daqiqa\n"

    # Excel faylga yozish
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sessions")
        stats.to_excel(writer, index=False, sheet_name="Summary")
    output.seek(0)

    await update.message.reply_text(summary_text, parse_mode="Markdown")
    try:
        await update.message.reply_document(
            InputFile(output, filename="hisobot.xlsx")
        )
        # Qisqa pauza: bot resurslarini bo'shatish uchun
        await asyncio.sleep(1)
    except Exception as e:
        await update.message.reply_text(f"❌ Excel fayl yuborilmadi: {e}")

    # Grafik (Top 10 xodim bo'yicha ish vaqti)
    top_stats = stats.head(10)
    fig, ax = plt.subplots(figsize=(8, 6))
    ax.bar(top_stats["Employee"], top_stats["Duration"], color="skyblue")
    ax.set_title("Top 10 xodim - Ish vaqti")
    ax.set_ylabel("Daqiqa")
    plt.xticks(rotation=45)
    fig.tight_layout()

    photo_buf = io.BytesIO()
    plt.savefig(photo_buf, format="png")
    photo_buf.seek(0)
    plt.close()

    try:
        await update.message.reply_photo(
            photo=InputFile(photo_buf, filename="grafik.png")
        )
    except Exception as e:
        await update.message.reply_text(f"❌ Grafik yuborilmadi: {e}")

# Botni ishga tushurish
if __name__ == "__main__":
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    # Fayl extension uchun to'g'ri filter: filters.Document.FileExtension("xlsx")
    app.add_handler(MessageHandler(filters.Document.FileExtension("xlsx"), handle_file))
    print("🤖 Bot ishlayapti...")
    app.run_polling()
