from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, filters, ContextTypes
import pandas as pd
from io import BytesIO

def split_dataframe_by_empty_row(df):
    groups = []
    current = []
    for _, row in df.iterrows():
        if row.isnull().all():
            if current:
                groups.append(pd.DataFrame(current, columns=df.columns))
                current = []
        else:
            current.append(row)
    if current:
        groups.append(pd.DataFrame(current, columns=df.columns))
    return groups

# Fungsi konversi Excel ke VCF
def convert_excel_to_vcf_multiple(file_bytes, filename):
    xls = pd.ExcelFile(file_bytes)
    vcf_files = []
    for sheet_name in xls.sheet_names:
        df = xls.parse(sheet_name)
        tables = split_dataframe_by_empty_row(df)
        for idx, table in enumerate(tables, 1):
            vcf_data = ''
            for _, row in table.iterrows():
                name = str(row.get('Nama') or row.get('Name') or '')
                phone = str(row.get('Telepon') or row.get('Phone') or row.get('HP') or '')
                if name and phone:
                    vcf_data += f"BEGIN:VCARD\nVERSION:3.0\nFN:{name}\nTEL;TYPE=CELL:{phone}\nEND:VCARD\n"
            if vcf_data:
                # Penamaan file: jika hanya 1 tabel, cukup nama file excel saja
                if len(tables) == 1:
                    vcf_filename = f"{filename}.vcf"
                else:
                    vcf_filename = f"{filename}_tabel{idx}.vcf"
                vcf_files.append((vcf_filename, vcf_data))
    return vcf_files
# Fungsi penanganan file
async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file = await update.message.document.get_file()
    file_bytes = BytesIO()
    await file.download_to_memory(out=file_bytes)
    file_bytes.seek(0)

    base_filename = update.message.document.file_name.rsplit('.', 1)[0]
    vcf_files = convert_excel_to_vcf_multiple(file_bytes, base_filename)

    if not vcf_files:
        await update.message.reply_text("Tidak ada kontak yang ditemukan.")
        return

    for vcf_filename, vcf_data in vcf_files:
        vcf_buffer = BytesIO(vcf_data.encode('utf-8'))
        vcf_buffer.seek(0)
        await update.message.reply_document(document=vcf_buffer, filename=vcf_filename)

# Jalankan bot
if __name__ == '__main__':
    app = ApplicationBuilder().token('7700640296:AAE8vXLyjMOrpP8Ms_58G4v71PhiltUaUts').build()
    app.add_handler(MessageHandler(filters.Document.ALL, handle_file))
    app.run_polling()
