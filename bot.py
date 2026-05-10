import telebot
import pandas as pd
import os
import re
from flask import Flask
from threading import Thread

# SIZNING BOT TOKENINGIZ
TOKEN = '8758590017:AAG6G0smLcTkJSiTeX8SmOGL7Jwtbo7aqsI'

# FAYL YUKLAY OLADIGAN ASOSIY ADMIN ID'SI (O'zingizning ID ni yozing)
ADMIN_ID = 7845695013  

# QIDIRUVDAN FOYDALANISHGA RUXSAT BERILGANLAR RO'YXATI (O'zingiz va dilerlar ID lari)
# ID larni vergul bilan ajratib yozasiz
ALLOWED_USERS = [7845695013, 300148507, 271145081, 8269421944, 604944814, 7628098805] 

bot = telebot.TeleBot(TOKEN)
DB_FILE = 'narxlar_bazasi.csv'

def load_db():
    if os.path.exists(DB_FILE):
        return pd.read_csv(DB_FILE)
    return pd.DataFrame()

db = {
    'baza': load_db()
}
admin_data = {}

def extract_core_model(name):
    name = str(name).lower()
    name = re.sub(r'\(.*?\)', '', name)
    name = re.sub(r'\b\d+([.,]\d+)?\s*mm\b', '', name)
    name = re.sub(r'\b(2\.8|3\.6|4|6|8|12|2\.8-12|2\.8~12)\b', '', name)
    
    words_to_remove = ['hikvision', 'mikrotik', 'dahua', 'switch', 'router', 'poe', 
                       'камера', 'уличная', 'купольная', 'ip', 'видеокамера', 'цилиндрическая']
    for w in words_to_remove:
        name = name.replace(w, '')
        
    name = re.sub(r'[^a-z0-9]', '', name)
    
    if name.startswith('dsids'): name = name[5:]
    elif name.startswith('ids'): name = name[3:]
    elif name.startswith('ds'): name = name[2:]
    
    if name.startswith('2cd'): name = name[3:]
    elif name.startswith('2ce'): name = name[3:]
    elif name.startswith('2de'): name = name[3:]
    elif name.startswith('2td'): name = name[3:]
    elif name.startswith('3e'): name = name[2:]
    
    return name

def extract_data_from_file(filepath):
    xls = pd.ExcelFile(filepath)
    all_data = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        col_nom = next((c for c in df.columns if str(c).strip().lower() in ['наименование', 'модель', 'model', 'tovar nomi']), None)
        col_narx = next((c for c in df.columns if 'цена' in str(c).lower() or 'price' in str(c).lower()), None)
        
        if col_nom and col_narx:
            df_clean = df.dropna(subset=[col_nom])
            df_clean = df_clean[[col_nom, col_narx]].copy()
            df_clean.columns = ['Tovar_Nomi', 'Narx']
            df_clean['Nom_Qidiruv'] = df_clean['Tovar_Nomi'].apply(extract_core_model)
            df_clean['Narx'] = pd.to_numeric(df_clean['Narx'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
            df_clean = df_clean[df_clean['Narx'] > 0]
            all_data.append(df_clean)
    xls.close()
    
    if all_data:
        result_df = pd.concat(all_data, ignore_index=True)
        result_df = result_df.groupby('Nom_Qidiruv', as_index=False).agg({'Tovar_Nomi': 'first', 'Narx': 'min'})
        return result_df
    return pd.DataFrame()

@bot.message_handler(commands=['start'])
def send_welcome(message):
    # 1. Begonalarni tekshirish (RO'YXATDA YO'QLAR UCHUN)
    if message.chat.id not in ALLOWED_USERS:
        bot.reply_to(message, "⛔️ Kechirasiz, bu bot yopiq tizimda ishlaydi. Sizga ruxsat berilmagan.")
        return

    # 2. O'zimiznikilar uchun
    if message.chat.id == ADMIN_ID:
        bot.reply_to(message, "👑 Salom, Admin!\nPrayslarni yangilash va bazaga yuklash uchun /update_prices buyrug'ini bosing.")
    else:
        bot.reply_to(message, "Salom, Hamkor! Men narx qidiruvchi aqlli botman. 🔎\n\nMenga kerakli tovar modelini yozing (masalan: 1043) va men narxlarni qidirib beraman!")

@bot.message_handler(commands=['update_prices'])
def update_prices_cmd(message):
    if message.chat.id != ADMIN_ID:
        bot.reply_to(message, "❌ Kechirasiz, bu buyruq faqat Asosiy Admin uchun.")
        return
    admin_data[message.chat.id] = {}
    msg = bot.reply_to(message, "Iltimos, **NEG praysini** (1-fayl) tashlang:")
    bot.register_next_step_handler(msg, process_admin_neg)

def process_admin_neg(message):
    try:
        chat_id = message.chat.id
        if not message.document:
            msg = bot.reply_to(message, "Fayl kutilmoqda. Iltimos NEG Excel faylini yuboring:")
            bot.register_next_step_handler(msg, process_admin_neg)
            return

        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        file1_name = f"neg_admin_{chat_id}.xlsx"
        with open(file1_name, 'wb') as new_file:
            new_file.write(downloaded_file)
            
        admin_data[chat_id]['neg'] = file1_name
        msg = bot.reply_to(message, "NEG praysi saqlandi ✅\n\nEndi **KSS praysini** (2-fayl) tashlang:")
        bot.register_next_step_handler(msg, process_admin_kss)
    except Exception as e:
        bot.reply_to(message, f"Xato: {e}")

def process_admin_kss(message):
    try:
        chat_id = message.chat.id
        if not message.document:
            msg = bot.reply_to(message, "Fayl kutilmoqda. KSS Excel faylini yuboring:")
            bot.register_next_step_handler(msg, process_admin_kss)
            return

        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        file2_name = f"kss_admin_{chat_id}.xlsx"
        with open(file2_name, 'wb') as new_file:
            new_file.write(downloaded_file)
            
        bot.reply_to(message, "Baza yig'ilmoqda va aqlli tozalash ishlayapti... Kuting ⏳")
        
        file1 = admin_data[chat_id]['neg']
        file2 = file2_name
        
        df1 = extract_data_from_file(file1)
        df2 = extract_data_from_file(file2)
        
        if df1.empty and df2.empty:
            bot.send_message(chat_id, "Ikkala fayl ham bo'sh. Baza yangilanmadi.")
            return

        df_natija = pd.merge(df1, df2, on='Nom_Qidiruv', how='outer', suffixes=('_NEG', '_KSS'))
        df_natija['Narx_NEG'] = df_natija['Narx_NEG'].fillna(0)
        df_natija['Narx_KSS'] = df_natija['Narx_KSS'].fillna(0)
        df_natija['Tovar_Nomi'] = df_natija['Tovar_Nomi_NEG'].fillna(df_natija['Tovar_Nomi_KSS'])
        
        yakuniy_baza = df_natija[['Nom_Qidiruv', 'Tovar_Nomi', 'Narx_NEG', 'Narx_KSS']]
        db['baza'] = yakuniy_baza
        yakuniy_baza.to_csv(DB_FILE, index=False)
        
        os.remove(file1)
        os.remove(file2)
        
        bot.send_message(chat_id, f"🎉 Baza muvaffaqiyatli yangilandi!\nJami xotiraga olingan tovarlar: {len(df_natija)} ta.")
    except Exception as e:
        bot.reply_to(message, f"Bazaga yozishda xatolik: {e}")

@bot.message_handler(content_types=['text'])
def search_product(message):
    # Begonalarni tekshirish (RO'YXATDA YO'QLAR UCHUN)
    if message.chat.id not in ALLOWED_USERS:
        bot.reply_to(message, "⛔️ Kechirasiz, bu bot yopiq tizimda ishlaydi. Sizga ruxsat berilmagan.")
        return

    query = str(message.text).strip()
    if query.startswith('/'): return
    df = db['baza']
    if df.empty:
        bot.reply_to(message, "❌ Hozircha bazada tovarlar yo'q. Iltimos, admin prayslarni yuklashini kuting.")
        return
        
    clean_query = extract_core_model(query)
    if len(clean_query) < 2:
        bot.reply_to(message, "Iltimos, model nomini aniqroq yozing (kamida 3 ta harf/raqam).")
        return
        
    results = df[df['Nom_Qidiruv'].astype(str).str.contains(clean_query, na=False)]
    if results.empty:
        bot.reply_to(message, "😔 Kechirasiz, bunday tovar hech qaysi do'konda topilmadi.")
        return
        
    response = f"🔍 Natijalar ({len(results)} ta topildi):\n\n"
    for index, row in results.head(10).iterrows():
        n_neg = float(row['Narx_NEG'])
        n_kss = float(row['Narx_KSS'])
        tovar_nomi = row['Tovar_Nomi']
        
        if n_neg > 0 and n_kss > 0:
            str_neg = f"{n_neg} $"
            str_kss = f"{n_kss} $"
            if n_neg < n_kss:
                farq = n_kss - n_neg
                foiz = (farq / n_kss) * 100
                xulosa = f"NEG arzon 🟢 (Farq: {round(farq, 2)} $, {round(foiz, 1)}%)"
            elif n_kss < n_neg:
                farq = n_neg - n_kss
                foiz = (farq / n_neg) * 100
                xulosa = f"KSS arzon 🔵 (Farq: {round(farq, 2)} $, {round(foiz, 1)}%)"
            else:
                xulosa = "Narxlar bir xil ⚖️"
        elif n_neg > 0 and n_kss == 0:
            str_neg = f"{n_neg} $"
            str_kss = "Sotuvda yo'q ❌"
            xulosa = "Faqat NEG do'konida bor 🟢"
        elif n_kss > 0 and n_neg == 0:
            str_neg = "Sotuvda yo'q ❌"
            str_kss = f"{n_kss} $"
            xulosa = "Faqat KSS do'konida bor 🔵"
        else:
            continue
            
        response += f"📦 **{tovar_nomi}**\n▫️ NEG narxi: {str_neg}\n▫️ KSS narxi: {str_kss}\n✅ **Xulosa:** {xulosa}\n➖➖➖➖➖➖➖➖\n"
        
    if len(results) > 10:
        response += f"\n...va yana {len(results) - 10} ta tovar. Aniqroq qidirish uchun modelni to'liqroq yozing."
    bot.send_message(message.chat.id, response, parse_mode='Markdown')

# ======= 24/7 SERVER UCHUN FLASK QISMI =======
app = Flask(__name__)

@app.route('/')
def home():
    return "Bot 24/7 ishlamoqda! Himoyalangan rejim faol."

def run():
    app.run(host="0.0.0.0", port=int(os.environ.get('PORT', 8080)))

def keep_alive():
    t = Thread(target=run)
    t.start()

if __name__ == '__main__':
    keep_alive() 
    print("Xavfsiz va Miyali bot 24/7 server rejimida ishga tushdi...")
    bot.infinity_polling()
