import streamlit as st
import pandas as pd
import sqlite3
import io

DB_NAME = "group_companies_pnl.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS unified_pnl (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_name TEXT,
            business_type TEXT,
            fiscal_year INTEGER,
            sheet_name TEXT,
            account_code TEXT,
            master_category TEXT,
            particulars TEXT,
            amount REAL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

init_db()

st.set_page_config(page_title="Executive P&L Standardizer", layout="wide")
st.title("📊 P&L Standardizer (Custom Tailored for Indobest Format)")

# Sidebar Configuration
st.sidebar.header("📝 File Configuration")
company_name = st.sidebar.text_input("ကုမ္ပဏီ နာမည်", value="Myanmar Indobest (Shan Gyi)")
business_type = st.sidebar.selectbox("Business Type", ["Trading/Distribution", "Production", "Retail", "Telecom"])
fiscal_year = st.sidebar.number_input("Fiscal Year", min_value=2020, max_value=2030, value=2026)
uploaded_file = st.sidebar.file_uploader("P&L Excel/CSV ဖိုင် တင်ပါ", type=["csv", "xlsx"])

if uploaded_file:
    # CSV သို့မဟုတ် Excel ဖြစ်မဖြစ် စစ်ဆေးဖတ်ရှုခြင်း
    if uploaded_file.name.endswith('.csv'):
        # အပေါ်က အလွတ်တွေ ကျော်ဖတ်ဖို့ အဆင်ပြေအောင် အရင်စစ်မယ်
        raw_df = pd.read_csv(uploaded_file, skiprows=4)
    else:
        raw_df = pd.read_excel(uploaded_file, skiprows=4)
        
    # Column နာမည်များ သန့်စင်ခြင်း
    raw_df.columns = [str(c).strip() for c in raw_df.columns]
    
    # ဖိုင်ထဲက အဓိက လိုချင်တဲ့ Column တွေကို ဆွဲထုတ်ခြင်း
    # အလွတ်တွေ ကြားညှပ်နေတာကြောင့် 'Code', 'Description' နဲ့ ပထမဆုံး တွေ့တဲ့ ဂဏန်း Column ကို ယူမယ်
    if 'Code' in raw_df.columns and 'Description' in raw_df.columns:
        
        # နာမည်မရှိတဲ့ ကော်လံတွေကို ဖယ်ပြီး Amount ကော်လံကို ရှာခြင်း
        amount_col = None
        for col in raw_df.columns:
            if col not in ['Code', 'Description', ''] and not col.startswith('Unnamed'):
                amount_col = col
                break
        
        if not amount_col:
            # တကယ်လို့ ရှာမတွေ့ရင် တတိယမြောက် ကော်လံကို ယူမယ်
            amount_col = raw_df.columns[3]
            
        st.info(f"📍 Reading Data: `Code`, `Description` နှင့် Amount အဖြစ် `[{amount_col}]` ကော်လံကို အသုံးပြုထားသည်။")
        
        processed_rows = []
        
        for idx, row in raw_df.iterrows():
            code = str(row['Code']).strip()
            desc = str(row['Description']).strip()
            amount_val = row[amount_col]
            
            # အလွတ်လိုင်းများနှင့် စုစုပေါင်းလိုင်းများအား ဖယ်ထုတ်ခြင်း
            if pd.isna(amount_val) or desc == "" or "nan" in desc.lower(): continue
            if "total" in desc.lower() or "net profit" in desc.lower() or "စုစုပေါင်း" in desc: continue
            
            # --- 💡 DOUBLE COUNTING TRAP မှ ကာကွယ်ခြင်း Logic ---
            # ကုဒ်နံပါတ်က '000' နဲ့ ဆုံးနေရင် ဒါက Main Category (Subtotal) ဖော်ပြချက် ဖြစ်လို့ 
            # ဒေတာတွေ ထပ်မပေါင်းမိအောင် Main Category ကို ကျော်ပြီး တကယ့် Sub-items (Child rows) ကိုပဲ သိမ်းပါမယ်။
            if code.endswith('000'): 
                continue # Main Group ကို ကျော်ပြီး အသေးစိတ်စာရင်းကိုပဲ ယူမယ်
                
            # ဂဏန်းဟုတ်မဟုတ် စစ်ဆေးခြင်း
            try:
                clean_amount = float(amount_val)
            except ValueError:
                continue
                
            # --- 💡 ACCOUNT CODE အလိုက် MASTER CATEGORY အလိုအလျောက် သတ်မှတ်ခြင်း ---
            master_cat = "Other"
            if code.startswith('5'):
                master_cat = "Revenue"
            elif code.startswith('6'):
                master_cat = "COGS"
            elif code.startswith('7'):
                master_cat = "OPEX"
            elif code.startswith('9'):
                master_cat = "Other_Income_Expense"
                
            processed_rows.append({
                'company_name': company_name,
                'business_type': business_type,
                'fiscal_year': fiscal_year,
                'sheet_name': "Trading",
                'account_code': code,
                'master_category': master_cat,
                'particulars': desc,
                'amount': clean_amount
            })
            
        if processed_rows:
            final_df = pd.DataFrame(processed_rows)
            
            st.write("### 🔍 ပုံစံညှိပြီးသား ဒေတာများ Preview (စစ်ဆေးရန်)")
            # User က မျက်စိနဲ့စစ်ပြီး လိုအပ်က Category ပြင်နိုင်ရန် st.data_editor သုံးထားသည်
            edited_df = st.data_editor(final_df, use_container_width=True)
            
            if st.button("💾 ဒေတာများကို Database ထဲသို့ အပြီးသတ်သိမ်းဆည်းမည်", type="primary"):
                conn = sqlite3.connect(DB_NAME)
                edited_df.to_sql('unified_pnl', conn, if_exists='append', index=False)
                conn.close()
                st.balloons()
                st.success(f"🎉 {company_name} ၏ ဒေတာများကို ပုံစံတူ Ledger အဖြစ် SQLite DB ထဲသို့ အောင်မြင်စွာ သွင်းပြီးပါပြီ။")
        else:
            st.error("❌ ဖိုင်ထဲမှ သင့်တော်သော ဒေတာလိုင်းများကို ရှာမတွေ့ပါ။ Format ကို ပြန်လည်စစ်ဆေးပါ။")
    else:
        st.error("❌ Excel ဖိုင်ထဲတွင် 'Code' နှင့် 'Description' ကော်လံခေါင်းစဉ်များ ရှာမတွေ့ပါ။ Header Row နေရာ မှားယွင်းနေနိုင်ပါသည်။")