import streamlit as st
import pandas as pd
import plotly.express as px
import io

# 1. Page Configuration (App ရဲ့ မျက်နှာပြင် Layout ပြင်ဆင်ခြင်း)
st.set_page_config(page_title="Smart Audit Automation Tool", page_icon="🛡️", layout="wide")

st.title("🛡️ Smart Audit Automation Dashboard")
st.markdown("---")

# 2. Sidebar ဖိုင်တင်ရန်နေရာ
st.sidebar.header("📁 Data Upload")
uploaded_file = st.sidebar.file_uploader("Excel သို့မဟုတ် CSV ဖိုင်ကို တင်ပါ", type=["xlsx", "csv"])

# 3. အဓိက လုပ်ဆောင်ချက်များ
if uploaded_file is not None:
    try:
        # File Type အလိုက် ဖတ်ခြင်း
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        st.sidebar.success("✅ ဖိုင်တင်ခြင်း အောင်မြင်သည်!")
        
        # Data Preview ပြသခြင်း
        st.subheader("📋 Raw Data Preview")
        st.dataframe(df.head(10), use_container_width=True)
        st.caption(u"စုစုပေါင်း စာရင်းအရေအတွက်: {:,} ခု | Columns အရေအတွက်: {} ခု".format(len(df), len(df.columns)))
        
        st.markdown("---")
        
        # Column Name များကို User ရွေးချယ်ခိုင်းခြင်း (Error ကာကွယ်ရန်)
        st.subheader("⚙️ Audit Parameters Setting")
        col1, col2 = st.columns(2)
        with col1:
            amount_col = st.selectbox("ငွေပမာဏ (Amount) ရှိသော Column ကို ရွေးပါ", df.columns)
            threshold_amount = st.number_input("စစ်ဆေးလိုသည့် အမြင့်ဆုံး ငွေပမာဏ သတ်မှတ်ရန် (Threshold)", value=10000)
        with col2:
            status_col = st.selectbox("ငွေပေးချေမှုစနစ် (သို့) အခြေအနေ ပြသော Column ကို ရွေးပါ (Optional)", [None] + list(df.columns))
            status_value = st.text_input("ရှာဖွေလိုသည့် အခြေအနေ (ဥပမာ - Credit, Failed, Pending)", value="Credit")

        st.markdown("---")
        
        # --- AUDIT RULES EXECUTION ---
        st.subheader("🔍 Audit Exception Reports (စစ်ဆေးတွေ့ရှိချက်များ)")
        
        # Rule 1: High Value Transactions
        df[amount_col] = pd.to_numeric(df[amount_col], errors='coerce') # ဂဏန်းပုံစံပြောင်းခြင်း
        high_value_df = df[df[amount_col] >= threshold_amount]
        
        # Rule 2: Specific Status (e.g., Credit / Suspicious)
        if status_col:
            status_df = df[df[status_col].astype(str).str.lower() == status_value.lower()]
        else:
            status_df = pd.DataFrame()
            
        # Rule 3: Missing Values (စာရင်းကျန်နေသော ကွက်လပ်များ)
        missing_data_df = df[df.isnull().any(axis=1)]
        
        # တဘ်များခွဲ၍ ပြသခြင်း
        tab1, tab2, tab3 = st.tabs([f"💰 High Value (>= {threshold_amount})", f"⚠️ Status: {status_value}", "❌ Missing Fields"])
        
        with tab1:
            st.warning(f"ငွေပမာဏ မြင့်မားသော စာရင်းပေါင်း {len(high_value_df)} ခု တွေ့ရှိရသည်။")
            st.dataframe(high_value_df, use_container_width=True)
            
        with tab2:
            if status_col:
                st.info(f"သတ်မှတ်ထားသော အခြေအနေနှင့် ကိုက်ညီသည့် စာရင်းပေါင်း {len(status_df)} ခု တွေ့ရှိရသည်။")
                st.dataframe(status_df, use_container_width=True)
            else:
                st.write("စစ်ဆေးရန် Column မရွေးချယ်ထားပါ။")
                
        with tab3:
            st.error(u"အချက်အလက် လိုအပ်နေသော (Blank/Null ဖြစ်နေသော) စာရင်းပေါင်း {} ခု တွေ့ရှိရသည်။".format(len(missing_data_df)))
            st.dataframe(missing_data_df, use_container_width=True)
            
        st.markdown("---")
        
        # 4. Data Visualization (အနှစ်ချုပ် Chart ပြသခြင်း)
        st.subheader("📊 Visual Analytics")
        if status_col:
            fig = px.box(df, x=status_col, y=amount_col, title="Transaction Amount Distribution by Status")
            st.plotly_chart(fig, use_container_width=True)
            
        # 5. Export Report (Excel ဖိုင်တစ်ခုတည်းဖြင့် Sheet အလိုက် ဒေါင်းလုဒ်ဆွဲရန်)
        st.subheader("📥 Export Audit Report")
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            high_value_df.to_excel(writer, sheet_name='High_Value_Alerts', index=False)
            if not status_df.empty:
                status_df.to_excel(writer, sheet_name='Status_Alerts', index=False)
            missing_data_df.to_excel(writer, sheet_name='Missing_Data_Alerts', index=False)
            
        st.download_button(
            label="🚀 Audit Exceptions Report (Excel) ကို ဒေါင်းလုဒ်လုပ်ရန်",
            data=buffer.getvalue(),
            file_name="Audit_Final_Exceptions_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Error တက်သွားပါသဖြင့် ဖိုင် သို့မဟုတ် ကော်လံ ရွေးချယ်မှု ပြန်စစ်ပေးပါ။ အသေးစိတ်: {e}")

else:
    # ဖိုင်မတင်ရသေးခင် ပြသမည့် မျက်နှာပြင်
    st.info("💡 ဘယ်ဘက် Sidebar မှတစ်ဆင့် Audit စစ်ဆေးမည့် Excel သို့မဟုတ် CSV ဖိုင်ကို စတင် တင်သွင်း (Upload) ပေးပါ။")
