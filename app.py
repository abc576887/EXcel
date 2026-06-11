import streamlit as st
import pandas as pd
import plotly.express as px
import io

# 1. Page Configuration
st.set_page_config(page_title="Professional Audit Engine", page_icon="⚖️", layout="wide")

st.title("⚖️ Advanced Ledger & Audit Automation Engine")
st.markdown("Power Query သဘောတရားအတိုင်း Format မတူသော Sheet များကို ညှိဖတ်နိုင်ပြီး **Debit/Credit Logic** များကို စစ်ဆေးကာ **App ပေါ်တွင် ပြင်ဆင်ပြီးမှ Report ပြန်ထုတ်နိုင်သော** စနစ် ဖြစ်ပါသည်။")
st.markdown("---")

# 2. Sidebar Upload
st.sidebar.header("📁 Data Upload Center")
uploaded_file = st.sidebar.file_uploader("စစ်ဆေးမည့် Excel Ledger ဖိုင်ကို တင်ပါ", type=["xlsx"])

if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
        
        st.sidebar.success(f"📂 ဖိုင်ဖတ်ခြင်းအောင်မြင်သည်။ Sheet ({len(sheet_names)}) ခု ပါဝင်သည်။")
        
        # --- POWER QUERY STYLE DYNAMIC MAPPING ---
        st.subheader("🛠️ Step 1: Power Query Style Column Alignment")
        st.info("💡 Sheet အသီးသီးရှိ ကော်လံခေါင်းစဉ်များ မတူညီပါက စံသတ်မှတ်ချက်အတိုင်း အောက်တွင် ညှိပေးပါ။")
        
        sheet_mappings = {}
        df_list = []
        
        # Sheet တစ်ခုချင်းစီအတွက် Mapping ယူခြင်း
        for sheet in sheet_names:
            preview_df = pd.read_excel(uploaded_file, sheet_name=sheet, nrows=2)
            cols = list(preview_df.columns)
            
            with st.expander(f"📄 Sheet Name: {sheet} - Columns Configuration"):
                col_m1, col_m2, col_m3 = st.columns(3)
                with col_m1:
                    m_vch = st.selectbox("Voucher/ID Column", cols, key=f"vch_{sheet}")
                    m_acc = st.selectbox("Account Name Column", cols, key=f"acc_{sheet}")
                with col_m2:
                    m_deb = st.selectbox("Debit (ဝင်ငွေ/ကြွေးမြီ) Column", cols, key=f"deb_{sheet}")
                with col_m3:
                    m_cre = st.selectbox("Credit (ထွက်ငွေ/ရရန်) Column", cols, key=f"cre_{sheet}")
                
                sheet_mappings[sheet] = {
                    'vch': m_vch, 'acc': m_acc, 'deb': m_deb, 'cre': m_cre
                }
        
        # ဒေတာများကို ပေါင်းစပ်ပြီး Standard Format ပြောင်းလဲခြင်း
        for sheet in sheet_names:
            raw_df = pd.read_excel(uploaded_file, sheet_name=sheet)
            if not raw_df.empty:
                maps = sheet_mappings[sheet]
                
                processed_df = pd.DataFrame()
                processed_df['Voucher_No'] = raw_df[maps['vch']].astype(str).str.strip()
                processed_df['Account_Name'] = raw_df[maps['acc']].astype(str).str.strip()
                processed_df['Debit'] = pd.to_numeric(raw_df[maps['deb']], errors='coerce').fillna(0)
                processed_df['Credit'] = pd.to_numeric(raw_df[maps['cre']], errors='coerce').fillna(0)
                processed_df['Sheet_Source'] = sheet
                
                # မူရင်း ကော်လံများကိုပါ တစ်ခါတည်း သိမ်းထားမည်
                for col in raw_df.columns:
                    if col not in maps.values():
                        processed_df[f"Orig_{col}"] = raw_df[col]
                        
                df_list.append(processed_df)
                
        # Main Combined DataFrame
        combined_df = pd.concat(df_list, ignore_index=True)
        st.success("✅ Power Query ပုံစံအတိုင်း ဒေတာအားလုံးကို စံနှုန်းတစ်ခုတည်းအဖြစ် ပေါင်းစပ်ပြီးပါပြီ။")
        
        st.markdown("---")
        
        # --- STEP 2: LIVE DATA EDITING (IN-APP DATA EDITOR) ---
        st.subheader("📝 Step 2: Live Ledger Data Review & Editing")
        st.markdown("💡 အောက်ပါ ဇယားကွက်ထဲတွင် မှားယွင်းနေသော စာရင်းများကို **တိုက်ရိုက်ကလစ်နှိပ်၍ ပြင်ဆင်နိုင်ပါသည်။** ပြင်ဆင်ပြီးပါက အောက်က Audit Result တွင် ချက်ချင်း ပြောင်းလဲသွားမည် ဖြစ်သည်။")
        
        # Streamlit Data Editor ကိုသုံးပြီး အပြန်အလှန် ပြင်ဆင်နိုင်အောင် လုပ်ခြင်း
        edited_df = st.data_editor(combined_df, num_rows="dynamic", use_container_width=True, key="ledger_editor")
        
        st.markdown("---")
        
        # --- STEP 3: ADVANCED AUDIT LOGIC RULES ---
        st.subheader("🔍 Step 3: Debit / Credit Audit Exceptions Report")
        
        # Threshold Settings
        high_val_limit = st.number_input("🚨 သံသယဖြစ်ဖွယ် အဝိုင်းလိုက် ငွေပမာဏကြီးများ သတ်မှတ်ရန် (High-Value Threshold)", value=50000)
        
        # 1. Unbalanced / Invalid Transactions Rule
        # Logic: Debit ကော Credit ကော တူညီစွာ ရှိနေခြင်း (သို့) နှစ်ခုလုံး သုည ဖြစ်နေခြင်း
        unbalanced_df = edited_df[
            ((edited_df['Debit'] == edited_df['Credit']) & (edited_df['Debit'] != 0)) | 
            ((edited_df['Debit'] == 0) & (edited_df['Credit'] == 0))
        ]
        
        # 2. Duplicate Vouchers Rule
        # Logic: Voucher နံပါတ် တူညီပြီး စာရင်းနှစ်ခါထပ်နေခြင်း
        duplicate_vch_df = edited_df[edited_df.duplicated(subset=['Voucher_No'], keep=False) & (edited_df['Voucher_No'] != 'nan')]
        
        # 3. High Value & Round Sum Anomaly
        # Logic: သတ်မှတ်ပမာဏထက် ကြီးမားသော စာရင်းများ
        high_value_df = edited_df[(edited_df['Debit'] >= high_val_limit) | (edited_df['Credit'] >= high_val_limit)]
        
        # 4. Missing Accounts
        missing_acc_df = edited_df[(edited_df['Account_Name'].isna()) | (edited_df['Account_Name'] == 'nan') | (edited_df['Account_Name'] == '')]
        
        # Displaying Results in Tabs
        t1, t2, t3, t4 = st.tabs([
            f"❌ Unbalanced / Invalid Entries ({len(unbalanced_df)})", 
            f"🆔 Duplicate Vouchers ({len(duplicate_vch_df)})", 
            f"🚨 High Value >= {high_val_limit:,} ({len(high_value_df)})",
            f"❓ Missing Accounts ({len(missing_acc_df)})"
        ])
        
        with t1:
            st.warning("Debit နှင့် Credit တစ်ပြိုင်နက်တည်း ရှိနေသော (သို့) နှစ်ခုလုံး သုညဖြစ်နေသော လွဲမှားသည့် စာရင်းများ")
            st.dataframe(unbalanced_df, use_container_width=True)
        with t2:
            st.warning("Voucher နံပါတ် ထပ်နေသဖြင့် အမှား သို့မဟုတ် Double Entry ဖြစ်နိုင်ခြေရှိသော စာရင်းများ")
            st.dataframe(duplicate_vch_df, use_container_width=True)
        with t3:
            st.error(f"ငွေပမာဏ ကြီးမားလွန်းသဖြင့် အထူးစစ်ဆေးရန် လိုအပ်သော စာရင်းများ")
            st.dataframe(high_value_df, use_container_width=True)
        with t4:
            st.info("စာရင်းခေါင်းစဉ် (Account Name) ကျန်ရစ်ခဲ့သော အချက်အလက်များ")
            st.dataframe(missing_acc_df, use_container_width=True)
            
        st.markdown("---")
        
        # --- STEP 4: VISUAL ANALYTICS ---
        st.subheader("📊 Step 4: Executive Visual Analytics")
        c1, c2 = st.columns(2)
        with c1:
            fig_deb = px.box(edited_df, x='Sheet_Source', y='Debit', title='Debit Entries Distribution by Sheet', color='Sheet_Source')
            st.plotly_chart(fig_deb, use_container_width=True)
        with c2:
            fig_cre = px.box(edited_df, x='Sheet_Source', y='Credit', title='Credit Entries Distribution by Sheet', color='Sheet_Source')
            st.plotly_chart(fig_cre, use_container_width=True)
            
        st.markdown("---")
        
        # --- STEP 5: EXPORT RE-DESIGNED CLEAN DATA ---
        st.subheader("📥 Step 5: Export Final Cleaned & Audited Ledger")
        st.markdown("အထက်တွင် **သင်ကိုယ်တိုင် ပြင်ဆင်ပြီးသား (Finalized Clean Data)** နှင့် Audit တွေ့ရှိချက်များကို Multi-sheet Excel ဖိုင်တစ်ခုတည်းအဖြစ် ပြန်လည် ထုတ်ယူနိုင်ပါသည်။")
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            # တည်းဖြတ်ပြီးသား Ledger တစ်ခုလုံးကို Sheet ၁ အနေနဲ့ ထုတ်မယ်
            edited_df.to_excel(writer, sheet_name='Final_Cleaned_Ledger', index=False)
            # အမှားတွေ့ရှိချက်များကို သီးသန့် Sheet များခွဲထုတ်မယ်
            if not unbalanced_df.empty:
                unbalanced_df.to_excel(writer, sheet_name='Audit_Unbalanced', index=False)
            if not duplicate_vch_df.empty:
                duplicate_vch_df.to_excel(writer, sheet_name='Audit_Duplicates', index=False)
            if not high_value_df.empty:
                high_value_df.to_excel(writer, sheet_name='Audit_High_Values', index=False)
                
        st.download_button(
            label="🚀 ပြင်ဆင်ပြီးသား Final Audit Report (Excel) ကို ဒေါင်းလုဒ်လုပ်ရန်",
            data=buffer.getvalue(),
            file_name="Finalized_Audited_Ledger_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Error တက်သွားပါသည်။ ကော်လံရွေးချယ်မှု ပြန်စစ်ပေးပါ။ အသေးစိတ်: {e}")
else:
    st.info("💡 စတင်ရန် ဘယ်ဘက် Sidebar မှတစ်ဆင့် သင့်ရဲ့ General Ledger Excel ဖိုင်ကို တင်သွင်း (Upload) ပေးပါ။")
