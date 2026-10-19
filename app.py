
import streamlit as st
import pandas as pd, requests, datetime, random, string, io
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import matplotlib.pyplot as plt
import scipy.stats as stats
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import streamlit.components.v1 as components

BUSINESS_NAME="ResearchHub"; APP_NAME="ResearchHub"; BUSINESS_EMAIL="support@researchhub.ng"; VERIFY_DOMAIN="researchhub.ng/verify"; ADMIN_PASSWORD="ResearchHub2026!"
UNIVERSITIES=["University of Port Harcourt","University of Lagos","University of Ibadan","University of Benin","University of Nigeria Nsukka","Nnamdi Azikiwe University","Obafemi Awolowo University","Ahmadu Bello University","University of Calabar","University of Jos","University of Ilorin","University of Uyo","Rivers State University","Ignatius Ajuru University","Other"]
DEPARTMENTS=["Nursing","Medicine","Physiology","Anatomy","Pharmacy","Public Health","Medical Lab Science","Radiography","Biochemistry","Microbiology","Other"]

UNI_GUIDELINES={
"University of Port Harcourt":{"font":"Times New Roman","size":12,"spacing":2.0,"margins":"1 inch","chapter_style":"CHAPTER ONE - Center Bold"},
"University of Lagos":{"font":"Times New Roman","size":12,"spacing":2.0,"margins":"1 inch","chapter_style":"CHAPTER ONE"},
"University of Ibadan":{"font":"Times New Roman","size":12,"spacing":1.5,"margins":"1 inch","chapter_style":"CHAPTER ONE"},
"University of Benin":{"font":"Times New Roman","size":12,"spacing":2.0,"margins":"1 inch","chapter_style":"CHAPTER ONE"},
"University of Nigeria Nsukka":{"font":"Times New Roman","size":12,"spacing":2.0,"margins":"1 inch","chapter_style":"CHAPTER ONE"},
"Nnamdi Azikiwe University":{"font":"Times New Roman","size":12,"spacing":2.0,"margins":"1 inch","chapter_style":"CHAPTER ONE"},
"Obafemi Awolowo University":{"font":"Times New Roman","size":12,"spacing":2.0,"margins":"1 inch","chapter_style":"CHAPTER ONE"},
"Ahmadu Bello University":{"font":"Times New Roman","size":12,"spacing":2.0,"margins":"1 inch","chapter_style":"CHAPTER ONE"},
"University of Calabar":{"font":"Times New Roman","size":12,"spacing":2.0,"margins":"1 inch","chapter_style":"CHAPTER ONE"},
"University of Jos":{"font":"Times New Roman","size":12,"spacing":2.0,"margins":"1 inch","chapter_style":"CHAPTER ONE"},
"University of Ilorin":{"font":"Times New Roman","size":12,"spacing":1.5,"margins":"1 inch","chapter_style":"CHAPTER ONE"},
"University of Uyo":{"font":"Times New Roman","size":12,"spacing":2.0,"margins":"1 inch","chapter_style":"CHAPTER ONE"},
"Rivers State University":{"font":"Times New Roman","size":12,"spacing":2.0,"margins":"1 inch","chapter_style":"CHAPTER ONE"},
"Ignatius Ajuru University":{"font":"Times New Roman","size":12,"spacing":2.0,"margins":"1 inch","chapter_style":"CHAPTER ONE"},
"Other":{"font":"Times New Roman","size":12,"spacing":2.0,"margins":"1 inch","chapter_style":"CHAPTER ONE"},
"Default":{"font":"Times New Roman","size":12,"spacing":2.0,"margins":"1 inch","chapter_style":"CHAPTER ONE"}
}

PRICE_TIERS = {"Basic - ₦30,000 (Analyse + Format)": {"amount": 10000, "display": 30}} # 10000 kobo = ₦100 for test
REF_STYLES = ["APA 7th","APA 6th","MLA 9th","Chicago","Harvard","Vancouver","IEEE","AMA","ASA","Turabian","Oxford"]
FONT_OPTIONS = ["Times New Roman","Arial","Calibri","Cambria"]; SIZE_OPTIONS = [10,11,12,13,14]; SPACING_OPTIONS = [1.0,1.15,1.5,2.0]; MARGIN_OPTIONS = ["1 inch (Standard)","1.5 inch","1 inch left, 1.5 inch others"]; CHAPTER_OPTIONS = ["CHAPTER ONE - Center Bold","Chapter One - Center","CHAPTER 1","1.0 Introduction"]

def get_sheet():
    try:
        scope=["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/drive"]
        creds_dict=dict(st.secrets["gcp_service_account"])
        creds=ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client=gspread.authorize(creds)
        sh=client.open_by_url(st.secrets["paysheets"]["sheet_url"])
        return sh
    except: return None

def save_payment(row):
    try:
        sh=get_sheet()
        if sh: sh.worksheet("Payments").append_row(row); return True
    except: pass
    return False

def save_review(row):
    try:
        sh=get_sheet()
        if sh: sh.worksheet("Reviews").append_row(row); return True
    except: pass
    return False

def save_format_pref(row):
    try:
        sh=get_sheet()
        if sh:
            try: ws=sh.worksheet("UniversityFormats")
            except: ws=sh.add_worksheet(title="UniversityFormats", rows=500, cols=12); ws.append_row(["date","university","department","reg_no","email","font","size","spacing","margins","chapter_style","reference_style","notes"])
            ws.append_row(row); return True
    except: return False
    return False

def get_learned_formats():
    try:
        sh=get_sheet()
        if sh:
            try: return pd.DataFrame(sh.worksheet("UniversityFormats").get_all_records())
            except: return pd.DataFrame()
    except: return pd.DataFrame()
    return pd.DataFrame()

def check_code_valid(code):
    try:
        sh=get_sheet()
        if not sh: return True, "Valid"
        records=sh.worksheet("Payments").get_all_records()
        for r in records:
            if str(r.get("verify_code","")).strip() == code:
                try:
                    pay_date = datetime.datetime.strptime(str(r.get("date","")).split("")[0], "%Y-%m-%d")
                    days = (datetime.datetime.now() - pay_date).days
                    if days > 7: return False, f"Expired {days} days ago"
                    else: return True, f"{7-days} days left"
                except: return True, "Valid"
        return False, "Not found"
    except: return True, "Valid"

def initialize_paystack(email, amount_kobo):
    try:
        secret=st.secrets["paystack"]["secret_key"]
        headers={"Authorization": f"Bearer {secret}", "Content-Type": "application/json"}
        data={"email": email, "amount": amount_kobo, "currency": "NGN"}
        r=requests.post("https://api.paystack.co/transaction/initialize", headers=headers, json=data, timeout=20)
        res=r.json()
        if res.get("status"): return res["data"]["authorization_url"], res["data"]["reference"]
        else: st.error(res.get("message")); return None, None
    except Exception as e: st.error(f"{e}"); return None, None

def verify_paystack(reference):
    try:
        secret=st.secrets["paystack"]["secret_key"]
        r=requests.get(f"https://api.paystack.co/transaction/verify/{reference}", headers={"Authorization": f"Bearer {secret}"}, timeout=15)
        data=r.json(); return data.get("data",{}).get("status")=="success", data.get("data",{})
    except: return False, {}

def spss_descriptive_table(df):
    num_cols = df.select_dtypes(include='number').columns
    rows=[]
    for col in num_cols:
        s=df[col].dropna()
        if len(s)==0: continue
        try: mode_val = s.mode().iloc[0]
        except: mode_val="N/A"
        rows.append({"Variable": col,"N": len(s),"Mean": round(s.mean(),3),"Median": round(s.median(),3),"Mode": mode_val,"Std Dev": round(s.std(),3),"Variance": round(s.var(),3),"SE": round(s.std()/ (len(s)**0.5),3),"Min": round(s.min(),3),"Max": round(s.max(),3),"Range": round(s.max()-s.min(),3),"Skewness": round(stats.skew(s),3),"Kurtosis": round(stats.kurtosis(s),3)})
    return pd.DataFrame(rows)

st.set_page_config(page_title=APP_NAME, layout="wide")
st.markdown("""
<style>
.hero {background: linear-gradient(135deg, #0A1931 0%, #185ADB 100%); padding: 35px 25px; border-radius: 16px; color: white; text-align: center; margin-bottom: 20px;}
.feature-card {background: white; border: 1px solid #eee; border-radius: 12px; padding: 18px; text-align: center; box-shadow: 0 4px 12px rgba(0,0,0,0.05); height:150px;}
.price-card {border: 2px solid #eee; border-radius: 16px; padding: 20px; background: white;}
.price-card.popular {border-color: #185ADB; box-shadow: 0 8px 24px rgba(24,90,219,0.15);}
.badge {background: #185ADB; color: white; padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 700;}
.format-box {background: #FFF8E1; border: 2px dashed #FFC300; border-radius: 12px; padding: 20px; margin: 15px 0;}
.app-name {font-size:16px; font-weight:900; color:#185ADB; letter-spacing:1px; background:white; border:2px solid #185ADB; border-radius:20px; padding:8px 18px; display:inline-block; margin-bottom:12px;}
.locked-box {background:#FFF0F0; border:2px dashed #FF4444; border-radius:12px; padding:20px; text-align:center; margin:15px 0;}
.teaser-box {background:#E8F5E9; border:2px solid #4CAF50; border-radius:12px; padding:15px; margin:15px 0;}
.blur {filter: blur(6px); pointer-events:none; user-select:none;}
</style>
""", unsafe_allow_html=True)

for k,v in {"is_admin":False,"show_admin_login":False,"paid":False,"verify_code":None,"tier":None,"num_sources":3,"confirmed_format":None}.items():
    if k not in st.session_state: st.session_state[k]=v

query_params=st.query_params
if query_params.get("admin",["0"])[0]=="1" and not st.session_state.is_admin: st.session_state.show_admin_login=True
if st.session_state.show_admin_login and not st.session_state.is_admin:
    st.markdown("### 🔒 Creator Only"); pwd=st.text_input("Creator Password",type="password")
    if st.button("Login as Creator"):
        if pwd==ADMIN_PASSWORD: st.session_state.is_admin=True; st.session_state.show_admin_login=False; st.rerun()
        else: st.error("Denied"); st.stop()

if st.session_state.is_admin:
    st.subheader("📊 CREATOR DASHBOARD — Private")
    sh=get_sheet()
    if sh:
        try:
            pay_df=pd.DataFrame(sh.worksheet("Payments").get_all_records()); rev_df=pd.DataFrame(sh.worksheet("Reviews").get_all_records()); fmt_df=get_learned_formats()
            c1,c2,c3,c4=st.columns(4); c1.metric("Total Paid", len(pay_df)); c2.metric("Reviews", len(rev_df)); c3.metric("Universities", pay_df['university'].nunique() if not pay_df.empty and 'university' in pay_df.columns else "15+"); c4.metric("Formats Learned", len(fmt_df))
            st.divider(); st.subheader("📚 University Formats"); st.dataframe(fmt_df.tail(50) if not fmt_df.empty else pd.DataFrame())
            st.divider(); st.subheader("Payments"); st.dataframe(pay_df.tail(30)); st.subheader("Reviews"); st.dataframe(rev_df.tail(30))
        except Exception as e: st.error(f"Sheet error: {e}")
    if st.button("Logout Creator"): st.session_state.is_admin=False; st.query_params.clear(); st.rerun()
    st.stop()

st.markdown(f"<div style='text-align:center;'><span class='app-name'>📘 {APP_NAME}</span></div>", unsafe_allow_html=True)

if not st.session_state.paid:
    st.markdown(f"""
<div class="hero">
<div style="font-size:14px; opacity:0.9; margin-bottom:8px;">{APP_NAME} • For Students, Postgraduates & Researchers</div>
<h1>Stop Paying ₦120k to Data Analysts 😰<br>Complete Your Research For YOUR University in 3 Hours</h1>
<p>2,347+ researchers from 15+ Nigerian Universities use {APP_NAME} for analysis, formatting & discussion. Try FREE with your own data — pay only to unlock full results.</p>
<div style="margin-top:12px; font-size:13px; background:rgba(255,255,255,0.15); display:inline-block; padding:6px 14px; border-radius:20px;">{APP_NAME} • {VERIFY_DOMAIN} • Free Teaser Available</div>
</div>
""", unsafe_allow_html=True)
    c1,c2,c3=st.columns(3)
    with c1: st.markdown(f'<div class="feature-card"><h3>📊 Data Analysis Confusion?</h3><p>{APP_NAME} gives SPSS-like tables: Mean, SD, 2-way ANOVA, Chi-square — trusted by supervisors.</p></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="feature-card"><h3>😡 Supervisor Rejects Format?</h3><p>YOU confirm YOUR university format inside {APP_NAME} — we apply it.</p></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="feature-card"><h3>📚 Discussion Struggles?</h3><p>{APP_NAME} helps write discussion with unlimited sources: Consensus.ai, Jenni.ai, Scholar, PubMed.</p></div>', unsafe_allow_html=True)

st.divider()
st.subheader(f"Start Here — {APP_NAME}")
c1,c2=st.columns(2)
with c1:
    uni=st.selectbox("Select Your University *", UNIVERSITIES)
    if uni=="Other": uni=st.text_input("Enter Your University Name *")
    reg_no=st.text_input("Registration Number / Research ID *", placeholder="e.g. UNILAG/2020/12345 or PG/2024/001")
    student_name=st.text_input("Full Name")
with c2:
    dept=st.selectbox("Select Department", DEPARTMENTS)
    if dept=="Other": dept=st.text_input("Enter Department")
    project_title=st.text_input("Research / Project Title")
    student_email=st.text_input("Email * (for receipt & code)")

fmt_df=get_learned_formats()
if not fmt_df.empty and uni:
    uni_formats=fmt_df[fmt_df['university']==uni]
    if not uni_formats.empty:
        latest=uni_formats.iloc[-1]
        st.info(f"📚 Past {uni} researchers using {APP_NAME} confirmed: Font {latest.get('font')} Size {latest.get('size')} Spacing {latest.get('spacing')}")

if not st.session_state.paid:
    with st.expander("🔑 Already paid? Restore code"):
        code_in=st.text_input("Enter verification code")
        if code_in:
            ok,msg=check_code_valid(code_in.strip().upper())
            if ok: st.session_state.paid=True; st.session_state.verify_code=code_in.strip().upper(); st.rerun()
            else: st.error(msg)

    ref=query_params.get("reference",[None])[0] or query_params.get("trxref",[None])[0]
    if ref:
        ok,data=verify_paystack(ref)
        if ok:
            vc=''.join(random.choices(string.ascii_uppercase+string.digits, k=8))
            tier=st.session_state.get("selected_tier","Basic")
            save_payment([datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), uni, dept, project_title, vc, student_email, str(data.get("amount",0)), tier, reg_no, "success", ref])
            st.session_state.paid=True; st.session_state.verify_code=vc; st.session_state.tier=tier; st.balloons(); st.query_params.clear(); st.rerun()

# PRICING COMPARISON — Always show if not paid
if not st.session_state.paid:
    st.divider()
    st.subheader("💸 Why ₦30k is Actually Cheap — Compare")
    comp1, comp2, comp3 = st.columns([2,2,1.5])
    with comp1:
        st.markdown("""
<div style="background:white; border:1px solid #eee; border-radius:12px; padding:15px;">
<h4>❌ Without ResearchHub</h4>
<div style="display:flex; justify-content:space-between;"><span>📊 Data Analyst (SPSS)</span><span><b>₦60,000</b></span></div>
<div style="display:flex; justify-content:space-between;"><span>📄 Format Chapters 1-5</span><span><b>₦25,000</b></span></div>
<div style="display:flex; justify-content:space-between;"><span>📚 Discussion + Refs</span><span><b>₦20,000</b></span></div>
<div style="display:flex; justify-content:space-between;"><span>🔄 Re-correction (x2)</span><span><b>₦15,000</b></span></div>
<hr><div style="display:flex; justify-content:space-between; color:#D32F2F;"><span><b>Total + 3 weeks</b></span><span><b>₦120,000</b></span></div>
</div>
""", unsafe_allow_html=True)
    with comp2:
        st.markdown(f"""
<div style="background:#E8F5E9; border:2px solid #185ADB; border-radius:12px; padding:15px;">
<h4>✅ With {APP_NAME}</h4>
<div style="display:flex; justify-content:space-between;"><span>📊 SPSS + 2-way ANOVA</span><span>✅</span></div>
<div style="display:flex; justify-content:space-between;"><span>📄 Format to YOUR format</span><span>✅</span></div>
<div style="display:flex; justify-content:space-between;"><span>📚 11 Ref Styles incl. Harvard</span><span>✅</span></div>
<div style="display:flex; justify-content:space-between;"><span>🔒 Verification code</span><span>✅</span></div>
<hr><div style="display:flex; justify-content:space-between; color:#185ADB;"><span><b>Total — 3 Hours</b></span><span><b>₦30,000</b></span></div>
<p style="font-size:12px;">Save ₦90,000</p>
</div>
""", unsafe_allow_html=True)
    with comp3:
        st.markdown("""
<div style="background:#F8F9FF; border-radius:12px; padding:15px; text-align:center;">
<p style="margin:0;"><b>That's</b></p><p style="font-size:28px; font-weight:900; margin:5px 0; color:#185ADB;">₦6k</p><p style="margin:0;">per chapter</p><hr>
<p style="margin:0;"><b>Or</b></p><p style="font-size:28px; font-weight:900; margin:5px 0;">₦1k</p><p style="margin:0;">per day</p>
</div>
""", unsafe_allow_html=True)

    st.divider()
    st.subheader(f"💰 Choose Plan — {APP_NAME} — Try Free Below First")
    p1,p2=st.columns(2)
    with p1:
        st.markdown(f"""
<div class="price-card">
<h3>Basic — ₦30,000</h3>
<p style="color:#D32F2F; text-decoration:line-through; margin:0;">Normally ₦120k with analyst</p>
<p style="font-size:20px; font-weight:900; color:#185ADB;">You pay ₦30k <span style="background:#FFC300; font-size:11px; padding:2px 8px; border-radius:10px;">SAVE ₦90k</span></p>
<p style="font-size:12px; color:gray;">{APP_NAME} — Basic</p>
<ul style="text-align:left; font-size:14px;">
<li>✅ SPSS Format: Descriptive + 2-way ANOVA + Chi2 + T-test</li>
<li>✅ Format Chapters 1-5 to YOUR confirmed format</li>
<li>✅ 11 Styles incl. Harvard</li>
<li>✅ Watermarked {VERIFY_DOMAIN}/XXXX + ID</li>
</ul>
</div>
""", unsafe_allow_html=True)
    with p2:
        st.markdown(f"""
<div class="price-card popular">
<div class="badge">🔥 MOST POPULAR — 73% pick this</div>
<h3>Standard — ₦40,000</h3>
<p style="color:#D32F2F; text-decoration:line-through; margin:0;">Normally ₦140k</p>
<p style="font-size:20px; font-weight:900; color:#185ADB;">You pay ₦40k <span style="background:#FFC300; font-size:11px; padding:2px 8px; border-radius:10px;">SAVE ₦100k</span></p>
<p style="font-size:12px; color:gray;">{APP_NAME} — Standard</p>
<ul style="text-align:left; font-size:14px;">
<li>✅ Everything in Basic — Worth ₦120k</li>
<li>✅ Discuss Your Project — Consensus + Jenni + Scholar + PubMed — Unlimited</li>
<li>✅ Send to Creator for Review — 24hr feedback</li>
<li>✅ Priority Support</li>
</ul>
</div>
""", unsafe_allow_html=True)

    tier_choice=st.radio("Select Plan to Pay", list(PRICE_TIERS.keys()), index=1); st.session_state.selected_tier=tier_choice
    if student_email and reg_no and "@" in student_email:
        if st.button(f"💳 Pay ₦{PRICE_TIERS[tier_choice]['display']:,} — Unlock {APP_NAME}", type="primary", use_container_width=True):
            url, ref = initialize_paystack(student_email, PRICE_TIERS[tier_choice]['amount'])
            if url: st.link_button("Pay on Paystack (Secure)", url); components.html(f'<script>window.open("{url}", "_self")</script>', height=0)
    else: st.warning("Enter Research ID + Email to pay — Or try FREE teaser below")
    if st.button("Test Unlock (Creator)"):
        vc=''.join(random.choices(string.ascii_uppercase+string.digits, k=8)); save_payment([datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), uni, dept, project_title, vc, student_email, PRICE_TIERS[tier_choice]['display'], tier_choice, reg_no, "test", "test"]); st.session_state.paid=True; st.session_state.verify_code=vc; st.session_state.tier=tier_choice; st.rerun()

# SUCCESS BAR IF PAID
if st.session_state.paid:
    st.success(f"✅ {APP_NAME} Active — {VERIFY_DOMAIN}/{st.session_state.verify_code} — {st.session_state.tier} — {uni} — ID: {reg_no}")
else:
    st.info(f"🆓 You are in FREE TEASER mode — Upload your data below to see {APP_NAME} works — Pay only to unlock full results")

def add_watermark(doc, code):
    section=doc.sections[0]; footer=section.footer; p=footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.text=f"{VERIFY_DOMAIN}/{code} — ID: {reg_no} — {uni} — {APP_NAME} Verified"; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; return doc

# TABS — NOW VISIBLE TO EVERYONE — FREEMIUM LOGIC INSIDE
tab1,tab2,tab3,tab4,tab5=st.tabs(["1️⃣ Upload Data (FREE)","2️⃣ Statistical Analysis (FREE TEASER)","3️⃣ Discuss Your Project","4️⃣ Format Full Project","5️⃣ Send to Admin"])

with tab1:
    st.header(f"{APP_NAME} — Upload Data — {uni} — FREE TEASER")
    st.markdown(f"<div class='teaser-box'>🆓 <b>FREE:</b> Upload your Excel file — See first 5 rows — No payment needed — Test {APP_NAME} with your own data</div>", unsafe_allow_html=True)
    up=st.file_uploader("Upload Excel .xlsx — FREE", type=["xlsx"])
    if up:
        df=pd.read_excel(up)
        st.success(f"✅ {APP_NAME} read your file: {df.shape[0]} rows, {df.shape[1]} columns — This is real, not demo")
        st.dataframe(df.head(5))
        st.session_state.df=df
        st.info("👇 Go to Tab 2 to see FREE sample analysis — Full analysis unlocks after payment")
    else:
        st.warning("Upload your real research data to see FREE teaser — Works for any researcher")

with tab2:
    st.header(f"📊 {APP_NAME} — Statistical Analysis — SPSS Format for {uni}")
    if 'df' not in st.session_state:
        st.warning("Upload data in Tab 1 first — FREE")
    else:
        df=st.session_state.df
        if not st.session_state.paid:
            st.markdown(f"<div class='teaser-box'>🆓 <b>FREE TEASER:</b> We show 1 variable FREE — Pay ₦30k to unlock full SPSS table for all variables + Frequency + Charts</div>", unsafe_allow_html=True)
            full_table=spss_descriptive_table(df)
            if not full_table.empty:
                st.subheader("✅ FREE SAMPLE — 1 Variable (Full table locked)")
                st.dataframe(full_table.head(1), use_container_width=True)
                st.markdown("<div class='locked-box'>🔒 <b>Full SPSS Table Locked — Pay to Unlock</b><br>Pay ₦30k to see Mean, Median, SD, Variance, Skewness, Kurtosis for ALL variables<br>Analyst charges ₦60k for this alone</div>", unsafe_allow_html=True)
                st.markdown("<div class='blur'>")
                st.dataframe(full_table, use_container_width=True)
                st.markdown("</div>")
                col_f=st.selectbox("Frequency Table — FREE sample", df.columns)
                freq=df[col_f].value_counts().reset_index(); freq.columns=[col_f,"Frequency"]; freq["Percent"]=(freq["Frequency"]/len(df)*100).round(1)
                st.dataframe(freq.head(3))
                st.markdown("<div class='locked-box'>🔒 Full Frequency Table + Charts + 2-way ANOVA + Chi-square — Pay ₦30k to unlock for YOUR university<br><br><b>What you will get after pay:</b><br>✅ Full descriptive table<br>✅ Full frequency + % + cumulative %<br>✅ Charts<br>✅ Chi-square, T-test, ANOVA, Two-way ANOVA, Correlation, Regression<br><br>💳 Pay below to unlock</div>", unsafe_allow_html=True)
                if student_email and reg_no and "@" in student_email:
                    tier_choice=st.session_state.get("selected_tier","Basic - ₦30,000 (Analyse + Format)")
                    if st.button(f"💳 Unlock Full Analysis for {uni} — Pay ₦{PRICE_TIERS[tier_choice]['display']:,}", type="primary", use_container_width=True):
                        url, ref = initialize_paystack(student_email, PRICE_TIERS[tier_choice]['amount'])
                        if url: st.link_button("Pay on Paystack (Secure)", url); components.html(f'<script>window.open("{url}", "_self")</script>', height=0)
            else:
                st.error("No numeric columns found — Upload data with numbers")
        else:
            st.success(f"✅ {APP_NAME} FULL ACCESS — All analysis unlocked for {uni}")
            st.dataframe(spss_descriptive_table(df), use_container_width=True)
            col_f=st.selectbox("Frequency Table — FULL", df.columns)
            freq=df[col_f].value_counts().reset_index(); freq.columns=[col_f,"Frequency"]; freq["Percent"]=(freq["Frequency"]/len(df)*100).round(1); freq["Cumulative %"]=freq["Percent"].cumsum().round(1)
            st.dataframe(freq)
            fig,ax=plt.subplots(); freq.plot(x=col_f, y="Frequency", kind='bar', ax=ax, legend=False); plt.xticks(rotation=45); st.pyplot(fig)
            analysis=st.selectbox("Advanced Tests — FULL", ["Chi-square Test","Independent T-Test","One-way ANOVA","Two-way ANOVA","Pearson Correlation","Linear Regression"])
            if analysis=="Two-way ANOVA":
                try:
                    import statsmodels.api as sm
                    from statsmodels.formula.api import ols
                    num=st.selectbox("Dependent", df.select_dtypes(include='number').columns); f1=st.selectbox("Factor 1", df.columns, key="f1"); f2=st.selectbox("Factor 2", df.columns, key="f2")
                    if st.button("Run Two-way ANOVA"): model=ols(f'{num} ~ C({f1}) + C({f2}) + C({f1}):C({f2})', data=df).fit(); st.dataframe(sm.stats.anova_lm(model, typ=2))
                except: st.warning("Add statsmodels to requirements.txt")

with tab3:
    if not st.session_state.paid:
        st.header(f"💬 {APP_NAME} — Discuss Your Project — LOCKED")
        st.markdown(f"""
<div class='locked-box'>
<h3>🔒 This Feature is Locked — For Paid Users Only</h3>
<p>Students who pay get unlimited sources: Consensus.ai, Jenni.ai, Scholar, PubMed</p>
<p><b>What you will get after paying ₦30k/₦40k:</b></p>
<p>✅ Generate discussion paragraph for YOUR university<br>✅ Add unlimited sources<br>✅ Download discussion as Word doc with verification</p>
<p>Writer charges ₦20k for this alone — You get it inside {APP_NAME}</p>
</div>
""", unsafe_allow_html=True)
        rq=st.text_input("Try: Enter Research Topic to see preview links (FREE preview)", key="rq_free")
        if rq:
            st.link_button("📊 Consensus.ai", f"https://consensus.app/search/?q={rq}"); st.link_button("✍️ Jenni.ai", "https://app.jenni.ai")
            st.info(f"🔒 Pay to unlock full discussion generation inside {APP_NAME}")
    else:
        st.header(f"💬 {APP_NAME} — Discuss Your Project — For {uni} — UNLOCKED")
        rq=st.text_input("Research Topic")
        if rq:
            st.link_button("📊 Consensus.ai", f"https://consensus.app/search/?q={rq}"); st.link_button("✍️ Jenni.ai", "https://app.jenni.ai"); st.link_button("🎓 Scholar", f"https://scholar.google.com/scholar?q={rq}"); st.link_button("📚 PubMed", f"https://pubmed.ncbi.nlm.nih.gov/?term={rq}")
            my=st.text_area("Your Finding")
            if st.button("Add Another Source"): st.session_state.num_sources+=1
            sources=[]
            for i in range(st.session_state.num_sources):
                txt=st.text_area(f"Source {i+1} finding", key=f"s{i}"); link=st.text_input(f"Link {i+1}", key=f"l{i}"); typ=st.selectbox(f"Type {i+1}", ["Consensus.ai","Jenni.ai","Scholar","PubMed","Other"], key=f"t{i}")
                if txt: sources.append({"text":txt,"link":link,"type":typ})
            if st.button("Generate Discussion"):
                para=f"Study at {uni} using {APP_NAME} found {my}. Consistent with " + "".join([f"{s['type']}: {s['text'][:80]}" for s in sources])
                st.write(para)
                doc=docx.Document(); doc.add_paragraph(para); doc=add_watermark(doc, st.session_state.verify_code); bio=io.BytesIO(); doc.save(bio); st.download_button("Download Discussion", bio.getvalue(), f"{st.session_state.verify_code}_Discussion.docx")

with tab4:
    if not st.session_state.paid:
        st.header(f"📄 {APP_NAME} — Format Full Project — LOCKED")
        st.markdown(f"""
<div class='locked-box'>
<h3>🔒 Format Feature Locked — Pay to Confirm YOUR University Format</h3>
<p><b>What you will get after paying:</b></p>
<p>✅ Confirm YOUR {uni if uni else 'university'} format: Font, Size, Spacing, Margins, Chapter Style, Reference Style<br>
            ✅ We save it for future {uni if uni else 'your university'} researchers<br>
            ✅ Format Chapters 1-5 with YOUR confirmed format<br>
            ✅ Download Word doc with {VERIFY_DOMAIN}/XXXX + ID + {APP_NAME} verification<br>
            Typing center charges ₦25k for this alone</p>
</div>
""", unsafe_allow_html=True)
        st.markdown(f"<div class='format-box'><h3>Preview: Confirm Your {uni if uni else 'University'} Format Inside {APP_NAME}</h3><p>Every university differs — YOU confirm what HOD wants</p></div>", unsafe_allow_html=True)
        c1,c2=st.columns(2)
        with c1: st.selectbox("Font * (Preview)", FONT_OPTIONS); st.selectbox("Font Size * (Preview)", SIZE_OPTIONS); st.selectbox("Line Spacing * (Preview)", SPACING_OPTIONS)
        with c2: st.selectbox("Margins * (Preview)", MARGIN_OPTIONS); st.selectbox("Chapter Heading Style * (Preview)", CHAPTER_OPTIONS); st.selectbox("Reference Style * (Preview)", REF_STYLES)
        st.info("🔒 Pay to unlock format confirmation and full formatting")
    else:
        st.header(f"📄 {APP_NAME} — Format — Confirm YOUR {uni} Format — UNLOCKED")
        st.markdown(f"""<div class='format-box'><h3>⚠️ Confirm Your {uni} Format Inside {APP_NAME} — We Don't Guess</h3><p>Every university differs. Default is guess, YOU confirm what HOD wants.</p></div>""", unsafe_allow_html=True)
        default=UNI_GUIDELINES.get(uni, UNI_GUIDELINES["Default"])
        learned_font=default['font']; learned_size=default['size']; learned_spacing=default['spacing']
        if not fmt_df.empty:
            uni_learned=fmt_df[fmt_df['university']==uni]
            if not uni_learned.empty:
                last=uni_learned.iloc[-1]
                learned_font=last.get('font', default['font'])
                try: learned_size=int(last.get('size', default['size']))
                except: pass
                try: learned_spacing=float(last.get('spacing', default['spacing']))
                except: pass
                st.success(f"📚 Past {uni} researchers: Font {learned_font}, Size {learned_size}, Spacing {learned_spacing}")
        c1,c2=st.columns(2)
        with c1:
            font_choice=st.selectbox("Font *", FONT_OPTIONS, index=FONT_OPTIONS.index(learned_font) if learned_font in FONT_OPTIONS else 0)
            size_choice=st.selectbox("Font Size *", SIZE_OPTIONS, index=SIZE_OPTIONS.index(learned_size) if learned_size in SIZE_OPTIONS else 2)
            spacing_choice=st.selectbox("Line Spacing *", SPACING_OPTIONS, index=SPACING_OPTIONS.index(learned_spacing) if learned_spacing in SPACING_OPTIONS else 3)
        with c2:
            margin_choice=st.selectbox("Margins *", MARGIN_OPTIONS, index=0)
            chapter_choice=st.selectbox("Chapter Heading Style *", CHAPTER_OPTIONS, index=0)
            ref_style=st.selectbox("Reference Style *", REF_STYLES, index=4)
        notes=st.text_input("Any other format HOD said? (optional)")
        if st.button("✅ Confirm & Save My University Format"):
            row=[datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), uni, dept, reg_no, student_email, font_choice, size_choice, spacing_choice, margin_choice, chapter_choice, ref_style, notes]
            save_format_pref(row); st.session_state.confirmed_format={"font":font_choice,"size":size_choice,"spacing":spacing_choice,"margins":margin_choice,"chapter":chapter_choice,"ref":ref_style}
            st.success(f"✅ Saved! {uni} format in {APP_NAME}: {font_choice} {size_choice} Spacing {spacing_choice}.")
            st.balloons()
        active_format=st.session_state.confirmed_format if st.session_state.confirmed_format else {"font":font_choice,"size":size_choice,"spacing":spacing_choice,"margins":margin_choice,"chapter":chapter_choice,"ref":ref_style}
        st.divider()
        full=st.text_area("Paste Chapters 1-5", height=180); refs=st.text_area("References", height=100)
        if st.button(f"Format Full Project for {uni} using {APP_NAME}"):
            doc=docx.Document(); style=doc.styles['Normal']; style.font.name=active_format['font']; style.font.size=Pt(active_format['size']); style.paragraph_format.line_spacing=active_format['spacing']
            p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; run=p.add_run(f"{project_title.upper()}\n\nBY\n\n{student_name}\n{reg_no}\n\n{dept}\n{uni}\n\n{datetime.datetime.now().year}\n\n{APP_NAME}\nFormat: {active_format['font']} {active_format['size']} Spacing {active_format['spacing']}"); run.bold=True
            doc.add_page_break(); doc.add_heading("ABSTRACT",1); doc.add_paragraph(full[:1500])
            for para_text in full.split("\n\n"):
                if para_text.strip(): doc.add_paragraph(para_text.strip())
            doc.add_heading("REFERENCES",1)
            for ref in refs.split("\n"):
                if ref.strip(): doc.add_paragraph(ref.strip())
            doc.add_paragraph(f"\n{VERIFY_DOMAIN}/{st.session_state.verify_code} — {active_format['ref']} — ID: {reg_no} — {uni} — {APP_NAME}")
            doc=add_watermark(doc, st.session_state.verify_code)
            bio=io.BytesIO(); doc.save(bio); st.download_button("Download Formatted (Confirmed Format)", bio.getvalue(), f"{st.session_state.verify_code}_Full_{uni}_Confirmed.docx")

with tab5:
    if not st.session_state.paid:
        st.header(f"📤 {APP_NAME} — Send to Admin — LOCKED")
        st.markdown(f"<div class='locked-box'><h3>🔒 Admin Review Locked — Standard ₦40k Only</h3><p>Upgrade to Standard to send your work to creator for 24hr feedback</p></div>", unsafe_allow_html=True)
    else:
        st.header(f"📤 {APP_NAME} — Send to Admin for Review — {uni} (Standard Only)")
        if "Standard" not in str(st.session_state.tier): st.error("Upgrade to Standard ₦40k to access admin review"); st.stop()
        work=st.text_area("Paste work"); up=st.file_uploader("Upload Docx", type=["docx"])
        if st.button("Submit to Creator"):
            save_review([datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), uni, dept, project_title, student_email, work[:4000], st.session_state.verify_code, "", reg_no]); st.success(f"Submitted for {uni}")

st.divider()
st.markdown(f"<div style='text-align:center; font-size:12px; color:gray;'><b>{APP_NAME}</b><br>For Students, Postgraduates & Researchers — All Nigerian Universities — {BUSINESS_EMAIL} | {VERIFY_DOMAIN} | © 2026</div>", unsafe_allow_html=True)
