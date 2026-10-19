import streamlit as st
import pandas as pd, requests, datetime, random, string, io, re
from collections import Counter
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import matplotlib.pyplot as plt
import scipy.stats as stats
import gspread
from oauth2client.service_account import ServiceAccountCredentials

BUSINESS_NAME="ResearchHub"; BUSINESS_EMAIL="support@researchhub.ng"; VERIFY_DOMAIN="researchhub.ng/verify"; ADMIN_PASSWORD="ResearchHub2026!"
UNIVERSITIES=["University of Port Harcourt","University of Lagos","University of Ibadan","University of Benin","University of Nigeria Nsukka","Nnamdi Azikiwe University","Obafemi Awolowo University","Ahmadu Bello University","University of Calabar","University of Jos","University of Ilorin","University of Uyo","Rivers State University","Ignatius Ajuru University","Other"]
DEPARTMENTS=["Nursing","Medicine","Physiology","Anatomy","Pharmacy","Public Health","Medical Lab Science","Radiography","Biochemistry","Microbiology","Other"]
UNI_GUIDELINES={"University of Port Harcourt":{"font":"Times New Roman","size":12,"spacing":2.0},"Default":{"font":"Times New Roman","size":12,"spacing":2.0}}

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
        if sh: sh.worksheet("Payments").append_row(row)
    except: pass

def save_review(row):
    try:
        sh=get_sheet()
        if sh: sh.worksheet("Reviews").append_row(row)
    except: pass

def verify_paystack(reference):
    try:
        secret=st.secrets["paystack"]["secret_key"]
        r=requests.get(f"https://api.paystack.co/transaction/verify/{reference}", headers={"Authorization": f"Bearer {secret}"}, timeout=15)
        return r.json().get("data",{}).get("status")=="success"
    except: return False

st.set_page_config(page_title=BUSINESS_NAME, layout="wide")
st.markdown(f'<div style="font-size:26px;font-weight:800;color:#0A1931;">{BUSINESS_NAME} — Student Tool 🇳🇬</div>', unsafe_allow_html=True)
st.markdown(f'<div style="font-size:12px;color:gray;">{BUSINESS_EMAIL} | {VERIFY_DOMAIN} | Works with Consensus.ai + Jenni.ai</div>', unsafe_allow_html=True)

query_params=st.query_params
if 'is_admin' not in st.session_state: st.session_state.is_admin=False
if 'show_admin_login' not in st.session_state: st.session_state.show_admin_login=False
if 'paid' not in st.session_state: st.session_state.paid=False
if 'qual_df' not in st.session_state: st.session_state.qual_df=pd.DataFrame()

if query_params.get("admin",["0"])[0]=="1" and not st.session_state.is_admin: st.session_state.show_admin_login=True
if st.session_state.show_admin_login and not st.session_state.is_admin:
    st.info("Admin Private"); pwd=st.text_input("Password",type="password")
    if st.button("Login"):
        if pwd==ADMIN_PASSWORD: st.session_state.is_admin=True; st.session_state.show_admin_login=False; st.rerun()
        else: st.error("Wrong"); st.stop()

if st.session_state.is_admin:
    st.subheader("📊 ADMIN — ResearchHub Dashboard — Google Sheet Live")
    sh=get_sheet()
    if sh:
        try:
            pay=sh.worksheet("Payments").get_all_records(); rev=sh.worksheet("Reviews").get_all_records()
            pay_df=pd.DataFrame(pay); rev_df=pd.DataFrame(rev)
            st.metric("Total Paid", len(pay_df))
            if not pay_df.empty: st.dataframe(pay_df.tail(30)); st.download_button("Download Payments CSV", pay_df.to_csv(index=False), "payments.csv")
            st.divider(); st.subheader("📥 Submissions for Review")
            if not rev_df.empty:
                st.dataframe(rev_df.tail(30))
                for i,row in rev_df.iterrows():
                    if not str(row.get("feedback","")).strip():
                        st.markdown(f"**{row['verify']} | {row['title']} | {row['email']}**")
                        st.text_area(f"Work {row['verify']}", str(row['work'])[:1000], key=f"work_{i}", height=80)
                        fb=st.text_input(f"Feedback for {row['verify']}", key=f"fb_{i}")
                        if st.button(f"Save Feedback {row['verify']}", key=f"save_{i}"):
                            sh.worksheet("Reviews").update_cell(i+2, 8, fb)
                            st.success("Saved to Sheet — student will see in Tab 5")
            else: st.write("No reviews yet")
        except Exception as e: st.error(f"Sheet error: {e}")
    if st.button("Logout Admin"): st.session_state.is_admin=False; st.query_params.clear(); st.rerun()
    st.stop()

c1,c2=st.columns(2)
with c1:
    uni=st.selectbox("Select Your University", UNIVERSITIES)
    if uni=="Other": uni=st.text_input("Enter University")
with c2:
    dept=st.selectbox("Select Department", DEPARTMENTS)
    if dept=="Other": dept=st.text_input("Enter Department")
project_title=st.text_input("Project Title"); student_email=st.text_input("Your Email for Receipt & Feedback")

if not st.session_state.paid:
    st.divider()
    ref=query_params.get("reference",[None])[0]
    if ref:
        if verify_paystack(ref):
            verify_code=''.join(random.choices(string.ascii_uppercase+string.digits, k=8))
            row=[datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), uni, dept, project_title, verify_code, student_email, "30000", "success", ref]
            save_payment(row); st.session_state.paid=True; st.session_state.verify_code=verify_code
            st.success(f"✅ Payment Verified! Code: {verify_code}"); st.query_params.clear(); st.rerun()
        else: st.error("Payment not verified — contact support@researchhub.ng")
    try:
        public_key=st.secrets["paystack"]["public_key"]
        st.subheader("Pay ₦30,000 to Unlock ResearchHub")
        st.markdown(f"""
<script src="https://js.paystack.co/v1/inline.js"></script>
<script>
        function payWithPaystack(){{
          var handler = PaystackPop.setup({{
            key: '{public_key}',
            email: '{student_email or "student@test.com"}',
            amount: 3000000,
            currency: "NGN",
            ref: 'RESEARCHHUB_'+Math.floor((Math.random()*1000000000)+1),
            callback: function(response){{ window.location.href = window.location.href.split('?')[0] + "?reference=" + response.reference; }},
            onClose: function(){{ alert('Cancelled'); }}
          }}); handler.openIframe();
        }}
</script>
<button onclick="payWithPaystack()" style="background:#0A1931;color:white;padding:14px 28px;border:none;border-radius:8px;font-weight:800;font-size:16px;">💳 Pay ₦30k with Paystack</button>
""", unsafe_allow_html=True)
    except:
        st.warning("Add Paystack keys in Secrets — Test Mode")
        if st.button("Pay Test Mode (Temporary)"):
            verify_code=''.join(random.choices(string.ascii_uppercase+string.digits, k=8))
            save_payment([datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), uni, dept, project_title, verify_code, student_email, "30000", "test"])
            st.session_state.paid=True; st.session_state.verify_code=verify_code; st.rerun()
    st.stop()
else:
    verify_code=st.session_state.verify_code
    st.success(f"✅ Active — {VERIFY_DOMAIN}/{verify_code}")

tab1,tab2,tab3,tab4,tab5=st.tabs(["1️⃣ Upload Data","2️⃣ SPSS + Qualitative","3️⃣ Discuss with Jenni + Consensus","4️⃣ Format Full Project","5️⃣ Send to Admin for Review"])

with tab1:
    st.header("Upload Data"); up=st.file_uploader("Excel .xlsx", type=["xlsx"])
    if up: df=pd.read_excel(up); st.dataframe(df.head()); st.session_state.df=df

with tab2:
    st.header("📊 SPSS + Qualitative"); st.subheader("PART A: Quantitative")
    if 'df' not in st.session_state: st.warning("Upload in Tab1")
    else:
        df=st.session_state.df
        analysis=st.selectbox("SPSS Test", ["Descriptive (Mean, SD, Frequency)","Chi-square Test","Independent T-Test","One-way ANOVA","Pearson Correlation","Linear Regression"])
        if analysis=="Descriptive (Mean, SD, Frequency)":
            st.dataframe(df.describe()); col_f=st.selectbox("Frequency", df.columns); st.dataframe(df[col_f].value_counts().reset_index()); fig,ax=plt.subplots(); df[col_f].value_counts().plot(kind='bar',ax=ax); st.pyplot(fig)
        elif analysis=="Chi-square Test":
            c1=st.selectbox("Row", df.columns, key="chi1"); c2=st.selectbox("Col", df.columns, key="chi2")
            if st.button("Run Chi-square"): cont=pd.crosstab(df[c1],df[c2]); chi2,p,dof,exp=stats.chi2_contingency(cont); st.dataframe(cont); st.write(f"Chi2={chi2:.3f}, p={p:.4f}")
        elif analysis=="Independent T-Test":
            num=st.selectbox("Numeric", df.select_dtypes(include='number').columns); grp=st.selectbox("Group", df.columns)
            if st.button("Run T-Test"): groups=df[grp].unique()[:2]; g1=df[df[grp]==groups[0]][num].dropna(); g2=df[df[grp]==groups[1]][num].dropna(); t,p=stats.ttest_ind(g1,g2); st.write(f"t={t:.3f}, p={p:.4f}")
        elif analysis=="One-way ANOVA":
            num=st.selectbox("Num ANOVA", df.select_dtypes(include='number').columns, key="anova_num"); grp=st.selectbox("Group ANOVA", df.columns, key="anova_group")
            if st.button("Run ANOVA"): groups=[df[df[grp]==g][num].dropna() for g in df[grp].unique()]; f,p=stats.f_oneway(*groups); st.write(f"F={f:.3f}, p={p:.4f}")
        elif analysis=="Pearson Correlation":
            nums=df.select_dtypes(include='number').columns.tolist(); x=st.selectbox("X", nums, key="cx"); y=st.selectbox("Y", nums, key="cy")
            if st.button("Run Correlation"): r,p=stats.pearsonr(df[x].dropna(), df[y].dropna()); st.write(f"r={r:.3f}, p={p:.4f}")
        elif analysis=="Linear Regression":
            nums=df.select_dtypes(include='number').columns.tolist(); x=st.selectbox("Ind X", nums, key="regx"); y=st.selectbox("Dep Y", nums, key="regy")
            if st.button("Run Regression"): slope,intercept,r,p,err=stats.linregress(df[x].dropna(), df[y].dropna()); st.write(f"Y={intercept:.2f}+{slope:.2f}X, R={r:.3f}, p={p:.4f}")
    st.divider(); st.subheader("PART B: Qualitative")
    qual_opt=st.selectbox("Method", ["Thematic Analysis","Content Analysis","Word Frequency"])
    qual_text=st.text_area("Paste transcripts", height=150)
    if qual_text:
        if qual_opt=="Word Frequency":
            words=re.findall(r'\w+', qual_text.lower()); stop=['the','and','is','to','of','a','in','that','it','for','was','were','are']
            words=[w for w in words if w not in stop and len(w)>3]; freq=Counter(words).most_common(20); df_freq=pd.DataFrame(freq, columns=["Word","Frequency"]); st.dataframe(df_freq); st.session_state.qual_df=df_freq; fig,ax=plt.subplots(); df_freq.plot(x="Word",y="Frequency",kind='bar',ax=ax,legend=False); plt.xticks(rotation=45); st.pyplot(fig)
        elif qual_opt=="Thematic Analysis":
            themes={"Perception":["perceive","think","believe"],"Challenge":["challenge","difficult","problem","stress"],"Support":["support","help","mentor"],"Knowledge":["know","knowledge","learn"]}; found=[]
            for theme,kws in themes.items():
                cnt=sum(qual_text.lower().count(k) for k in kws)
                if cnt>0: st.write(f"- {theme}: {cnt}"); found.append([theme,cnt])
            if found: st.session_state.qual_df=pd.DataFrame(found, columns=["Theme","Count"])
        else: lines=[l for l in qual_text.split("\n") if l.strip()]; st.dataframe(pd.DataFrame({"Response":lines[:20]})); st.session_state.qual_df=pd.DataFrame({"Response":lines})

with tab3:
    st.header("💬 Discuss with Jenni.ai + Consensus.ai")
    rq=st.text_input("Research Topic / Question", placeholder="E.g. Prevalence of stress among nursing students in Nigeria")
    if rq:
        consensus_url=f"https://consensus.app/search/?q={rq.replace('', '%20')}"
        scholar_url=f"https://scholar.google.com/scholar?q={rq.replace('', '+')}"
        pubmed_url=f"https://pubmed.ncbi.nlm.nih.gov/?term={rq.replace('', '+')}"
        st.subheader("🔍 Step1: Find Relevant Articles — Click")
        colA,colB=st.columns(2)
        with colA: st.link_button("📊 Search Consensus.ai", consensus_url); st.link_button("📚 Search PubMed", pubmed_url)
        with colB: st.link_button("✍️ Open Jenni.ai", "https://app.jenni.ai"); st.link_button("🎓 Google Scholar", scholar_url)
        st.divider(); st.subheader("Step2: Paste what you found")
        my=st.text_area("Your Own Finding (Tab2)", height=80)
        c1,c2=st.columns(2)
        with c1: s1=st.text_area("Article1 from Consensus.ai", height=100); l1=st.text_input("Link1")
        with c2: s2=st.text_area("Article2 from Jenni.ai", height=100); l2=st.text_input("Link2")
        s3=st.text_area("Article3 optional", height=60)
        dtype=st.selectbox("Your Finding vs Articles", ["Similar / Agrees","Contradicts / Different","Mixed"])
        if st.button("Generate Discussion"):
            if not my or not s1: st.error("Enter finding + 1 article")
            else:
                if dtype=="Similar / Agrees": para=f"The present study found that {my}. This agrees with {s1[:120]} as retrieved from Consensus.ai. Similarly {s2 if s2 else 'other studies via Jenni.ai reported comparable outcomes'}. Similarity due to similar population in {uni}."
                elif dtype=="Contradicts / Different": para=f"The present study found that {my}. This contradicts {s1[:120]} found on Consensus.ai. Also {s2 if s2 else ''} via Jenni.ai had variance. Difference due to setting in {uni}."
                else: para=f"The present study found that {my}. Partially consistent with {s1[:120]}. Mixed literature from Jenni.ai suggests contextual factors."
                st.write(para)
                doc=docx.Document(); doc.styles['Normal'].font.name="Times New Roman"; doc.styles['Normal'].font.size=Pt(12); doc.styles['Normal'].paragraph_format.line_spacing=2.0
                doc.add_heading("Discussion - Consensus + Jenni",1); doc.add_paragraph(para); doc.add_heading("Articles",2)
                if s1: doc.add_paragraph(s1 + (f" Link:{l1}" if l1 else "")); 
                if s2: doc.add_paragraph(s2 + (f" Link:{l2}" if l2 else ""))
                if s3: doc.add_paragraph(s3)
                doc.add_paragraph(f"\n{VERIFY_DOMAIN}/{verify_code}")
                bio=io.BytesIO(); doc.save(bio); st.download_button("Download Discussion", bio.getvalue(), file_name=f"{verify_code}_Discussion.docx")

with tab4:
    st.header("📄 Format Entire Project")
    guideline=UNI_GUIDELINES.get(uni, UNI_GUIDELINES["Default"])
    st.info(f"{uni}: {guideline['font']} {guideline['size']} Spacing {guideline['spacing']}")
    full=st.text_area("Paste Chapters 1-5", height=180); refs=st.text_area("References", height=90)
    style_sel=st.selectbox("Citation", ["APA 7th","APA 6th","MLA","Vancouver"])
    if st.button("Format Full Project"):
        doc=docx.Document(); style=doc.styles['Normal']; style.font.name=guideline['font']; style.font.size=Pt(guideline['size']); style.paragraph_format.line_spacing=guideline['spacing']
        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; run=p.add_run(f"{project_title.upper()}\n\nBY\n\n{dept}\n{uni}\n\n{datetime.datetime.now().year}"); run.bold=True; doc.add_page_break()
        doc.add_heading("ABSTRACT",1).alignment=WD_ALIGN_PARAGRAPH.CENTER; doc.add_paragraph(full[:1200])
        for para_text in full.split("\n\n"):
            if para_text.strip(): doc.add_paragraph(para_text.strip())
        doc.add_heading("REFERENCES",1)
        for ref in refs.split("\n"):
            if ref.strip(): pr=doc.add_paragraph(ref.strip()); pr.paragraph_format.left_indent=Pt(36); pr.paragraph_format.first_line_indent=Pt(-36)
        doc.add_paragraph(f"\n{VERIFY_DOMAIN}/{verify_code} — {style_sel}")
        bio=io.BytesIO(); doc.save(bio); st.download_button("Download Formatted", bio.getvalue(), file_name=f"{verify_code}_Formatted.docx")

with tab5:
    st.header("📤 Send Work to Admin for Review & Get Feedback")
    st.info("Submit work — saved permanently to Google Sheet")
    work_text=st.text_area("Paste your work", height=180)
    up_review=st.file_uploader("Or Upload Docx", type=["docx"], key="rev_up")
    if st.button("Submit for Review"):
        if not work_text and not up_review: st.error("Paste work")
        elif not student_email: st.error("Enter email in top")
        else:
            content=work_text + (f" [File:{up_review.name}]" if up_review else "")
            row=[datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), uni, dept, project_title, student_email, content[:3000], verify_code, ""]
            save_review(row); st.success(f"✅ Submitted! Code {verify_code} — saved to Google Sheet")
    st.divider(); st.subheader("🔔 Check Feedback")
    check=st.text_input("Enter verification code")
    if check:
        sh=get_sheet()
        if sh:
            try:
                rev_df=pd.DataFrame(sh.worksheet("Reviews").get_all_records())
                found=rev_df[rev_df['verify']==check]
                if found.empty: st.warning("No submission for this code in Sheet")
                else:
                    for _,r in found.iterrows():
                        if str(r.get("feedback","")).strip(): st.success(f"Feedback from Admin: {r['feedback']}")
                        else: st.info("Submitted — Admin not yet feedback, check later")
            except Exception as e: st.error(f"Sheet read error: {e}")

st.divider()
st.markdown(f"<div style='text-align:center; font-size:11px; color:gray;'><b>{BUSINESS_NAME}</b><br>{BUSINESS_EMAIL} | {VERIFY_DOMAIN}<br>Works with Consensus.ai + Jenni.ai | Paystack + Google Sheet Live | © 2026</div>", unsafe_allow_html=True)
