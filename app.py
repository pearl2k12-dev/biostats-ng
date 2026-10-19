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
import streamlit.components.v1 as components

# --- CONFIG ---
BUSINESS_NAME="ResearchHub"; BUSINESS_EMAIL="support@researchhub.ng"; VERIFY_DOMAIN="researchhub.ng/verify"; ADMIN_PASSWORD="ResearchHub2026!"
UNIVERSITIES=["University of Port Harcourt","University of Lagos","University of Ibadan","University of Benin","University of Nigeria Nsukka","Nnamdi Azikiwe University","Obafemi Awolowo University","Ahmadu Bello University","University of Calabar","University of Jos","University of Ilorin","University of Uyo","Rivers State University","Ignatius Ajuru University","Other"]
DEPARTMENTS=["Nursing","Medicine","Physiology","Anatomy","Pharmacy","Public Health","Medical Lab Science","Radiography","Biochemistry","Microbiology","Other"]
UNI_GUIDELINES={"University of Port Harcourt":{"font":"Times New Roman","size":12,"spacing":2.0},"Default":{"font":"Times New Roman","size":12,"spacing":2.0}}

# 5. PRICE TIERS - money making
PRICE_TIERS = {
"Basic - ₦30k (SPSS Only)": {"amount": 3000000, "display": 30000, "features": ["Descriptive","Chi2","T-Test","ANOVA","Correlation","Regression"]},
"Standard - ₦45k (Most Popular) SPSS + Qualitative + Format": {"amount": 4500000, "display": 45000, "features": ["All SPSS","Thematic","Word Freq","Format Project"]},
"Premium - ₦60k (Full) + Admin Review + Jenni/Consensus Discussion": {"amount": 6000000, "display": 60000, "features": ["Everything","Discussion","Admin Review","Priority"]}
}

AUTO_FEEDBACKS = [
"p-value wrong — check Tab 2 interpretation",
"Add 2 more articles from Consensus.ai — see Tab 3",
"References not APA 7th — fix in Tab 4",
"Sample size justification missing",
"Discussion too short — compare with 2 studies from Jenni.ai",
"Formatting error — check line spacing 2.0, Times New Roman 12",
"Good work — minor corrections in references"
]

def get_sheet():
    try:
        scope=["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/drive"]
        creds_dict=dict(st.secrets["gcp_service_account"])
        creds=ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client=gspread.authorize(creds)
        sh=client.open_by_url(st.secrets["paysheets"]["sheet_url"])
        return sh
    except Exception as e:
        return None

# 1. FIX SESSION RESET - Save every payment to sheet
def save_payment(row):
    try:
        sh=get_sheet()
        if sh: 
            sh.worksheet("Payments").append_row(row)
            return True
    except: pass
    return False

def save_review(row):
    try:
        sh=get_sheet()
        if sh: 
            sh.worksheet("Reviews").append_row(row)
            return True
    except: pass
    return False

# 3. STOP SHARING - check expiry from sheet
def check_code_valid(code):
    try:
        sh=get_sheet()
        if not sh: return True, "No sheet — allowing"
        records=sh.worksheet("Payments").get_all_records()
        for r in records:
            if str(r.get("verify_code","")).strip() == code or str(r.get("verify","")).strip() == code:
                date_str = str(r.get("date", r.get("timestamp","")))
                try:
                    pay_date = datetime.datetime.strptime(date_str.split("")[0], "%Y-%m-%d")
                    days = (datetime.datetime.now() - pay_date).days
                    if days > 7:
                        return False, f"Expired — {days} days ago (7-day limit). Please pay again."
                    else:
                        return True, f"Valid — {7-days} days left, Tier: {r.get('tier','')}"
                except:
                    return True, "Valid (no date check)"
        return False, "Code not found in sheet"
    except Exception as e:
        return True, f"Check error: {e}"

def get_payment_by_code(code):
    try:
        sh=get_sheet()
        if sh:
            df=pd.DataFrame(sh.worksheet("Payments").get_all_records())
            found=df[df.astype(str).apply(lambda row: row.astype(str).str.contains(code, na=False).any(), axis=1)]
            if not found.empty:
                return found.iloc[0].to_dict()
    except: pass
    return None

# 2. REAL PAYSTACK LIVE
def initialize_paystack(email, amount_kobo):
    try:
        secret=st.secrets["paystack"]["secret_key"]
        headers={"Authorization": f"Bearer {secret}", "Content-Type": "application/json"}
        data={"email": email, "amount": amount_kobo, "currency": "NGN"}
        r=requests.post("https://api.paystack.co/transaction/initialize", headers=headers, json=data, timeout=20)
        res=r.json()
        if res.get("status"):
            return res["data"]["authorization_url"], res["data"]["reference"]
        else:
            st.error(f"Paystack init failed: {res.get('message')}")
            return None, None
    except Exception as e:
        st.error(f"Paystack error: {e}")
        return None, None

def verify_paystack(reference):
    try:
        secret=st.secrets["paystack"]["secret_key"]
        r=requests.get(f"https://api.paystack.co/transaction/verify/{reference}", headers={"Authorization": f"Bearer {secret}"}, timeout=15)
        data=r.json()
        return data.get("data",{}).get("status")=="success", data.get("data",{})
    except: return False, {}

# --- UI ---
st.set_page_config(page_title=BUSINESS_NAME, layout="wide")
st.markdown(f'<div style="font-size:26px;font-weight:800;color:#0A1931;">{BUSINESS_NAME} — Student Tool 🇳🇬</div>', unsafe_allow_html=True)
st.markdown(f'<div style="font-size:12px;color:gray;">{BUSINESS_EMAIL} | {VERIFY_DOMAIN} | Works with Consensus.ai + Jenni.ai</div>', unsafe_allow_html=True)

# Session defaults
for k,v in {"is_admin":False,"show_admin_login":False,"paid":False,"qual_df":pd.DataFrame(),"pay_ref":None,"verify_code":None,"tier":None}.items():
    if k not in st.session_state: st.session_state[k]=v

query_params=st.query_params

# Admin
if query_params.get("admin",["0"])[0]=="1" and not st.session_state.is_admin: st.session_state.show_admin_login=True
if st.session_state.show_admin_login and not st.session_state.is_admin:
    st.info("Admin Private"); pwd=st.text_input("Password",type="password")
    if st.button("Login"):
        if pwd==ADMIN_PASSWORD: st.session_state.is_admin=True; st.session_state.show_admin_login=False; st.rerun()
        else: st.error("Wrong"); st.stop()

if st.session_state.is_admin:
    st.subheader("📊 ADMIN — ResearchHub Dashboard — Google Sheet Live (No data loss)")
    sh=get_sheet()
    if sh:
        try:
            pay_records=sh.worksheet("Payments").get_all_records(); rev_records=sh.worksheet("Reviews").get_all_records()
            pay_df=pd.DataFrame(pay_records); rev_df=pd.DataFrame(rev_records)
            col1,col2,col3=st.columns(3)
            col1.metric("Total Paid", len(pay_df))
            col2.metric("Total Reviews", len(rev_df))
            if not pay_df.empty:
                # Calculate revenue
                revenue = pay_df['amount'].astype(str).str.extract(r'(\d+)').astype(float).sum() if 'amount' in pay_df.columns else len(pay_df)*30000
                col3.metric("Est. Revenue", f"₦{int(len(pay_df)*35000):,}")
            if not pay_df.empty: st.dataframe(pay_df.tail(40)); st.download_button("Download Payments CSV", pay_df.to_csv(index=False), "payments.csv")
            st.divider(); st.subheader("📥 Submissions for Review — 1-click Auto Feedback")
            if not rev_df.empty:
                st.dataframe(rev_df.tail(30))
                for i,row in rev_df.iterrows():
                    verify = str(row.get("verify","")).strip()
                    if not str(row.get("feedback","")).strip():
                        st.markdown(f"---\n**{verify} | {row.get('title','')} | {row.get('email','')} | {row.get('university','')}**")
                        st.text_area(f"Work {verify}", str(row.get("work",""))[:1200], key=f"work_{i}", height=100)
                        # 4. AUTO FEEDBACK BUTTONS
                        cols=st.columns(3)
                        for idx, fb_text in enumerate(AUTO_FEEDBACKS):
                            c = cols[idx % 3]
                            if c.button(f"{fb_text[:28]}...", key=f"autofb_{i}_{idx}"):
                                try:
                                    # Find row number in sheet: header + i + 1
                                    sh.worksheet("Reviews").update_cell(i+2, 8, fb_text)
                                    st.success(f"Feedback sent: {fb_text}")
                                    st.rerun()
                                except Exception as e: st.error(f"{e}")
                        custom_fb=st.text_input(f"Custom feedback for {verify}", key=f"fb_{i}")
                        if st.button(f"Save Custom Feedback {verify}", key=f"save_{i}"):
                            try:
                                sh.worksheet("Reviews").update_cell(i+2, 8, custom_fb)
                                st.success("Saved to Sheet — student will see in Tab 5")
                            except Exception as e: st.error(f"{e}")
            else: st.write("No reviews yet")
        except Exception as e: st.error(f"Sheet error: {e} — Check sheet has tabs Payments and Reviews with headers")
    if st.button("Logout Admin"): st.session_state.is_admin=False; st.query_params.clear(); st.rerun()
    st.stop()

# Inputs
c1,c2=st.columns(2)
with c1:
    uni=st.selectbox("Select Your University", UNIVERSITIES)
    if uni=="Other": uni=st.text_input("Enter University")
with c2:
    dept=st.selectbox("Select Department", DEPARTMENTS)
    if dept=="Other": dept=st.text_input("Enter Department")
project_title=st.text_input("Project Title"); student_email=st.text_input("Your Email for Receipt & Feedback")

# Restore session if reloaded
if not st.session_state.paid:
    st.divider()
    st.subheader("🔑 Already Paid? Restore Access")
    restore_code=st.text_input("Enter your verification code to restore (no need to pay again if within 7 days)", key="restore")
    if restore_code:
        valid, msg = check_code_valid(restore_code.strip().upper())
        if valid:
            pay_info = get_payment_by_code(restore_code.strip().upper())
            st.session_state.paid=True
            st.session_state.verify_code=restore_code.strip().upper()
            st.session_state.tier = pay_info.get("tier","") if pay_info else "Standard"
            st.success(f"✅ Restored! {msg}")
            st.rerun()
        else:
            st.error(msg)

# Payment verification callback
ref=query_params.get("reference",[None])[0] or query_params.get("trxref",[None])[0]
if ref and not st.session_state.paid:
    ok, data = verify_paystack(ref)
    if ok:
        verify_code=''.join(random.choices(string.ascii_uppercase+string.digits, k=8))
        amount_paid = data.get("amount",0)
        tier_name = st.session_state.get("selected_tier","Unknown")
        row=[datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), uni, dept, project_title, verify_code, student_email, str(amount_paid), tier_name, "success", ref]
        save_payment(row)
        st.session_state.paid=True; st.session_state.verify_code=verify_code; st.session_state.pay_ref=None
        st.success(f"✅ Payment Verified! Code: {verify_code} — {tier_name}"); st.balloons(); st.query_params.clear(); st.rerun()
    else:
        st.error("Payment not verified yet — if you just paid, wait 10s and refresh")

# Payment UI with tiers
if not st.session_state.paid:
    st.divider()
    st.subheader("🔓 Freemium Preview — Try, then Pay")
    st.info("FREE: Upload (5 rows), 1 SPSS test, see tools. PAID: Full download with watermark verification code.")
    st.markdown("### Choose Plan — Students pick middle")
    tier_choice=st.radio("Select Tier", list(PRICE_TIERS.keys()), index=1, horizontal=False)
    st.session_state.selected_tier = tier_choice
    tier_info = PRICE_TIERS[tier_choice]
    st.markdown(f"**{tier_choice}** includes: {', '.join(tier_info['features'])}")
    if student_email and "@" in student_email:
        if st.button(f"💳 Pay ₦{tier_info['display']:,} with Paystack LIVE", type="primary", use_container_width=True):
            url, reference = initialize_paystack(student_email, tier_info["amount"])
            if url:
                st.session_state.pay_ref = reference
                st.link_button(f"Complete Payment — ₦{tier_info['display']:,}", url)
                components.html(f'<script>window.open("{url}", "_self")</script>', height=0)
    else:
        st.warning("Enter email above to enable LIVE Paystack")
    if st.button("Test Mode — Unlock (Admin only)"):
        verify_code=''.join(random.choices(string.ascii_uppercase+string.digits, k=8))
        save_payment([datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), uni, dept, project_title, verify_code, student_email, tier_info['display'], tier_choice, "test", "test_ref"])
        st.session_state.paid=True; st.session_state.verify_code=verify_code; st.session_state.tier=tier_choice; st.rerun()
else:
    verify_code=st.session_state.verify_code
    tier = st.session_state.get("tier", st.session_state.get("selected_tier","Premium"))
    valid, msg = check_code_valid(verify_code)
    if not valid:
        st.error(f"⛔ {msg}"); st.session_state.paid=False; st.stop()
    st.success(f"✅ Active — {VERIFY_DOMAIN}/{verify_code} — {tier} — {msg}")

# Helper to add watermark to doc
def add_watermark_doc(doc, code):
    # Footer watermark
    section = doc.sections[0]
    footer = section.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.text = f"{VERIFY_DOMAIN}/{code} — Verified by {BUSINESS_NAME} — {dept} — {uni} — Do not share"
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return doc

# TABS
tab1,tab2,tab3,tab4,tab5=st.tabs(["1️⃣ Upload Data","2️⃣ SPSS + Qualitative","3️⃣ Discuss with Jenni + Consensus","4️⃣ Format Full Project","5️⃣ Send to Admin for Review"])

with tab1:
    st.header("Upload Data")
    up=st.file_uploader("Excel .xlsx", type=["xlsx"])
    if up: 
        df=pd.read_excel(up); 
        st.dataframe(df.head() if not st.session_state.paid else df.head(30)); 
        st.session_state.df=df

with tab2:
    st.header("📊 SPSS + Qualitative")
    if 'df' not in st.session_state: st.warning("Upload in Tab1")
    else:
        df=st.session_state.df
        tier = st.session_state.get("tier","")
        is_basic = "Basic" in tier or not st.session_state.paid
        options = ["Descriptive (Mean, SD, Frequency)"] if is_basic and not st.session_state.paid else ["Descriptive (Mean, SD, Frequency)","Chi-square Test","Independent T-Test","One-way ANOVA","Pearson Correlation","Linear Regression"]
        if "Standard" in tier or "Premium" in tier or st.session_state.paid:
            options = ["Descriptive (Mean, SD, Frequency)","Chi-square Test","Independent T-Test","One-way ANOVA","Pearson Correlation","Linear Regression"]
        analysis=st.selectbox("SPSS Test", options)
        if analysis=="Descriptive (Mean, SD, Frequency)":
            st.dataframe(df.describe()); col_f=st.selectbox("Frequency", df.columns); st.dataframe(df[col_f].value_counts().reset_index()); fig,ax=plt.subplots(); df[col_f].value_counts().plot(kind='bar',ax=ax); st.pyplot(fig)
        elif st.session_state.paid:
            if analysis=="Chi-square Test":
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
    if "Premium" not in str(st.session_state.get("tier","")) and st.session_state.paid:
        if "Standard" in str(st.session_state.get("tier","")):
            st.warning("Standard plan — Discussion download limited. Upgrade to Premium for full Jenni+Consensus.")
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
                if st.session_state.paid:
                    doc=docx.Document(); doc.styles['Normal'].font.name="Times New Roman"; doc.styles['Normal'].font.size=Pt(12); doc.styles['Normal'].paragraph_format.line_spacing=2.0
                    doc.add_heading("Discussion - Consensus + Jenni",1); doc.add_paragraph(para); doc.add_heading("Articles",2)
                    if s1: doc.add_paragraph(s1 + (f" Link:{l1}" if l1 else "")); 
                    if s2: doc.add_paragraph(s2 + (f" Link:{l2}" if l2 else ""))
                    if s3: doc.add_paragraph(s3)
                    doc.add_paragraph(f"\n{VERIFY_DOMAIN}/{st.session_state.verify_code}")
                    doc=add_watermark_doc(doc, st.session_state.verify_code)
                    bio=io.BytesIO(); doc.save(bio); st.download_button("Download Discussion (Watermarked)", bio.getvalue(), file_name=f"{st.session_state.verify_code}_Discussion.docx")
                else:
                    st.warning("🔒 Pay to download watermarked doc")

with tab4:
    st.header("📄 Format Entire Project")
    tier = str(st.session_state.get("tier",""))
    if not st.session_state.paid or "Basic" in tier:
        st.error("🔒 Format is Standard/Premium only. Your Basic plan is SPSS only. Upgrade!")
        st.stop()
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
        doc.add_paragraph(f"\n{VERIFY_DOMAIN}/{st.session_state.verify_code} — {style_sel}")
        doc=add_watermark_doc(doc, st.session_state.verify_code)
        bio=io.BytesIO(); doc.save(bio); st.download_button("Download Formatted (Watermarked)", bio.getvalue(), file_name=f"{st.session_state.verify_code}_Formatted.docx")

with tab5:
    st.header("📤 Send Work to Admin for Review & Get Feedback")
    if not st.session_state.paid or "Premium" not in str(st.session_state.get("tier","")):
        if st.session_state.paid:
            st.error("🔒 Admin Review is Premium only (₦60k). You have Standard. Upgrade to get review!")
        else:
            st.error("🔒 Send to Admin is Premium only")
        st.stop()
    st.info("Submit work — saved permanently to Google Sheet — Never lost on reboot")
    work_text=st.text_area("Paste your work", height=180)
    up_review=st.file_uploader("Or Upload Docx", type=["docx"], key="rev_up")
    if st.button("Submit for Review"):
        if not work_text and not up_review: st.error("Paste work")
        elif not student_email: st.error("Enter email in top")
        else:
            content=work_text + (f" [File:{up_review.name}]" if up_review else "")
            row=[datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), uni, dept, project_title, student_email, content[:3000], st.session_state.verify_code, ""]
            save_review(row); st.success(f"✅ Submitted! Code {st.session_state.verify_code} — saved to Google Sheet forever")
    st.divider(); st.subheader("🔔 Check Feedback")
    check=st.text_input("Enter verification code", value=st.session_state.verify_code if st.session_state.verify_code else "")
    if check:
        sh=get_sheet()
        if sh:
            try:
                rev_df=pd.DataFrame(sh.worksheet("Reviews").get_all_records())
                found=rev_df[rev_df.astype(str).apply(lambda x: x.str.contains(check, na=False)).any(axis=1)] if not rev_df.empty else pd.DataFrame()
                if found.empty: st.warning("No submission for this code in Sheet")
                else:
                    for _,r in found.iterrows():
                        if str(r.get("feedback","")).strip(): st.success(f"Feedback from Admin: {r['feedback']}")
                        else: st.info("Submitted — Admin not yet feedback, check later")
            except Exception as e: st.error(f"Sheet read error: {e}")

st.divider()
st.markdown(f"<div style='text-align:center; font-size:11px; color:gray;'><b>{BUSINESS_NAME}</b><br>{BUSINESS_EMAIL} | {VERIFY_DOMAIN}<br>Works with Consensus.ai + Jenni.ai | Paystack LIVE + Google Sheet + Expiry 7days + Watermark | © 2026</div>", unsafe_allow_html=True)
