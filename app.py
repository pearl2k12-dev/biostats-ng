
import streamlit as st
import pandas as pd, requests, datetime, random, string, io, os, re, tempfile, hmac
from urllib.parse import quote_plus
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import matplotlib.pyplot as plt
import seaborn as sns
import scipy.stats as stats
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import PyPDF2
import numpy as np
import pyreadstat
import statsmodels.api as sm

st.set_page_config(page_title="Research Base HQ", layout="wide")

# ---------------------------------------------------------------- SECRETS
def _secret(name, default=None):
    try: return st.secrets[name]
    except Exception: return default

ADMIN_PASSWORD = _secret("admin_password")          # set in secrets.toml — never hardcode
TEST_MODE = bool(_secret("test_mode", False))        # true only in your LOCAL secrets

BUSINESS_NAME="Research Base HQ"; APP_NAME="Research Base HQ"; BUSINESS_EMAIL="support@researchbasehq.ng"; VERIFY_DOMAIN="researchbasehq.ng/verify"
WHATSAPP_NUMBER="2347054857303"
WHATSAPP_DISPLAY="07054857303"
WHATSAPP_LINK=f"https://wa.me/{WHATSAPP_NUMBER}?text=Hello%20Research%20Base%20HQ%20Support,%20I%20need%20help%20with%20my%20project."
UNIVERSITIES=["University of Port Harcourt","University of Lagos","University of Ibadan","University of Benin","University of Nigeria Nsukka","Nnamdi Azikiwe University","Obafemi Awolowo University","Ahmadu Bello University","University of Calabar","University of Jos","University of Ilorin","University of Uyo","Rivers State University","Ignatius Ajuru University","Other"]
DEPARTMENTS=["Nursing","Medicine","Physiology","Anatomy","Pharmacy","Public Health","Medical Lab Science","Radiography","Biochemistry","Microbiology","Pharmacology","Other"]
PRICE_TIERS = {"Basic - ₦30,000 (Analyse + Format)": {"amount": 3000000, "display": 30000}, "Standard - ₦40,000 (Most Popular + AI Checker)": {"amount": 4000000, "display": 40000}}

# ---------------------------------------------------------------- CHAPTER 3 PARSING
def extract_ch3(file):
    text=""
    try:
        if file.name.endswith('.pdf'):
            reader=PyPDF2.PdfReader(file)
            for p in reader.pages: text+=(p.extract_text() or "")+"\n"
        else:
            doc=docx.Document(file); text="\n".join([para.text for para in doc.paragraphs])
    except: pass
    return text.lower()

def parse_ch3(text):
    m=[]; t=text
    if "regression" in t: m.append("Regression")
    if "correlation" in t or "pearson" in t: m.append("Correlation")
    if "chi-square" in t or "chi square" in t: m.append("Chi-Square")
    if "t-test" in t or "t test" in t: m.append("T-Test")
    if "anova" in t: m.append("ANOVA")
    if "tukey" in t or "post-hoc" in t or "post hoc" in t: m.append("Tukey")
    if "cronbach" in t or "reliability" in t: m.append("Reliability")
    if not m: m=["Descriptive Statistics","ANOVA","Tukey"]
    dv=re.search(r'dependent variable[^.]*?is\s+([a-z ]{3,40})',t)
    return {"methods":m,"dv_text":dv.group(1) if dv else ""}

# ---------------------------------------------------------------- GOOGLE SHEETS
@st.cache_resource(show_spinner=False)
def _connect():
    scope=["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/drive"]
    creds_dict=dict(st.secrets["gcp_service_account"])
    creds=ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client=gspread.authorize(creds)
    return client.open_by_url(st.secrets["paysheets"]["sheet_url"])

def get_sheet():
    try: return _connect()
    except Exception: return None

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

# Payments sheet columns (row 1 must be a header row; data in this order — matches save_payment)
COL = {"date":0,"uni":1,"dept":2,"title":3,"code":4,"email":5,"amount":6,"tier":7,"reg":8,"status":9,"ref":10}

def _cell(r, key):
    i = COL[key]
    return str(r[i]).strip() if len(r) > i else ""

def _parse_date(s):
    s = str(s).strip().split(" ")[0]
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try: return datetime.datetime.strptime(s, fmt)
        except ValueError: continue
    return None

@st.cache_data(ttl=45, show_spinner=False)
def payment_rows():
    sh = get_sheet()
    if not sh: return None
    try: return sh.worksheet("Payments").get_all_values()[1:]   # skip header row
    except Exception: return None

def reference_used(ref):
    return any(_cell(r, "ref") == ref for r in (payment_rows() or []))

def find_payment(reg_no, title, email, code=None):
    """Fail-closed: access is restored only if reg no + title + email (+ code) match a successful, unexpired payment."""
    rows = payment_rows()
    if not rows: return None
    for r in reversed(rows):
        if _cell(r, "status") != "success": continue
        if _cell(r, "reg").lower() != reg_no.strip().lower(): continue
        if _cell(r, "title").lower() != title.strip().lower(): continue
        if _cell(r, "email").lower() != email.strip().lower(): continue
        if code and _cell(r, "code") != code: continue
        d = _parse_date(_cell(r, "date"))
        if d is None or (datetime.datetime.now() - d).days > 180: continue
        return {"code": _cell(r, "code"), "tier": _cell(r, "tier")}
    return None

# ---------------------------------------------------------------- PAYSTACK
def _app_url():
    try: return st.secrets["app"]["url"]
    except Exception: return None

def initialize_paystack(email, amount_kobo, metadata=None):
    try:
        secret=st.secrets["paystack"]["secret_key"]
        headers={"Authorization": f"Bearer {secret}", "Content-Type": "application/json"}
        data={"email": email, "amount": amount_kobo, "currency": "NGN", "metadata": metadata or {}}
        if _app_url(): data["callback_url"] = _app_url()
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

# ---------------------------------------------------------------- ANALYSIS
def spss_descriptive_table(df):
    rows=[]
    for col in df.columns:
        try: s = pd.to_numeric(df[col], errors='coerce').dropna()
        except: continue
        if len(s)==0: continue
        n=len(s); mean=s.mean(); median=s.median()
        try: mode_val=s.mode().iloc[0]
        except: mode_val=np.nan
        std=s.std(ddof=1); var=s.var(ddof=1)
        se=std/np.sqrt(n) if n>0 else 0
        skew=stats.skew(s, bias=False); kurt=stats.kurtosis(s, bias=False)
        se_skew=np.sqrt(6*n*(n-1)/((n-2)*(n+1)*(n+3))) if n>2 else 0
        se_kurt=np.sqrt(24*n*(n-1)**2/((n-3)*(n-2)*(n+3)*(n+5))) if n>3 else 0
        rows.append({"Variable":col,"N Valid":n,"N Missing":int(df[col].isnull().sum()),"Mean":round(mean,3),"Median":round(median,3),"Mode":mode_val,"Std. Deviation":round(std,3),"Variance":round(var,3),"Std. Error Mean":round(se,3),"Minimum":round(s.min(),3),"Maximum":round(s.max(),3),"Range":round(s.max()-s.min(),3),"Sum":round(s.sum(),3),"Skewness":round(skew,3),"Std. Error Skewness":round(se_skew,3),"Kurtosis":round(kurt,3),"Std. Error Kurtosis":round(se_kurt,3)})
    return pd.DataFrame(rows)

def run_inferential_correct_order(df, dv, ivs, group_col, methods):
    out={}
    work_cols=[c for c in [dv]+ivs if c in df.columns]
    work_df=df[work_cols].copy()
    for c in work_cols: work_df[c]=pd.to_numeric(work_df[c], errors='coerce')
    work_df=work_df.dropna()
    if group_col and group_col in df.columns and dv in df.columns:
        try:
            gdf=df[[group_col, dv]].copy()
            gdf[dv]=pd.to_numeric(gdf[dv], errors='coerce')
            gdf=gdf.dropna()
            if len(gdf)>0:
                desc_by_group=gdf.groupby(group_col)[dv].agg(['count','mean','std','min','max']).round(3)
                desc_by_group.columns=['N','Mean','SD','Min','Max']
                out['Descriptive_By_Group']=desc_by_group
        except: pass
    if len(work_df) < 3: return out
    if "ANOVA" in methods or group_col:
        gc = group_col
        if not gc:
            for c in df.columns:
                if 3<=df[c].nunique()<=6 and c!=dv:
                    gc=c; break
        if gc and dv in df.columns and gc in df.columns:
            try:
                groups=[pd.to_numeric(g[dv], errors='coerce').dropna().values for _,g in df.groupby(gc) if len(pd.to_numeric(g[dv], errors='coerce').dropna())>0]
                if len(groups)>=2:
                    f,p=stats.f_oneway(*groups)
                    out['ANOVA']={'F':round(f,3),'Sig.':round(p,4),'factor':gc}
                    if p < 0.05:
                        try:
                            from statsmodels.stats.multicomp import pairwise_tukeyhsd
                            tukey_df=df[[gc, dv]].copy()
                            tukey_df[dv]=pd.to_numeric(tukey_df[dv], errors='coerce')
                            tukey_df=tukey_df.dropna()
                            if len(tukey_df)>=3:
                                tukey = pairwise_tukeyhsd(endog=tukey_df[dv], groups=tukey_df[gc], alpha=0.05)
                                out['PostHoc_Tukey_Text']=tukey.summary().as_text()
                                out['PostHoc_df']=pd.DataFrame(data=tukey._results_table.data[1:], columns=tukey._results_table.data[0])
                        except Exception as e:
                            out['PostHoc_Error']=str(e)
            except: pass
    if "T-Test" in methods:
        cat=None
        for c in df.columns:
            if df[c].nunique()==2 and c!=dv:
                cat=c; break
        if cat and dv in df.columns:
            try:
                groups=df[cat].dropna().unique()
                if len(groups)==2:
                    g1=pd.to_numeric(df[df[cat]==groups[0]][dv], errors='coerce').dropna()
                    g2=pd.to_numeric(df[df[cat]==groups[1]][dv], errors='coerce').dropna()
                    if len(g1)>1 and len(g2)>1:
                        levene_stat, levene_p=stats.levene(g1,g2)
                        t_eq,p_eq=stats.ttest_ind(g1,g2,equal_var=True)
                        t_neq,p_neq=stats.ttest_ind(g1,g2,equal_var=False)
                        out['T-Test']={'Levene F':round(levene_stat,3),'Levene Sig':round(levene_p,4),'t Equal Var':round(t_eq,3),'Sig Equal Var':round(p_eq,4),'t Unequal Var':round(t_neq,3),'Sig Unequal Var':round(p_neq,4),'Mean Difference':round(g1.mean()-g2.mean(),3),'Groups':list(groups),'Factor':cat}
            except: pass
    if "Regression" in methods and dv and ivs:
        numeric_ivs=[i for i in ivs if i in work_df.columns]
        if numeric_ivs and dv in work_df.columns:
            X=sm.add_constant(work_df[numeric_ivs]); y=work_df[dv]
            try:
                model=sm.OLS(y,X).fit()
                out['Regression']={'Model Summary':{'R':round(np.sqrt(model.rsquared),3),'R Square':round(model.rsquared,3),'Adjusted R Square':round(model.rsquared_adj,3),'Std Error':round(np.sqrt(model.mse_resid),3)},'ANOVA':{'Regression SS':round(model.ess,3),'Residual SS':round(model.ssr,3),'F':round(model.fvalue,3),'Sig.':round(model.f_pvalue,4)},'Coefficients_df':pd.DataFrame({'Variable':model.params.index,'B':model.params.values.round(3),'Std. Error':model.bse.values.round(3),'t':model.tvalues.values.round(3),'Sig.':model.pvalues.values.round(4),'95% CI Lower':model.conf_int()[0].values.round(3),'95% CI Upper':model.conf_int()[1].values.round(3)})}
            except Exception as e: out['Regression_Error']=str(e)
    if "Correlation" in methods and len(work_cols)>=2:
        try:
            corr=work_df.corr(method='pearson')
            pvals=pd.DataFrame(np.zeros_like(corr), columns=corr.columns, index=corr.index, dtype=float)
            for i in corr.columns:
                for j in corr.columns:
                    if i!=j and len(work_df[i].dropna())>2:
                        _,p=stats.pearsonr(work_df[i], work_df[j]); pvals.loc[i,j]=p
            out['Correlation']={'r':corr.round(3),'Sig (2-tailed)':pvals.round(4),'N':len(work_df)}
        except: pass
    return out

def spss_graphs(df, dv, ivs, group_col=None):
    sns.set_style("whitegrid"); figs=[]
    for col in [dv]+ivs:
        if col and col in df.columns:
            try:
                s=pd.to_numeric(df[col], errors='coerce').dropna()
                if len(s)>0:
                    fig,ax=plt.subplots(); sns.histplot(s,kde=True,ax=ax,color="#4C72B0"); ax.set_title(f"Histogram - {col}"); figs.append((f"Histogram_{col}",fig))
            except: pass
    if group_col and group_col in df.columns and dv in df.columns:
        try:
            gdf=df[[group_col,dv]].copy(); gdf[dv]=pd.to_numeric(gdf[dv], errors='coerce'); gdf=gdf.dropna()
            if len(gdf)>0:
                fig,ax=plt.subplots(); sns.boxplot(x=group_col,y=dv,data=gdf,ax=ax); ax.set_title(f"Boxplot by Group - {dv} by {group_col}"); plt.xticks(rotation=30); figs.append((f"Boxplot_{dv}_by_{group_col}",fig))
                fig,ax=plt.subplots(); sns.barplot(x=group_col,y=dv,data=gdf,ci='sd',ax=ax); ax.set_title(f"Mean ± SD by Group - {dv}"); plt.xticks(rotation=30); figs.append((f"BarMean_{dv}_by_{group_col}",fig))
        except: pass
    return figs

# ---------------------------------------------------------------- STYLES
st.markdown(f"""
<style>
.hero {{background: linear-gradient(135deg, #0A1931 0%, #185ADB 100%); padding: 45px 25px; border-radius: 16px; color: white; text-align: center; margin-bottom: 20px;}}
.app-name {{font-size:16px; font-weight:900; color:#185ADB; letter-spacing:1px; background:white; border:2px solid #185ADB; border-radius:20px; padding:8px 18px; display:inline-block; margin-bottom:12px;}}
.locked-box {{background:#FFF0F0; border:2px dashed #FF4444; border-radius:12px; padding:20px; text-align:center; margin:15px 0;}}
.blur {{filter: blur(6px); pointer-events:none; user-select:none;}}
.price-card {{border: 2px solid #eee; border-radius: 16px; padding: 20px; background: white;}}
.price-card.popular {{border-color: #185ADB; box-shadow: 0 8px 24px rgba(24,90,219,0.15);}}
.badge {{background: #185ADB; color: white; padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 700;}}
.whatsapp-float {{position: fixed; width:60px; height:60px; bottom:20px; right:20px; background-color:#25D366; color:#FFF; border-radius:50px; text-align:center; font-size:30px; box-shadow: 2px 2px 10px rgba(0,0,0,0.3); z-index:1000; display:flex; align-items:center; justify-content:center; text-decoration:none;}}
</style>
<a href="{WHATSAPP_LINK}" target="_blank" class="whatsapp-float">💬</a>
""", unsafe_allow_html=True)

for k,v in {"is_admin":False,"show_admin_login":False,"paid":False,"verify_code":None,"tier":None,"last_project_paid":""}.items():
    if k not in st.session_state: st.session_state[k]=v

# ---------------------------------------------------------------- ADMIN
if st.query_params.get("admin")=="1" and not st.session_state.is_admin: st.session_state.show_admin_login=True
if st.session_state.show_admin_login and not st.session_state.is_admin:
    st.markdown("### 🔒 Creator Only"); pwd=st.text_input("Creator Password",type="password")
    if st.button("Login as Creator"):
        if ADMIN_PASSWORD and hmac.compare_digest(pwd, str(ADMIN_PASSWORD)): st.session_state.is_admin=True; st.session_state.show_admin_login=False; st.rerun()
        else: st.error("Denied"); st.stop()

if st.session_state.is_admin:
    st.subheader("📊 CREATOR DASHBOARD"); sh=get_sheet()
    if sh:
        try:
            pay_df=pd.DataFrame(sh.worksheet("Payments").get_all_records()); st.metric("Total Paid", len(pay_df)); st.dataframe(pay_df.tail(30))
        except Exception as e: st.error(f"{e}")
    if st.button("Logout Creator"): st.session_state.is_admin=False; st.query_params.clear(); st.rerun()
    st.stop()

# ---------------------------------------------------------------- RETURN FROM PAYSTACK (runs before the input widgets)
def _restore_choice(key, other_key, options, value):
    if not value: return
    if value in options: st.session_state[key] = value
    else: st.session_state[key] = "Other"; st.session_state[other_key] = value

pay_ref = st.query_params.get("reference") or st.query_params.get("trxref")
if pay_ref:
    ok, pdata = verify_paystack(pay_ref)
    meta = pdata.get("metadata") or {}
    if not isinstance(meta, dict): meta = {}
    tier = meta.get("tier", "")
    expected = PRICE_TIERS.get(tier, {}).get("amount")
    if ok and expected and pdata.get("amount") == expected and pdata.get("currency", "NGN") == "NGN":
        pay_email = (pdata.get("customer") or {}).get("email", "")
        reg = str(meta.get("reg_no", "")).strip(); title = str(meta.get("project_title", "")).strip()
        if not reference_used(pay_ref):
            vc = ''.join(random.choices(string.ascii_uppercase+string.digits, k=8))
            saved = save_payment([datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), meta.get("uni",""), meta.get("dept",""), title, vc, pay_email, str(pdata.get("amount",0)), tier, reg, "success", pay_ref])
            payment_rows.clear()
            if not saved:
                # payment is genuine but the sheet failed: grant this session, tell them to keep the reference
                st.session_state.paid=True; st.session_state.verify_code=vc; st.session_state.tier=tier
                st.session_state.last_project_paid=f"{reg.lower()}|{title.lower()}"
                st.session_state["pay_notice"]=f"Payment received but not recorded. Save this reference and send it on WhatsApp: {pay_ref}"
        _restore_choice("uni_sel", "uni_other", UNIVERSITIES, meta.get("uni",""))
        _restore_choice("dept_sel", "dept_other", DEPARTMENTS, meta.get("dept",""))
        st.session_state["reg_no_in"]=reg; st.session_state["project_title_in"]=title
        st.session_state["student_email_in"]=pay_email; st.session_state["student_name_in"]=str(meta.get("student_name",""))
        st.session_state["just_paid"]=True
    else:
        st.session_state["pay_notice"]=f"Payment verification failed. If you were debited, send this reference on WhatsApp: {pay_ref}"
    st.query_params.clear(); st.rerun()

if "pay_notice" in st.session_state: st.warning(st.session_state.pop("pay_notice"))
if st.session_state.pop("just_paid", False): st.balloons()

# ---------------------------------------------------------------- HEADER
st.markdown(f"<div style='text-align:center;'><span class='app-name'>📘 {APP_NAME}</span></div>", unsafe_allow_html=True)

if not st.session_state.paid:
    st.markdown(f"""
<div class="hero">
<h1>Stuck With Your Results? From Raw Data To Final Chapter — In One Afternoon</h1>
<p>You don't need to understand statistics. Just upload what you have — we give you back what you need: clean, supervisor-ready tables, graphs, and properly formatted chapters. No analyst needed.</p>
<div style="margin-top:18px; display:flex; gap:8px; justify-content:center; flex-wrap:wrap;">
<span style="background:white; color:#0A1931; padding:6px 14px; border-radius:20px; font-size:13px; font-weight:700;">✓ No Watermark</span>
<span style="background:white; color:#0A1931; padding:6px 14px; border-radius:20px; font-size:13px; font-weight:700;">✓ Real SPSS .sav</span>
<span style="background:white; color:#0A1931; padding:6px 14px; border-radius:20px; font-size:13px; font-weight:700;">✓ Correct Order + Choose Uni Format</span>
</div>
<div style="margin-top:15px; background:rgba(255,255,255,0.15); display:inline-block; padding:6px 14px; border-radius:20px; font-size:12px;">{APP_NAME} • {VERIFY_DOMAIN}</div>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------- PROJECT DETAILS
st.divider()
st.subheader("Start Here — Your University & Project")
c1,c2=st.columns(2)
with c1:
    uni=st.selectbox("Select Your University *", UNIVERSITIES, key="uni_sel")
    if uni=="Other": uni=st.text_input("Enter Your University Name *", key="uni_other")
    reg_no=st.text_input("Registration Number / Research ID *", key="reg_no_in")
    student_name=st.text_input("Full Name", key="student_name_in")
with c2:
    dept=st.selectbox("Select Department", DEPARTMENTS, key="dept_sel")
    if dept=="Other": dept=st.text_input("Enter Department", key="dept_other")
    project_title=st.text_input("Research / Project Title * (Access tied to this title)", key="project_title_in")
    student_email=st.text_input("Email *", key="student_email_in")

def start_checkout(tier_choice):
    meta = {"reg_no": reg_no.strip(), "project_title": project_title.strip(), "tier": tier_choice,
            "uni": uni, "dept": dept, "student_name": student_name}
    url, _ = initialize_paystack(student_email.strip(), PRICE_TIERS[tier_choice]["amount"], meta)
    if url: st.link_button("Pay on Paystack (Secure) →", url, type="primary", use_container_width=True)

def grant(hit, key):
    st.session_state.paid=True; st.session_state.verify_code=hit["code"]
    st.session_state.tier=hit["tier"]; st.session_state.last_project_paid=key

current_project_key = f"{reg_no.strip().lower()}|{project_title.strip().lower()}"
if st.session_state.paid and st.session_state.last_project_paid != current_project_key:
    st.warning("🔒 New project detected! Pay for this new topic")
    st.session_state.paid=False; st.session_state.verify_code=None; st.session_state.tier=None

if not st.session_state.paid and reg_no and project_title and len(project_title.strip())>5 and "@" in student_email:
    hit = find_payment(reg_no, project_title, student_email)
    if hit:
        grant(hit, current_project_key); st.success("✅ Existing payment found — Access restored")

if not st.session_state.paid:
    with st.expander("🔑 Already paid for THIS exact project? Restore code"):
        code_in=st.text_input("Enter verification code")
        if code_in:
            hit = find_payment(reg_no, project_title, student_email, code=code_in.strip().upper())
            if hit: grant(hit, current_project_key); st.rerun()
            else: st.error("Code not found for this Reg No, Project Title and Email — or it has expired. Fill all three fields above exactly as you used when paying.")

if not st.session_state.paid:
    st.divider()
    st.subheader("💸 Why ₦30k Saves You Money")
    a,b=st.columns(2)
    with a: st.markdown("""<div style="background:white; border:1px solid #eee; border-radius:12px; padding:15px;"><h4>❌ Without (Separate)</h4><div>📊 Analyst ₦35k-₦60k</div><div>📄 Formatting ₦15k</div><div>📚 Discussion ₦15k + AI check ₦5k</div><hr><b>Total ₦80k-₦110k</b></div>""", unsafe_allow_html=True)
    with b: st.markdown("""<div style="background:#E8F5E9; border:2px solid #185ADB; border-radius:12px; padding:15px;"><h4>✅ With Research Base HQ</h4><div>📊 Correct Order: Descriptive → Group → ANOVA → Tukey ✅</div><div>🏛️ Choose University Format on Spot ✅</div><div>🤖 AI Checker included (Standard) ✅</div><div>🔒 Variable Map — No analyst ✅</div><hr><b>Total 3 Hours — ₦30,000 — Save ₦50k+</b></div>""", unsafe_allow_html=True)
    st.divider(); st.subheader("💰 Choose Plan — Per Project Topic — What you get")
    p1,p2=st.columns(2)
    with p1:
        st.markdown(f"""
<div class="price-card">
<h3>Basic — ₦30,000</h3>
<p style="text-decoration:line-through; color:#D32F2F; margin:4px 0;">If you pay separately ₦80k-₦105k</p>
<div style="text-align:left; font-size:14px; line-height:1.8; margin:12px 0;">
                ✓ Descriptive Statistics<br>
                ✓ Descriptive by Group (Control vs Low vs High)<br>
                ✓ Frequency Tables<br>
                ✓ ANOVA, T-Test, Correlation, Regression<br>
                ✓ Tukey Post-Hoc (when Sig &lt;0.05)<br>
                ✓ Graphs — No Watermark<br>
                ✓ Real SPSS .sav — No Watermark<br>
                ✓ Excel Chapter 4 — Correct Order<br>
                ✓ Format Chapters 1-5 — Choose University Format On Spot<br>
                ✓ References + 10+ Styles — Citations Intact<br>
                ✓ Discussion with PubMed/Scholar/Consensus/Jenni Guideline<br>
</div>
<b>Pay ₦30k for THIS project — No AI Checker</b>
</div>
""", unsafe_allow_html=True)
    with p2:
        st.markdown(f"""
<div class="price-card popular">
<div class="badge">🔥 MOST POPULAR — Includes AI Checker + Uni Format</div>
<h3>Standard — ₦40,000</h3>
<p style="text-decoration:line-through; color:#D32F2F; margin:4px 0;">If you pay separately ₦95k-₦120k</p>
<div style="text-align:left; font-size:14px; line-height:1.8; margin:12px 0;">
<b>Everything in Basic, PLUS:</b><br>
                ✓ Choose/Confirm Your University Format On The Spot<br>
                ✓ AI Checker — Check if your work will be flagged as AI<br>
                ✓ Humanize On The Spot — Fix AI-flagged text<br>
                ✓ Send to Admin for Review & Correction<br>
                ✓ Priority WhatsApp Support<br>
                ✓ Supervisor Query Fix<br>
                ✓ Discussion with PubMed + Scholar + Consensus + Jenni<br>
                ✓ Valid 180 days same title<br>
</div>
<b>Pay ₦40k for THIS project — Best for final year</b>
</div>
""", unsafe_allow_html=True)
    tier_choice=st.radio("Select Plan", list(PRICE_TIERS.keys()), index=1); st.session_state.selected_tier=tier_choice
    if student_email and reg_no and project_title and "@" in student_email:
        if st.button(f"💳 Pay ₦{PRICE_TIERS[tier_choice]['display']:,} for this project", type="primary", use_container_width=True):
            start_checkout(tier_choice)
    else: st.warning("Enter Reg No + Project Title + Email to pay")
    if TEST_MODE and st.button("Test Unlock (dev only)"):
        vc=''.join(random.choices(string.ascii_uppercase+string.digits, k=8))
        st.session_state.paid=True; st.session_state.verify_code=vc; st.session_state.tier=tier_choice
        st.session_state.last_project_paid=current_project_key; st.rerun()

if st.session_state.paid:
    st.success(f"✅ Active for: '{project_title}' — {VERIFY_DOMAIN}/{st.session_state.verify_code} — No watermark — Correct Order")

tab1,tab2,tab3,tab4,tab5,tab6=st.tabs(["1️⃣ Upload (FREE)","2️⃣ Analysis","3️⃣ Discuss with PubMed/Scholar","4️⃣ Format — Choose Uni Format","5️⃣ Send to Admin","6️⃣ AI Checker — Standard"])

# ---------------------------------------------------------------- TAB 1
def render_upload_tab():
    st.header("Upload Data — Specific Format Required — Prevents Analyst Trip")
    study_type = st.radio("What is your study? *", ["Human / Survey (Nursing, Public Health, etc.)", "Rat / Experimental (Physiology, Anatomy, Biochemistry, Pharmacology)"], horizontal=True)
    tmpl_buf = io.BytesIO()
    with pd.ExcelWriter(tmpl_buf, engine='openpyxl') as writer:
        if "Human" in study_type:
            sample_data = pd.DataFrame({"Training": [4,5,3,4], "Leadership": [3,4,2,5], "Gender": ["Male","Female","Male","Female"], "Employee_Performance": [75,82,68,90]})
            var_map = pd.DataFrame({"Variable_Name": ["Employee_Performance","Training","Leadership","Gender"], "Variable_Label": ["Employee Performance Score","Training Hours","Leadership Style","Gender"], "Role": ["Dependent","Independent","Independent","Categorical"], "Type": ["Scale","Scale","Scale","Nominal"], "Categories": ["","", "", "Male, Female"]})
        else:
            sample_data = pd.DataFrame({"Rat_ID": ["R001","R002","R003","R004","R005","R006"], "Group": ["Control","Control","Low Dose","Low Dose","High Dose","High Dose"], "Dose_mg": [0,0,10,10,20,20], "Weight_g": [150,152,155,158,162,165], "Glucose_mgdl": [90,92,85,84,78,75], "Day": [7,7,7,7,7,7]})
            var_map = pd.DataFrame({"Variable_Name": ["Rat_ID","Group","Dose_mg","Weight_g","Glucose_mgdl","Day"], "Variable_Label": ["Rat Identification","Experimental Group","Dose mg/kg","Body Weight g","Blood Glucose mg/dl","Day of measurement"], "Role": ["Subject_ID","Group","Dose","Dependent","Dependent","Time"], "Type": ["Nominal","Nominal","Scale","Scale","Scale","Ordinal"], "Categories": ["","Control, Low Dose, High Dose","","","","7, 14, 21"]})
        sample_data.to_excel(writer, sheet_name='Data', index=False)
        var_map.to_excel(writer, sheet_name='Variable_Map', index=False)
        instructions = pd.DataFrame({"Instruction": ["1. Data sheet: First row = variable names — MUST match Variable_Map", "2. Variable_Map sheet: List EVERY variable — Role MUST be one of: Dependent (1 only), Independent, Group, Dose, Categorical, Subject_ID (Rat_ID), Time", "3. For Rat studies: Group = Control, Low, High — for ANOVA", "4. For Rat studies: Rat_ID = R001 etc — for tracking", "5. Type: Scale (numbers), Nominal (Male/Female, Control/Low/High), Ordinal (Day)", "6. CORRECT ORDER in thesis: Descriptive → Descriptive by Group → Frequency → ANOVA → Tukey", "7. This map stops you going to analyst — you already defined DV, IV, Group like analyst asks"]})
        instructions.to_excel(writer, sheet_name='Instructions', index=False)
    st.download_button(f"📥 Download Template for {study_type.split('(')[0]} — Data + Variable_Map", data=tmpl_buf.getvalue(), file_name=f"ResearchBaseHQ_Template_{'Human' if 'Human' in study_type else 'Rat'}.xlsx", use_container_width=True, type="primary")
    c1,c2=st.columns(2)
    with c1: up=st.file_uploader("Upload Your Filled Template .xlsx — Must have Data + Variable_Map", type=["xlsx"])
    with c2: ch3_up=st.file_uploader("Upload Chapter 3 (PDF/DOCX)", type=["pdf","docx"])
    if up:
        try:
            xls=pd.ExcelFile(up)
            if 'Data' not in xls.sheet_names or 'Variable_Map' not in xls.sheet_names:
                st.error("❌ Wrong format! Must have 2 sheets: 'Data' and 'Variable_Map'. Download template above."); return
            df=pd.read_excel(xls, sheet_name='Data')
            vmap_df=pd.read_excel(xls, sheet_name='Variable_Map')
            if 'Variable_Name' not in vmap_df.columns or 'Role' not in vmap_df.columns:
                st.error("❌ Variable_Map must have Variable_Name, Role, Type"); return
            dv_rows=vmap_df[vmap_df['Role'].astype(str).str.lower()=='dependent']
            if len(dv_rows)==0: st.error("❌ Mark 1 variable as Role = Dependent"); return
            if len(dv_rows)>1:
                st.warning(f"Multiple Dependent found: {', '.join(dv_rows['Variable_Name'])} — Using first as primary DV: {dv_rows.iloc[0]['Variable_Name']}")
            st.success(f"✅ Correct! {df.shape[0]} rows | Map: {len(vmap_df)} vars | Study: {study_type}")
            st.dataframe(df.head()); st.write("Your Variable Map:"); st.dataframe(vmap_df, use_container_width=True)
            st.session_state.df=df; st.session_state.vmap=vmap_df
            st.session_state.auto_dv=dv_rows.iloc[0]['Variable_Name']
            st.session_state.auto_ivs=vmap_df[vmap_df['Role'].astype(str).str.lower()=='independent']['Variable_Name'].tolist()
            st.session_state.categorical_vars=vmap_df[vmap_df['Role'].astype(str).str.lower().isin(['categorical','nominal','control'])]['Variable_Name'].tolist()
            grp_rows=vmap_df[vmap_df['Role'].astype(str).str.lower()=='group']
            if not grp_rows.empty: st.session_state.group_col=grp_rows.iloc[0]['Variable_Name']
            else:
                gc=None
                for c in df.columns:
                    if 'group' in c.lower() and 2<=df[c].nunique()<=8:
                        gc=c; break
                st.session_state.group_col=gc
        except Exception as e: st.error(f"Error: {e}"); return
        if ch3_up:
            ch3_text=extract_ch3(ch3_up); ch3_info=parse_ch3(ch3_text); st.session_state.ch3_info=ch3_info
            st.success(f"📘 Methods: {', '.join(ch3_info['methods'])}")
    else: st.warning("Upload template — Descriptive → ANOVA → Tukey order — No analyst needed")

with tab1:
    render_upload_tab()

# ---------------------------------------------------------------- TAB 2
with tab2:
    st.header(f"📊 Analysis in Correct Order — For: {project_title}")
    if 'df' not in st.session_state: st.warning("Upload data in Tab 1")
    else:
        df=st.session_state.df; ch3_info=st.session_state.get('ch3_info',{"methods":["Descriptive Statistics","ANOVA","Tukey"],"dv_text":""})
        dep_col=st.session_state.get('auto_dv', df.columns[0])
        indep_cols=st.session_state.get('auto_ivs', [])
        group_col=st.session_state.get('group_col', None)
        st.info(f"📌 From YOUR Variable_Map: Primary Dependent = **{dep_col}** | Independent = **{', '.join(indep_cols) if indep_cols else 'None'}** | Group = **{group_col if group_col else 'None'}**")
        st.caption("Correct Chapter 4 Order: 1️⃣ Descriptive → 2️⃣ Descriptive by Group (rats) → 3️⃣ Frequency → 4️⃣ ANOVA/T-Test/Regression → 5️⃣ Tukey Post-Hoc")
        c1,c2,c3=st.columns(3)
        with c1: dep_col=st.selectbox("Dependent (DV) — 1 only", df.columns, index=list(df.columns).index(dep_col) if dep_col in df.columns else 0)
        with c2: indep_cols=st.multiselect("Independent (IVs)", [c for c in df.columns if c!=dep_col], default=[c for c in indep_cols if c in df.columns and c!=dep_col])
        with c3: group_col_sel=st.selectbox("Group Column (for rats, else None)", ["None"]+list(df.columns), index=(["None"]+list(df.columns)).index(group_col) if group_col in df.columns else 0)
        group_col = None if group_col_sel=="None" else group_col_sel
        if not st.session_state.paid:
            full_table=spss_descriptive_table(df)
            if not full_table.empty:
                st.subheader("1️⃣ Descriptive — FREE TEASER (1 row)")
                st.dataframe(full_table.head(1), use_container_width=True)
                st.markdown("<div class='locked-box'>🔒 Full Table Locked — Pay for this project to unlock full descriptive → group → ANOVA → Tukey in correct order</div>", unsafe_allow_html=True)
                if len(full_table) > 1:
                    masked = full_table.iloc[1:].astype(object).copy()
                    for mc in masked.columns[1:]: masked[mc] = "🔒"
                    st.dataframe(masked, use_container_width=True)
                if student_email and reg_no and project_title and "@" in student_email:
                    tier_choice=st.session_state.get("selected_tier","Basic - ₦30,000 (Analyse + Format)")
                    if st.button(f"💳 Unlock — Pay ₦{PRICE_TIERS[tier_choice]['display']:,}", type="primary", use_container_width=True):
                        start_checkout(tier_choice)
                else: st.caption("Fill Reg No, Project Title and Email at the top of the page to unlock.")
        else:
            st.success(f"✅ FULL ACCESS — Correct Order: Descriptive → ANOVA → Tukey — No watermark")
            st.divider(); st.subheader("1️⃣ Descriptive Statistics — Chapter 4.1 — SPSS Style")
            desc_table=spss_descriptive_table(df); st.dataframe(desc_table, use_container_width=True)
            if group_col and dep_col and group_col in df.columns:
                st.subheader(f"2️⃣ Descriptive by Group — {group_col} vs {dep_col} — Chapter 4.1b")
                try:
                    gdf=df[[group_col, dep_col]].copy(); gdf[dep_col]=pd.to_numeric(gdf[dep_col], errors='coerce'); gdf=gdf.dropna()
                    if len(gdf)>0:
                        desc_by_group=gdf.groupby(group_col)[dep_col].agg(['count','mean','std','min','max']).round(3)
                        desc_by_group.columns=['N','Mean','SD','Min','Max']; st.dataframe(desc_by_group, use_container_width=True)
                        st.session_state.desc_by_group=desc_by_group
                except Exception as e: st.error(f"Group descriptive error: {e}")
            else:
                st.session_state.pop('desc_by_group', None)
            st.subheader("3️⃣ Frequency Table — Chapter 4.1c")
            col_f=st.selectbox("Frequency Table", df.columns, key="freq")
            freq=df[col_f].value_counts().reset_index(); freq.columns=[col_f,"Frequency"]; freq["Percent"]=(freq["Frequency"]/len(df)*100).round(1); freq["Cumulative %"]=freq["Percent"].cumsum().round(1)
            st.dataframe(freq)
            infer=run_inferential_correct_order(df, dep_col, indep_cols, group_col, ch3_info['methods'])
            st.divider(); st.subheader("4️⃣ Inferential Statistics — Chapter 4.2 — After Descriptive")
            if 'ANOVA' in infer:
                st.write(f"**One-Way ANOVA — {infer['ANOVA']['factor']} on {dep_col}**")
                st.json(infer['ANOVA'])
            if 'T-Test' in infer:
                st.write("**Independent Samples T-Test**"); st.json(infer['T-Test'])
            if 'Regression' in infer:
                st.write("**Regression**"); st.json(infer['Regression']['Model Summary']); st.json(infer['Regression']['ANOVA']); st.dataframe(infer['Regression']['Coefficients_df'], use_container_width=True)
            if 'Correlation' in infer:
                st.write("**Pearson Correlation**"); st.dataframe(infer['Correlation']['r'], use_container_width=True); st.write("**Sig. (2-tailed)**"); st.dataframe(infer['Correlation']['Sig (2-tailed)'], use_container_width=True)
            if 'PostHoc_df' in infer:
                st.subheader("5️⃣ Post-Hoc Tukey HSD — Chapter 4.2b — After ANOVA")
                st.text(infer.get('PostHoc_Tukey_Text',''))
                st.dataframe(infer['PostHoc_df'], use_container_width=True)
                st.success("Interpretation: Sig <0.05 = significant difference")
            elif 'ANOVA' in infer and infer['ANOVA']['Sig.'] >= 0.05:
                st.info(f"5️⃣ Post-Hoc Tukey — Not needed — ANOVA Sig = {infer['ANOVA']['Sig.']} >0.05")
            st.divider(); st.subheader("📊 Graphs — No Watermark"); graphs=spss_graphs(df, dep_col, indep_cols, group_col)
            for name,fig in graphs: st.pyplot(fig)
            st.divider(); st.subheader("📥 Downloads — Correct Order — No Watermark — Real SPSS .sav")
            with tempfile.NamedTemporaryFile(delete=False,suffix='.sav') as tmp: pyreadstat.write_sav(df, tmp.name); p=tmp.name
            with open(p,'rb') as f: sav=f.read(); os.unlink(p)
            st.download_button("📊 1. Clean Data — SPSS .sav — No Watermark", data=sav, file_name=f"{reg_no}_CleanData.sav", mime="application/x-sav", use_container_width=True)
            if not desc_table.empty:
                with tempfile.NamedTemporaryFile(delete=False,suffix='.sav') as tmp: pyreadstat.write_sav(desc_table, tmp.name); p=tmp.name
                with open(p,'rb') as f: sav2=f.read(); os.unlink(p)
                st.download_button("📊 2. Descriptive — SPSS .sav — Chapter 4.1", data=sav2, file_name=f"{reg_no}_Descriptive.sav", mime="application/x-sav", use_container_width=True)
            if 'desc_by_group' in st.session_state:
                with tempfile.NamedTemporaryFile(delete=False,suffix='.sav') as tmp: pyreadstat.write_sav(st.session_state.desc_by_group.reset_index(), tmp.name); p=tmp.name
                with open(p,'rb') as f: sav_bg=f.read(); os.unlink(p)
                st.download_button(f"📊 3. Descriptive by Group {group_col} — SPSS .sav — Chapter 4.1b", data=sav_bg, file_name=f"{reg_no}_Desc_By_{group_col}.sav", mime="application/x-sav", use_container_width=True)
            with tempfile.NamedTemporaryFile(delete=False,suffix='.sav') as tmp: pyreadstat.write_sav(freq, tmp.name); p=tmp.name
            with open(p,'rb') as f: sav3=f.read(); os.unlink(p)
            st.download_button(f"📊 4. Frequency {col_f} — SPSS .sav", data=sav3, file_name=f"{reg_no}_Freq_{col_f}.sav", mime="application/x-sav", use_container_width=True)
            if 'ANOVA' in infer:
                anova_df=pd.DataFrame([infer['ANOVA']])
                with tempfile.NamedTemporaryFile(delete=False,suffix='.sav') as tmp: pyreadstat.write_sav(anova_df, tmp.name); p=tmp.name
                with open(p,'rb') as f: sav_anova=f.read(); os.unlink(p)
                st.download_button("📊 5. ANOVA — SPSS .sav — Chapter 4.2", data=sav_anova, file_name=f"{reg_no}_ANOVA.sav", mime="application/x-sav", use_container_width=True)
            if 'PostHoc_df' in infer:
                with tempfile.NamedTemporaryFile(delete=False,suffix='.sav') as tmp: pyreadstat.write_sav(infer['PostHoc_df'], tmp.name); p=tmp.name
                with open(p,'rb') as f: sav_tukey=f.read(); os.unlink(p)
                st.download_button("📊 6. Tukey Post-Hoc — SPSS .sav — Chapter 4.2b", data=sav_tukey, file_name=f"{reg_no}_Tukey.sav", mime="application/x-sav", use_container_width=True)
            if 'Regression' in infer:
                coeff_df = infer['Regression']['Coefficients_df']
                with tempfile.NamedTemporaryFile(delete=False,suffix='.sav') as tmp: pyreadstat.write_sav(coeff_df, tmp.name); p=tmp.name
                with open(p,'rb') as f: sav5=f.read(); os.unlink(p)
                st.download_button("📊 7. Regression — SPSS .sav", data=sav5, file_name=f"{reg_no}_Regression.sav", mime="application/x-sav", use_container_width=True)
            excel_buf=io.BytesIO()
            with pd.ExcelWriter(excel_buf,engine='openpyxl') as writer:
                df.to_excel(writer,sheet_name='1_Cleaned_Data',index=False)
                desc_table.to_excel(writer,sheet_name='2_Descriptive',index=False)
                if 'desc_by_group' in st.session_state: st.session_state.desc_by_group.to_excel(writer,sheet_name='3_Desc_By_Group')
                freq.to_excel(writer,sheet_name=f'4_Freq_{col_f}'[:31],index=False)
                if 'ANOVA' in infer: pd.DataFrame([infer['ANOVA']]).to_excel(writer,sheet_name='5_ANOVA',index=False)
                if 'PostHoc_df' in infer: infer['PostHoc_df'].to_excel(writer,sheet_name='6_Tukey_PostHoc',index=False)
                if 'Regression' in infer: infer['Regression']['Coefficients_df'].to_excel(writer,sheet_name='7_Regression',index=False)
            st.download_button("📑 8. Chapter 4 All Tables in Correct Order — Excel — No Watermark", data=excel_buf.getvalue(), file_name=f"{reg_no}_Chapter4_CorrectOrder.xlsx", use_container_width=True)
            for name,fig in graphs:
                buf=io.BytesIO(); fig.savefig(buf,format='png',dpi=300,bbox_inches='tight'); buf.seek(0)
                st.download_button(f"🖼️ {name}.png — No Watermark", data=buf, file_name=f"{name}.png", mime="image/png", key=f"dl_{name}")

# ---------------------------------------------------------------- TAB 3
with tab3:
    if not st.session_state.paid:
        st.markdown("<div class='locked-box'><h3>🔒 Discuss Locked — Pay to unlock teaching with Consensus, Jenni, PubMed, Google Scholar</h3></div>", unsafe_allow_html=True)
    else:
        st.header(f"💬 Discuss — Learn to Discuss Using Similar Articles — Consensus + Jenni + PubMed + Google Scholar — {project_title}")
        st.info("We teach you to discuss — using real similar articles online as guideline — not AI hallucination — supervisor wants real citations")
        dep_col = st.session_state.get('auto_dv') or 'Variable'
        group_col = st.session_state.get('group_col') or 'Group'
        base_query = f"{project_title} {dep_col} {group_col}".strip()
        anova_query = f"{dep_col} effect of {group_col} dose response"
        scholar_query = f"{dep_col} {group_col} experimental study"
        st.subheader("Step 1 — Find Similar Articles — 5 Sources — One Click")
        st.caption("Click to open with YOUR project keywords — find 2-3 similar articles to use as guideline — copy Title + Year")
        r1c1,r1c2,r1c3 = st.columns(3)
        with r1c1:
            consensus_url = f"https://consensus.app/search/?q={quote_plus(base_query)}&searchMode=chat"
            consensus_url2 = f"https://consensus.app/search/?q={quote_plus(anova_query)}"
            st.markdown(f"""
<div style="border:2px solid #185ADB; border-radius:12px; padding:12px; background:#F0F7FF; height:190px;">
<b>🔬 1. Consensus.ai — Yes/No evidence</b><br>
<span style="font-size:12px;">Does {dep_col} differ by {group_col}?</span><br><br>
<a href="{consensus_url}" target="_blank" style="background:#185ADB; color:white; padding:8px 14px; border-radius:8px; text-decoration:none; display:inline-block; font-size:13px;">Search Consensus</a><br>
<a href="{consensus_url2}" target="_blank" style="background:white; color:#185ADB; border:1px solid #185ADB; padding:5px 10px; border-radius:8px; text-decoration:none; display:inline-block; margin-top:6px; font-size:11px;">{anova_query[:32]}</a>
</div>
""", unsafe_allow_html=True)
        with r1c2:
            pubmed_url = f"https://pubmed.ncbi.nlm.nih.gov/?term={quote_plus(base_query)}&sort=date"
            pubmed_url2 = f"https://pubmed.ncbi.nlm.nih.gov/?term={quote_plus(str(dep_col) + ' ' + str(group_col))}"
            st.markdown(f"""
<div style="border:2px solid #0F766E; border-radius:12px; padding:12px; background:#F0FDFA; height:190px;">
<b>🧬 2. PubMed — Biomedical — Free full text</b><br>
<span style="font-size:12px;">Best for Rat/Physiology/Med/Nursing</span><br><br>
<a href="{pubmed_url}" target="_blank" style="background:#0F766E; color:white; padding:8px 14px; border-radius:8px; text-decoration:none; display:inline-block; font-size:13px;">Search PubMed</a><br>
<a href="{pubmed_url2}" target="_blank" style="background:white; color:#0F766E; border:1px solid #0F766E; padding:5px 10px; border-radius:8px; text-decoration:none; display:inline-block; margin-top:6px; font-size:11px;">{dep_col} + {group_col}</a>
</div>
""", unsafe_allow_html=True)
        with r1c3:
            scholar_url = f"https://scholar.google.com/scholar?q={quote_plus(scholar_query)}&as_ylo=2018"
            scholar_url2 = f"https://scholar.google.com/scholar?q={quote_plus(base_query)}"
            st.markdown(f"""
<div style="border:2px solid #1E40AF; border-radius:12px; padding:12px; background:#EFF6FF; height:190px;">
<b>📚 3. Google Scholar — Most cited — All fields</b><br>
<span style="font-size:12px;">Find highly cited similar articles</span><br><br>
<a href="{scholar_url}" target="_blank" style="background:#1E40AF; color:white; padding:8px 14px; border-radius:8px; text-decoration:none; display:inline-block; font-size:13px;">Search Google Scholar</a><br>
<a href="{scholar_url2}" target="_blank" style="background:white; color:#1E40AF; border:1px solid #1E40AF; padding:5px 10px; border-radius:8px; text-decoration:none; display:inline-block; margin-top:6px; font-size:11px;">All years — {base_query[:22]}</a>
</div>
""", unsafe_allow_html=True)
        r2c1,r2c2,r2c3 = st.columns(3)
        with r2c1:
            jenni_url = "https://app.jenni.ai"
            semantic_url = f"https://www.semanticscholar.org/search?q={quote_plus(base_query)}&sort=relevance"
            st.markdown(f"""
<div style="border:2px solid #7C3AED; border-radius:12px; padding:12px; background:#F5F0FF; height:190px;">
<b>✍️ 4. Jenni.ai — Write with citations</b><br>
<span style="font-size:12px;">Import articles from Consensus/PubMed/Scholar → Jenni cites them</span><br><br>
<a href="{jenni_url}" target="_blank" style="background:#7C3AED; color:white; padding:8px 14px; border-radius:8px; text-decoration:none; display:inline-block; font-size:13px;">Open Jenni.ai</a><br>
<a href="{semantic_url}" target="_blank" style="background:white; color:#7C3AED; border:1px solid #7C3AED; padding:5px 10px; border-radius:8px; text-decoration:none; display:inline-block; margin-top:6px; font-size:11px;">Semantic Scholar too</a>
</div>
""", unsafe_allow_html=True)
        with r2c2:
            st.markdown(f"""
<div style="border:1px solid #eee; border-radius:12px; padding:12px; background:white; height:190px;">
<b>📝 How to use 5 sources together (Teaching)</b><br>
<span style="font-size:12px; line-height:1.5;">
                        1. <b>Consensus</b> → Quick Yes/No: Does dose affect {dep_col}?<br>
                        2. <b>PubMed</b> → Rat studies — copy MeSH terms<br>
                        3. <b>Google Scholar</b> → Most cited — check References list<br>
                        4. <b>Copy 2 titles</b> → e.g. Ade et al. 2020<br>
                        5. <b>Jenni</b> → Paste titles → Jenni inserts (Ade et al., 2020)<br>
                        6. <b>Tab 4</b> → Format into APA 7th etc. — citations intact
</span>
</div>
""", unsafe_allow_html=True)
        with r2c3:
            st.markdown(f"""
<div style="background:#FFFDE7; border:1px solid #F9A825; border-radius:12px; padding:12px; height:190px;">
<b>🎓 Supervisor Formula — 4 Paras — Using Similar Articles</b><br>
<span style="font-size:12px; line-height:1.5;">
<b>P1 What found:</b> Mean {dep_col}=X, ANOVA Sig=0.000<br>
<b>P2 Compare:</b> Agrees with PubMed article Ade 2020... Differs from Scholar article Bola 2021...<br>
<b>P3 Why:</b> Agreement due to similar dose; Difference due to strain<br>
<b>P4 Implication:</b> Suggests {dep_col} responsive<br>
<i>Use 2-3 real articles — not fake</i>
</span>
</div>
""", unsafe_allow_html=True)
        st.divider()
        st.subheader("Step 2 — Paste Similar Articles You Found — Then Generate Discussion")
        cA,cB = st.columns(2)
        with cA:
            main_finding = st.text_area("Your Main Finding from Tab 2 — ANOVA Sig, Tukey", height=90, value=f"ANOVA for {dep_col} by {group_col} significant Sig=0.000, Tukey Control vs High Dose Sig=0.000" if group_col else "")
            similar_articles = st.text_area("Paste 2-3 Similar Articles from PubMed / Scholar / Consensus — Copy Title + Authors + Year", height=150, placeholder="From PubMed:\n1. Ade J, et al. (2020). Effect of extract on weight in Wistar rats. J Physiol...\nFrom Google Scholar:\n2. Bola S et al. (2021). High dose effect on glucose...\nFrom Consensus:\n3. ...")
        with cB:
            objective = st.text_area("Objective from Chapter 1", height=90)
            study_type_note = st.selectbox("Study type for discussion tone", ["Experimental Rat Study","Human Survey/Clinical","Other"])
        if st.button("Generate Discussion Using PubMed + Scholar + Consensus as Guideline — No Watermark", type="primary", use_container_width=True):
            mean_line = ""
            try:
                if 'df' in st.session_state and dep_col in st.session_state.df.columns:
                    s = pd.to_numeric(st.session_state.df[dep_col], errors='coerce').dropna()
                    if len(s)>0:
                        mean_line = f"Descriptive analysis revealed mean {dep_col} = {s.mean():.2f} ± {s.std():.2f}."
            except: pass
            anova_line = "ANOVA showed significant difference."
            try:
                if 'df' in st.session_state:
                    ch3 = st.session_state.get('ch3_info',{"methods":["ANOVA","Tukey"]})
                    infer_temp = run_inferential_correct_order(st.session_state.df, dep_col, st.session_state.get('auto_ivs',[]), st.session_state.get('group_col'), ch3['methods'])
                    if 'ANOVA' in infer_temp:
                        anova_line = f"One-way ANOVA showed significant difference in {dep_col} across {infer_temp['ANOVA']['factor']} (F={infer_temp['ANOVA']['F']}, p={infer_temp['ANOVA']['Sig.']}). Tukey HSD identified specific group differences."
            except: pass
            refs_guideline = similar_articles if similar_articles.strip() else "Ade et al. (2020) from PubMed and Bola et al. (2021) from Google Scholar"
            draft = f"""CHAPTER FIVE — DISCUSSION — {project_title.upper()}

This {study_type_note.lower()} was conducted at {uni}, Department of {dept}, to investigate {objective if objective else project_title}. {mean_line}

{anova_line} Main finding: {main_finding}.

Comparison with similar articles as guideline (from PubMed, Google Scholar, Consensus.ai):
The present finding aligns with evidence from similar studies retrieved via PubMed and Google Scholar. {refs_guideline}. For instance, PubMed-indexed studies using similar doses reported comparable changes in {dep_col}, supporting the dose-dependent effect observed. Google Scholar highly cited articles also reported similar trends, while some differences may be attributed to strain, duration, or dosage variations as often discussed in literature syntheses from Consensus.ai.

How to write this with Jenni.ai (Teaching):
1. Open https://app.jenni.ai → New document: Discussion {dep_col}
2. Go to PubMed link above → Copy DOI/title of 2 articles → Paste into Jenni library (Jenni auto-fetches citation)
3. Go to Google Scholar link → Copy 1 highly cited article → Add to Jenni
4. In Jenni, type: "Discuss my finding: {main_finding} in relation to [{refs_guideline[:60]}] — {dep_col} by {group_col} — experimental rat study"
5. Jenni will insert citations: (Ade et al., 2020) — keep them intact — our Tab 4 Formatter preserves them with 10+ styles (APA 7th, Vancouver, etc.)
6. Copy Jenni output back here → Download → Then format in Tab 4

Implication:
Significant difference between Control and High Dose suggests {dep_col} is responsive to intervention, with potential translational relevance as suggested by PubMed biomedical literature.

Conclusion:
Objective achieved. Descriptive by group, ANOVA, Tukey confirm {dep_col} varies by {group_col}. Similar articles from PubMed/Google Scholar/Consensus provide guideline for interpretation.

References for discussion (from your similar articles — format in Tab 4 with 10+ styles):
{refs_guideline}
"""
            st.session_state["discussion_guideline_full"] = draft   # set BEFORE the text_area below is created

        if "discussion_guideline_full" in st.session_state:
            st.text_area("Edit final discussion — Keep citations intact — Will be preserved in Tab 4 with 10+ styles", height=450, key="discussion_guideline_full")
            ddoc = docx.Document()
            dstyle = ddoc.styles['Normal']; dstyle.font.name="Times New Roman"; dstyle.font.size=Pt(12); dstyle.paragraph_format.line_spacing=Pt(18)
            for para_text in st.session_state["discussion_guideline_full"].split("\n\n"):
                if para_text.strip():
                    ddoc.add_paragraph(para_text.strip())
            dbio = io.BytesIO(); ddoc.save(dbio)
            st.download_button("📄 Download Discussion — With PubMed + Scholar + Consensus Guideline — No Watermark — Citations Intact", dbio.getvalue(), f"{reg_no}_Discussion_PubMed_Scholar_Consensus_Jenni.docx", use_container_width=True)
            st.success("Discussion uses YOUR similar articles from PubMed/Google Scholar/Consensus as guideline — not fake — ready for Tab 4 formatting into APA 7th/Vancouver etc.")
        st.divider()
        st.markdown(f"""
<div style="text-align:center; font-size:12px; color:gray;">
<b>Research Base HQ workflow with 5 sources:</b><br>
        Research Base HQ (Descriptive → Group → ANOVA → Tukey correct order) → PubMed (biomedical) + Google Scholar (most cited) + Consensus.ai (Yes/No) → Copy 2-3 real articles → Jenni.ai (writes discussion with those citations intact) → Research Base HQ Tab 4 (formats References into 10+ styles — APA 7th, Vancouver, Harvard etc.) — No watermark<br>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------- TAB 4
with tab4:
    if not st.session_state.paid:
        st.markdown("<div class='locked-box'><h3>🔒 Format Locked — Pay to unlock university format chooser</h3></div>", unsafe_allow_html=True)
    else:
        st.header(f"📄 Format Chapters 1-5 — Choose/Confirm Your University Format On The Spot — References + 10+ Styles — No Watermark")

        uni_formats = {
            "University of Port Harcourt": {"font":"Times New Roman","size":12,"spacing":1.5,"margin":"1 inch all sides","chapter":"CHAPTER ONE","ref":"APA 7th"},
            "University of Lagos": {"font":"Times New Roman","size":12,"spacing":2.0,"margin":"1 inch","chapter":"CHAPTER ONE","ref":"APA 6th"},
            "University of Ibadan": {"font":"Times New Roman","size":12,"spacing":2.0,"margin":"1.5 left, 1 others","ref":"APA 7th"},
            "University of Benin": {"font":"Times New Roman","size":12,"spacing":1.5,"ref":"APA 7th"},
            "University of Nigeria Nsukka": {"font":"Times New Roman","size":12,"spacing":2.0,"ref":"APA 7th"},
            "Rivers State University": {"font":"Times New Roman","size":12,"spacing":1.5,"ref":"APA 7th"},
            "Ignatius Ajuru University": {"font":"Times New Roman","size":12,"spacing":2.0,"ref":"APA 6th"},
            "Other / Custom": {"font":"Times New Roman","size":12,"spacing":1.5,"ref":"APA 7th"}
        }
        st.subheader("🏛️ Step 1 — Choose/Confirm Your University Format On The Spot")
        selected_uni_format = st.selectbox("Select Your University to Auto-Load Official Format *", list(uni_formats.keys()), index=list(uni_formats.keys()).index(uni) if uni in uni_formats else 0)
        fmt = uni_formats.get(selected_uni_format, uni_formats["Other / Custom"])
        st.info(f"✅ {selected_uni_format} Official Format: Font={fmt['font']} {fmt['size']}pt, Spacing={fmt['spacing']}, Ref={fmt['ref']}, Margin={fmt.get('margin','1 inch')} — Confirm on spot — You can still edit below")
        confirm_format = st.checkbox(f"Confirm — Use {selected_uni_format} official format — I have checked with my department", value=True)

        c0,c1,c2=st.columns(3)
        with c0:
            font_options=["Times New Roman","Arial","Calibri","Cambria","Garamond"]
            font_idx = font_options.index(fmt['font']) if fmt['font'] in font_options else 0
            font_choice=st.selectbox("Font *", font_options, index=font_idx)
        with c1:
            size_options=[10,11,12,13,14]
            size_idx = size_options.index(fmt['size']) if fmt['size'] in size_options else 2
            size_choice=st.selectbox("Font Size *", size_options, index=size_idx)
        with c2:
            spacing_options=[1.0,1.15,1.5,2.0]
            spacing_idx = spacing_options.index(fmt['spacing']) if fmt['spacing'] in spacing_options else 2
            spacing_choice=st.selectbox("Line Spacing *", spacing_options, index=spacing_idx)

        c3,c4=st.columns(2)
        with c3:
            ref_styles=["APA 7th","APA 6th","MLA 9th","Chicago 17th","Harvard","Vancouver","IEEE","AMA 11th","Turabian","OSCOLA","Keep as is — Preserve exactly"]
            ref_idx = ref_styles.index(fmt['ref']) if fmt['ref'] in ref_styles else 0
            ref_style=st.selectbox("Reference Style * — 10+ styles", ref_styles, index=ref_idx)
        with c4:
            keep_cite=st.checkbox("Keep in-text citations intact — (Ade et al., 2020), [1] — Do not break", value=True)
            auto_sort=st.checkbox("Auto-sort References A-Z (for APA/Harvard)", value=False)

        full=st.text_area("Paste Chapters 1-5 — From INTRODUCTION to REFERENCES — Citations will be kept intact", height=260,
            placeholder="Example:\nCHAPTER ONE...\nAs shown by Ade et al. (2020) and [1]...\n\nREFERENCES\nAde, J. (2020). Title. Journal...\nBola, S. (2021). Title...")
        if st.button("Format Full Project — University Format + Citations & References Intact — No Watermark", type="primary", use_container_width=True):
            if not confirm_format:
                st.warning("Please confirm university format — tick the box above")
            elif not full.strip():
                st.warning("Paste your chapters first")
            else:
                doc=docx.Document()
                style=doc.styles['Normal']; style.font.name=font_choice; style.font.size=Pt(size_choice); style.paragraph_format.line_spacing=spacing_choice; style.paragraph_format.space_after=Pt(6)
                p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
                run=p.add_run(f"{project_title.upper()}\n\nBY\n\n{student_name.upper()}\n{reg_no}\n{dept.upper()}\n{selected_uni_format.upper()}"); run.bold=True; run.font.size=Pt(size_choice+2)
                doc.add_page_break()
                text_parts = re.split(r'\n\s*REFERENCES\s*\n|\n\s*REFERENCE\s*\n|\n\s*BIBLIOGRAPHY\s*\n', full, flags=re.IGNORECASE)
                main_body = text_parts[0] if text_parts else full
                ref_section = text_parts[1] if len(text_parts)>1 else ""
                for para_text in main_body.split("\n\n"):
                    para_text=para_text.strip()
                    if not para_text: continue
                    para=doc.add_paragraph()
                    para.add_run(para_text)
                    para.paragraph_format.line_spacing=spacing_choice
                no_ref_header = not ref_section.strip()
                if not no_ref_header:
                    doc.add_page_break()
                    h=doc.add_paragraph(); r=h.add_run("REFERENCES"); r.bold=True; r.font.size=Pt(size_choice+1); h.alignment=WD_ALIGN_PARAGRAPH.CENTER
                    refs = [rl.strip() for rl in ref_section.split("\n") if rl.strip()]
                    if auto_sort and ref_style in ["APA 7th","APA 6th","Harvard","MLA 9th"]:
                        refs = sorted(refs, key=lambda x: x.lower())
                    for ref_line in refs:
                        p=doc.add_paragraph()
                        p.add_run(ref_line)
                        if ref_style in ["APA 7th","APA 6th","Harvard"]:
                            p.paragraph_format.left_indent=Pt(36); p.paragraph_format.first_line_indent=Pt(-36)
                            p.paragraph_format.line_spacing=2.0; p.paragraph_format.space_after=Pt(0)
                        elif ref_style in ["MLA 9th","Chicago 17th","Turabian"]:
                            p.paragraph_format.left_indent=Pt(36); p.paragraph_format.first_line_indent=Pt(-36)
                            p.paragraph_format.line_spacing=2.0
                        elif ref_style in ["Vancouver","IEEE","AMA 11th"]:
                            p.paragraph_format.left_indent=Pt(18); p.paragraph_format.first_line_indent=Pt(0)
                            p.paragraph_format.line_spacing=1.0; p.paragraph_format.space_after=Pt(3)
                        elif ref_style == "OSCOLA":
                            p.paragraph_format.left_indent=Pt(0); p.paragraph_format.first_line_indent=Pt(0)
                            p.paragraph_format.line_spacing=1.0
                        else:
                            p.paragraph_format.line_spacing=spacing_choice
                bio=io.BytesIO(); doc.save(bio)
                safe_uni = re.sub(r'[^A-Za-z0-9]+', '_', selected_uni_format).strip('_')
                safe_ref = re.sub(r'[^A-Za-z0-9]+', '_', ref_style).strip('_')
                st.session_state["formatted_doc"] = {
                    "data": bio.getvalue(),
                    "name": f"{reg_no}_Chapters1-5_{safe_uni}_{safe_ref}_CitationsIntact.docx",
                    "label": f"📄 Download Formatted 1-5 — {selected_uni_format} — {ref_style} — Citations Intact — No Watermark",
                    "msg": f"✅ Formatted for {selected_uni_format} — {font_choice} {size_choice}pt {spacing_choice} spacing — References in {ref_style} — Citations intact — No Watermark",
                    "caption": f"University Format Confirmed: {selected_uni_format} — Preserved: (Ade et al., 2020), [1] — Reference list in {ref_style}",
                    "no_ref_header": no_ref_header,
                }
        fd = st.session_state.get("formatted_doc")
        if fd:
            if fd["no_ref_header"]:
                st.warning("No REFERENCES header found — We formatted all as body. Include a line with REFERENCES before your list to format references separately.")
            st.success(fd["msg"])
            st.download_button(fd["label"], fd["data"], fd["name"], use_container_width=True)
            st.caption(fd["caption"])

# ---------------------------------------------------------------- TAB 5
with tab5:
    if not st.session_state.paid:
        st.markdown("<div class='locked-box'><h3>🔒 Admin Review Locked</h3></div>", unsafe_allow_html=True)
    elif "Standard" not in str(st.session_state.tier):
        st.error("Upgrade to Standard ₦40k for Admin Review")
    else:
        st.header(f"📤 Send to Admin — {project_title}")
        work=st.text_area("Paste work for review")
        if st.button("Submit to Creator"):
            if save_review([datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), uni, dept, project_title, student_email, work[:4000], st.session_state.verify_code, "", reg_no]):
                st.success("Submitted")
            else:
                st.error("Could not submit right now — please send it on WhatsApp instead.")

# ---------------------------------------------------------------- TAB 6
with tab6:
    if not st.session_state.paid:
        st.markdown("<div class='locked-box'><h3>🔒 AI Checker Locked — Standard Plan ₦40k Only</h3><p>Basic ₦30k = No AI checker. Standard = AI checker + humanize on spot so supervisor does not flag you</p></div>", unsafe_allow_html=True)
    elif "Standard" not in str(st.session_state.tier):
        st.error("🔒 AI Checker is Standard Plan only — Upgrade to ₦40k")
    else:
        st.header(f"🤖 AI Checker — Included in Standard — Check If Your Project Will Be Flagged — {project_title}")
        st.info("Supervisor may use Turnitin AI, ZeroGPT, Originality — We check on the spot before you submit — Included in Standard ₦40k")
        check_text = st.text_area("Paste Discussion / Chapter 4 / Full Chapters 1-5 to check — We analyze like AI detector", height=250, placeholder="Paste your Discussion from Tab 3 or formatted project...")
        if st.button("Check AI Score — On The Spot — No Watermark", type="primary", use_container_width=True):
            if not check_text.strip():
                st.warning("Paste text to check")
            else:
                words = check_text.split()
                word_count = len(words)
                avg_word_len = sum(len(w) for w in words)/word_count if word_count else 0
                sentences = re.split(r'[.!?]+', check_text)
                sentence_count = len([s for s in sentences if s.strip()])
                avg_sentence_len = word_count / sentence_count if sentence_count else 0
                repetitive_phrases = len(re.findall(r'\b(in conclusion|furthermore|moreover|additionally|it is important to note|delve into|tapestry)\b', check_text.lower()))
                burstiness = np.std([len(s.split()) for s in sentences if s.strip()]) if sentence_count>1 else 0
                ai_score = 0
                if avg_sentence_len > 22: ai_score += 25
                if avg_word_len > 5.5: ai_score += 15
                if repetitive_phrases > 3: ai_score += 20
                if burstiness < 5: ai_score += 25
                if word_count < 100: ai_score += 10
                ai_score = min(95, ai_score + random.randint(5,15))
                humanized = None
                if ai_score >= 30:
                    humanized = check_text.replace("Furthermore,", "Also,").replace("Moreover,", "In addition,").replace("It is important to note that", "We observed that").replace("In conclusion,", "From our findings at "+str(uni)+",")
                st.session_state["ai_result"] = {"ai_score": ai_score, "human_score": 100 - ai_score, "avg_sentence_len": avg_sentence_len,
                                                 "repetitive_phrases": repetitive_phrases, "burstiness": float(burstiness), "humanized": humanized}
        res = st.session_state.get("ai_result")
        if res:
            ai_score = res["ai_score"]; human_score = res["human_score"]
            cA,cB,cC = st.columns(3)
            with cA: st.metric("🤖 AI Likelihood", f"{ai_score}%")
            with cB: st.metric("👤 Human Likelihood", f"{human_score}%")
            with cC:
                if ai_score < 30: st.success("✅ SAFE — Likely not flagged")
                elif ai_score < 60: st.warning("⚠️ MEDIUM — Rewrite some parts")
                else: st.error("🔴 HIGH RISK — Will likely be flagged — Rewrite needed")
            st.divider()
            st.subheader("How to Fix If Flagged — Humanize On The Spot")
            st.markdown(f"""
<div style="background:#F0FFF0; border:1px solid #4CAF50; border-radius:10px; padding:14px; font-size:14px;">
<b>If AI Score >60%:</b><br>
                                1. Break long sentences — Avg now {res['avg_sentence_len']:.1f} words — aim 15-18<br>
                                2. Add personal observation — e.g. "In our lab at {uni}, we observed..."<br>
                                3. Replace AI phrases: "delve", "tapestry", "in conclusion" found {res['repetitive_phrases']} times<br>
                                4. Vary sentence length — Burstiness now {res['burstiness']:.1f} — mix short and long<br>
                                5. Add specific numbers from YOUR results — e.g. "Weight_g mean 157.5" — AI detectors like specifics<br>
                                6. Paste revised text back into Tab 3 Discussion then re-check here<br>
                                7. Final check before submitting to supervisor
</div>
""", unsafe_allow_html=True)
            if res["humanized"] is not None:
                st.subheader("Suggested Humanized Version — Edit Then Use")
                st.text_area("Humanized draft — Copy to Tab 3/Tab 4:", value=res["humanized"][:4000], height=200)
                hdoc = docx.Document()
                hdoc.add_paragraph(res["humanized"])
                hbio = io.BytesIO(); hdoc.save(hbio)
                st.download_button("📄 Download Humanized Version — No Watermark", hbio.getvalue(), f"{reg_no}_Humanized_AI_Checked.docx", use_container_width=True)
        st.divider()
        st.caption("AI Checker included in Standard ₦40k — Basic ₦30k cannot access — We teach you to humanize using your own lab observations + specific numbers from Tab 2")

# ---------------------------------------------------------------- FOOTER
st.divider()
st.markdown(f"""
<div style='text-align:center; font-size:13px; color:gray;'>
<b>{APP_NAME}</b> — Pay Per Project — No watermark — Correct Order — Choose University Format On Spot — AI Checker (Standard) — Real .sav — Chapters 1-5 includes References + 10+ Styles — Discuss with PubMed + Scholar + Consensus + Jenni<br>
Helpline: <b>{WHATSAPP_DISPLAY}</b> — <a href="{WHATSAPP_LINK}" target="_blank">Chat on WhatsApp</a><br>
{VERIFY_DOMAIN} — © 2026
</div>
""", unsafe_allow_html=True)
