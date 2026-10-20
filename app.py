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
import uuid, hashlib, smtplib, ssl
from email.message import EmailMessage
from sqlalchemy import create_engine, text as sql_text

st.set_page_config(page_title="Research Base HQ | Research Project Support", page_icon="📘", layout="wide", initial_sidebar_state="collapsed")

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
PRICE_TIERS = {"Basic - ₦30,000 (Analyse + Format)": {"amount": 3000000, "display": 30000}, "Standard - ₦40,000 (Most Popular + Writing Quality Review)": {"amount": 4000000, "display": 40000}}


# ---------------------------------------------------------------- PERSISTENT FUNNEL / LEAD / PROJECT EVENT STORE
# Set DATABASE_URL to a managed PostgreSQL URL in production. Local SQLite is for development only.
_DATABASE_URL = _secret("database_url") or os.getenv("DATABASE_URL") or "sqlite:///research_base_hq.db"

@st.cache_resource(show_spinner=False)
def _db_engine():
    return create_engine(_DATABASE_URL, pool_pre_ping=True, future=True)

def init_app_db():
    statements = [
        "CREATE TABLE IF NOT EXISTS leads (id VARCHAR(36) PRIMARY KEY, email VARCHAR(320) NOT NULL, source VARCHAR(120), campaign VARCHAR(120), consent BOOLEAN NOT NULL DEFAULT FALSE, created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP)",
        "CREATE TABLE IF NOT EXISTS funnel_events (id VARCHAR(36) PRIMARY KEY, event_name VARCHAR(100) NOT NULL, source VARCHAR(120), campaign VARCHAR(120), email_hash VARCHAR(64), project_key_hash VARCHAR(64), metadata TEXT, created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP)",
        "CREATE TABLE IF NOT EXISTS project_progress (project_key_hash VARCHAR(64) PRIMARY KEY, stage VARCHAR(80) NOT NULL DEFAULT 'project_started', source VARCHAR(120), campaign VARCHAR(120), updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP)"
    ]
    with _db_engine().begin() as conn:
        for statement in statements:
            conn.execute(sql_text(statement))

def _hash_value(value):
    value = str(value or '').strip().lower()
    return hashlib.sha256(value.encode('utf-8')).hexdigest() if value else None

def track_event(event_name, email=None, project_key=None, metadata=None):
    try:
        init_app_db()
        conn_source = str(st.session_state.get('utm_source', 'direct'))[:120]
        conn_campaign = str(st.session_state.get('utm_campaign', 'none'))[:120]
        with _db_engine().begin() as conn:
            conn.execute(sql_text("INSERT INTO funnel_events (id,event_name,source,campaign,email_hash,project_key_hash,metadata) VALUES (:id,:event_name,:source,:campaign,:email_hash,:project_key_hash,:metadata)"), {
                'id': str(uuid.uuid4()), 'event_name': str(event_name)[:100], 'source': conn_source,
                'campaign': conn_campaign, 'email_hash': _hash_value(email),
                'project_key_hash': _hash_value(project_key), 'metadata': str(metadata or '')[:1000]
            })
            if project_key and event_name in {'project_started','dataset_uploaded','analysis_completed','discussion_opened','formatting_opened','report_downloaded'}:
                conn.execute(sql_text("INSERT INTO project_progress (project_key_hash,stage,source,campaign,updated_at) VALUES (:key,:stage,:source,:campaign,CURRENT_TIMESTAMP) ON CONFLICT(project_key_hash) DO UPDATE SET stage=CASE WHEN CASE excluded.stage WHEN 'project_started' THEN 0 WHEN 'dataset_uploaded' THEN 1 WHEN 'analysis_completed' THEN 2 WHEN 'discussion_opened' THEN 3 WHEN 'formatting_opened' THEN 4 WHEN 'report_downloaded' THEN 5 ELSE 0 END >= CASE project_progress.stage WHEN 'project_started' THEN 0 WHEN 'dataset_uploaded' THEN 1 WHEN 'analysis_completed' THEN 2 WHEN 'discussion_opened' THEN 3 WHEN 'formatting_opened' THEN 4 WHEN 'report_downloaded' THEN 5 ELSE 0 END THEN excluded.stage ELSE project_progress.stage END, source=excluded.source, campaign=excluded.campaign, updated_at=CURRENT_TIMESTAMP"), {'key':_hash_value(project_key),'stage':str(event_name)[:80],'source':conn_source,'campaign':conn_campaign})
    except Exception as exc:
        # Analytics failure must never prevent the research workflow.
        st.session_state['analytics_warning'] = f"Analytics store unavailable: {type(exc).__name__}"

def send_resource_email(email):
    """Send the requested resource by SMTP when configured. No marketing sequence is sent here."""
    try:
        cfg = st.secrets["smtp"]
        host = str(cfg["host"]); port = int(cfg.get("port", 587))
        username = str(cfg["username"]); password = str(cfg["password"])
        from_email = str(cfg["from_email"])
        message = EmailMessage()
        message["Subject"] = "Your Research Base HQ Chapter Four Survival Kit"
        message["From"] = from_email
        message["To"] = email
        message.set_content("""Thank you for requesting the Research Base HQ Chapter Four Survival Kit.

Use the checklist below as a starting point, and verify your analysis against your approved methodology and supervisor's instructions.

BEFORE ANALYSIS
- Confirm objectives and hypotheses.
- Keep an untouched copy of the original dataset.
- Check missing values, duplicates, units and out-of-range values.
- Create a variable map.
- Confirm the statistical method with your approved study design.

REPORTING RESULTS
- Report sample sizes, test statistics and exact p-values where appropriate.
- Use post-hoc tests only when justified.
- Label tables and figures and verify all output.

DISCUSSION
- Address each objective.
- Compare only with real, verified articles.
- Discuss limitations and avoid unsupported claims.

FORMATTING
- Check the latest official departmental guide.
- Verify citations, references and numbering.

You can also download the checklist directly from the Research Base HQ page. This email delivers the resource you requested; promotional updates are sent only when you separately opted in.
""")
        context = ssl.create_default_context()
        with smtplib.SMTP(host, port, timeout=20) as server:
            server.starttls(context=context)
            server.login(username, password)
            server.send_message(message)
        return True
    except Exception:
        return False

def save_lead(email, consent):
    init_app_db()
    with _db_engine().begin() as conn:
        conn.execute(sql_text("INSERT INTO leads (id,email,source,campaign,consent) VALUES (:id,:email,:source,:campaign,:consent)"), {
            'id': str(uuid.uuid4()), 'email': email.strip().lower()[:320],
            'source': str(st.session_state.get('utm_source','direct'))[:120],
            'campaign': str(st.session_state.get('utm_campaign','none'))[:120], 'consent': bool(consent)
        })
    track_event('lead_captured', email=email)
    return send_resource_email(email)

def get_funnel_summary():
    init_app_db()
    with _db_engine().connect() as conn:
        rows = conn.execute(sql_text("SELECT event_name, COUNT(*) AS n FROM funnel_events GROUP BY event_name ORDER BY n DESC")).mappings().all()
        by_source = conn.execute(sql_text("SELECT source, event_name, COUNT(*) AS n FROM funnel_events GROUP BY source, event_name ORDER BY source, event_name")).mappings().all()
        leads = conn.execute(sql_text("SELECT COUNT(*) FROM leads")).scalar_one()
    return int(leads), {r['event_name']: int(r['n']) for r in rows}, [dict(r) for r in by_source]

def get_project_stage(project_key):
    try:
        init_app_db()
        with _db_engine().connect() as conn:
            row = conn.execute(sql_text("SELECT stage FROM project_progress WHERE project_key_hash=:key"), {'key': _hash_value(project_key)}).first()
        return row[0] if row else None
    except Exception:
        return None

def get_marketing_leads_df():
    init_app_db()
    with _db_engine().connect() as conn:
        rows = conn.execute(sql_text("SELECT email, source, campaign, consent, created_at FROM leads ORDER BY created_at DESC")).mappings().all()
    return pd.DataFrame([dict(r) for r in rows])

# Capture attribution on the first page view. Do not store raw research data in analytics.
st.session_state.setdefault('utm_source', str(st.query_params.get('utm_source', 'direct'))[:120])
st.session_state.setdefault('utm_campaign', str(st.query_params.get('utm_campaign', 'none'))[:120])
if not st.session_state.get('_page_view_tracked'):
    track_event('landing_page_view')
    st.session_state['_page_view_tracked'] = True

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
    dv=re.search(r'dependent variable[^.]*?is\s+([a-z ]{3,40})',t)
    return {"methods":m,"dv_text":dv.group(1) if dv else "", "needs_review": not bool(m)}

# ---------------------------------------------------------------- GOOGLE SHEETS
def _clean_private_key(pk):
    """Rebuild a clean PEM from however the key was pasted (missing header, literal \\n, stray spaces/quotes)."""
    pk = str(pk).strip().strip('"').strip("'").replace("\\n", "\n")
    body = pk.replace("-----BEGIN PRIVATE KEY-----", "").replace("-----END PRIVATE KEY-----", "")
    body = "".join(body.split())
    lines = [body[i:i+64] for i in range(0, len(body), 64)]
    return "-----BEGIN PRIVATE KEY-----\n" + "\n".join(lines) + "\n-----END PRIVATE KEY-----\n"

@st.cache_resource(show_spinner=False)
def _connect():
    scope=["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/drive"]
    creds_dict=dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"]=_clean_private_key(creds_dict.get("private_key", ""))
    creds=ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client=gspread.authorize(creds)
    target=str(st.secrets["paysheets"]["sheet_url"]).strip()
    if target.startswith("http"): return client.open_by_url(target)
    return client.open_by_key(target)      # also accepts just the spreadsheet ID

SHEET_ERR = {"msg": ""}

def get_sheet():
    try: return _connect()
    except Exception as e:
        SHEET_ERR["msg"] = f"{type(e).__name__}: {e}"
        return None

def sheet_hint(msg):
    m = msg.lower()
    if "paysheets" in m or "sheet_url" in m: return "The [paysheets] section with sheet_url is missing from your secrets."
    if "gcp_service_account" in m: return "The [gcp_service_account] section is missing from your secrets."
    if "has not been used" in m or "is disabled" in m or "accessnotconfigured" in m: return "Google Sheets API / Google Drive API are not enabled for your Google Cloud project. Enable both, wait a minute, then refresh."
    if "spreadsheetnotfound" in m or "404" in m: return "The sheet was not found. Check sheet_url, and that the sheet is shared with the service account email as Editor."
    if "403" in m or "permission" in m: return "Permission denied. Share the sheet with the service account email as Editor."
    if "invalid_grant" in m or "jwt" in m or "private key" in m or "pem" in m or "asn1" in m or "could not deserialize" in m: return "The private_key in your secrets is wrong or was pasted incorrectly. Paste the NEW key exactly as in the JSON file, including the BEGIN/END lines."
    if "nonvalidurlkey" in m or "invalid" in m and "url" in m: return "sheet_url is not a valid Google Sheets link or ID."
    return "See the message above."

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


# ---- extra helpers: payments table, reviews, sheet-safe text
PAY_COLS = ["date","uni","dept","project_title","verify_code","email","amount_kobo","tier","reg_no","status","reference"]
RCOL = {"date":0,"uni":1,"dept":2,"title":3,"email":4,"work":5,"code":6,"reply":7,"reg":8}

def _safe_cell(v):
    v = str(v)
    return (" " + v) if v[:1] in ("=", "+", "-", "@") else v

def _rcell(r, key):
    i = RCOL[key]
    return str(r[i]).strip() if len(r) > i else ""

@st.cache_data(ttl=20, show_spinner=False)
def review_rows():
    sh = get_sheet()
    if not sh: return None
    try: vals = sh.worksheet("Reviews").get_all_values()
    except Exception: return None
    return [(i, r) for i, r in enumerate(vals, start=1) if i > 1 and any(str(c).strip() for c in r)]

def _access_label(date_str):
    d = _parse_date(date_str)
    if d is None: return "Unknown date"
    left = 180 - (datetime.datetime.now() - d).days
    return f"Active ({left} days left)" if left > 0 else "Expired"

def payments_df():
    rows = payment_rows()
    if rows is None: return None
    data = [(list(r) + [""] * len(PAY_COLS))[:len(PAY_COLS)] for r in rows if any(str(c).strip() for c in r)]
    df = pd.DataFrame(data, columns=PAY_COLS)
    df["amount_naira"] = pd.to_numeric(df["amount_kobo"], errors="coerce").fillna(0) / 100
    df["plan"] = df["tier"].apply(lambda t: "Standard" if "Standard" in str(t) else ("Basic" if "Basic" in str(t) else "Other"))
    df["manual"] = df["reference"].astype(str).str.startswith("manual-")
    df["access"] = df["date"].apply(_access_label)
    return df

def _tier_for(meta_tier, amount):
    if meta_tier in PRICE_TIERS: return meta_tier
    for k, v in PRICE_TIERS.items():
        if v["amount"] == amount: return k
    return list(PRICE_TIERS.keys())[0]

# ---------------------------------------------------------------- PAYSTACK
def _app_url():
    try: return st.secrets["app"]["url"]
    except Exception: return None

def _paystack_key():
    try: k = str(st.secrets["paystack"]["secret_key"])
    except Exception: return None
    k = k.strip().strip('"').strip("'").strip()
    return k if k.startswith(("sk_live_", "sk_test_")) else None

def initialize_paystack(email, amount_kobo, metadata=None):
    try:
        secret=_paystack_key()
        if not secret:
            st.error("Payment is not set up correctly right now. Please contact support on WhatsApp."); return None, None
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
        secret=_paystack_key()
        if not secret: return False, {}
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
    rows_before = len(work_df)
    rows_excluded = int(work_df.isna().any(axis=1).sum())
    work_df=work_df.dropna()
    out['Data_Handling'] = {'Rows in original data': len(df), 'Rows available for selected numeric variables': rows_before, 'Rows excluded by missing/non-numeric selected values': rows_excluded, 'Rows used in complete-case analysis': len(work_df)}
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
    if "ANOVA" in methods:
        gc = group_col
        if gc and dv in df.columns and gc in df.columns:
            try:
                groups=[pd.to_numeric(g[dv], errors='coerce').dropna().values for _,g in df.groupby(gc) if len(pd.to_numeric(g[dv], errors='coerce').dropna())>0]
                if len(groups)>=2 and all(len(g) >= 2 for g in groups):
                    f,p=stats.f_oneway(*groups)
                    lev_stat, lev_p = stats.levene(*groups, center='median')
                    out['ANOVA']={'F':round(float(f),3),'Sig.':round(float(p),6),'factor':gc,
                                  'Levene statistic':round(float(lev_stat),3),'Levene p-value':round(float(lev_p),6),
                                  'group sizes':[int(len(g)) for g in groups]}
                    if p < 0.05 and "Tukey" in methods and lev_p >= 0.05:
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
                    elif p < 0.05 and "Tukey" in methods and lev_p < 0.05:
                        out['PostHoc_Error']='Tukey HSD was not run because the equal-variance check was not met. Ask whether a variance-robust post-hoc procedure is appropriate.'
            except: pass
    if "T-Test" in methods:
        cat = group_col if group_col in df.columns and group_col != dv else None
        if cat and df[cat].nunique(dropna=True) == 2 and dv in df.columns:
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
:root {{ --rbhq-navy:#102A43; --rbhq-blue:#2563EB; --rbhq-teal:#0F766E; --rbhq-ink:#172B4D; --rbhq-muted:#526579; --rbhq-border:#D9E2EC; }}
.block-container {{ max-width: 1180px; padding-top: 1.4rem; padding-bottom: 3rem; }}
.hero {{background: radial-gradient(circle at 90% 10%, rgba(59,130,246,.42), transparent 30%), linear-gradient(125deg, #102A43 0%, #163D70 60%, #1D4ED8 100%); padding: clamp(28px, 5vw, 58px) clamp(20px, 4vw, 54px); border-radius: 24px; color: white; text-align: left; margin: 12px 0 22px; box-shadow: 0 18px 44px rgba(16,42,67,.16); }}
.hero h1 {{ font-size: clamp(2rem, 4vw, 3.2rem); line-height: 1.08; letter-spacing: -0.04em; max-width: 800px; margin-bottom: 16px; }}
.hero p {{ font-size: 1.08rem; line-height: 1.65; color: #E6EEF8; max-width: 760px; }}
.app-name {{font-size:17px; font-weight:850; color:#102A43; letter-spacing:.2px; background:#F0F6FF; border:1px solid #D8E6FA; border-radius:14px; padding:9px 15px; display:inline-block; margin-bottom:8px;}}
.locked-box {{background:#F8FAFC; border:1px solid #CBD5E1; border-radius:14px; padding:22px; text-align:center; margin:15px 0; color:#334155;}}
.price-card {{border: 1px solid #D9E2EC; border-radius: 18px; padding: 24px; background: white; color:#172B4D; min-height:100%; box-shadow:0 4px 16px rgba(16,42,67,.04);}}
.price-card.popular {{border: 2px solid #2563EB; box-shadow: 0 12px 30px rgba(37,99,235,.12);}}
.badge {{background: #DBEAFE; color: #1D4ED8; padding: 5px 10px; border-radius: 20px; font-size: 12px; font-weight: 800; display:inline-block;}}
.section-kicker {{color:#2563EB; font-size:12px; text-transform:uppercase; letter-spacing:.12em; font-weight:800;}}
.feature-card {{border:1px solid #D9E2EC; border-radius:16px; padding:18px; min-height:140px; background:white;}}
.trust-note {{background:#F0FDF4; border:1px solid #BBF7D0; border-radius:14px; padding:14px 16px; color:#166534;}}
.small-muted {{font-size:13px; color:#526579;}}
.whatsapp-float {{position: fixed; width:54px; height:54px; bottom:20px; right:20px; background-color:#16A34A; color:#FFF; border-radius:50%; text-align:center; font-size:26px; box-shadow: 0 6px 20px rgba(0,0,0,0.2); z-index:1000; display:flex; align-items:center; justify-content:center; text-decoration:none;}}
[data-testid="stMetric"] {{background:#F8FAFC; border:1px solid #E2E8F0; padding:12px 14px; border-radius:12px;}}
@media (max-width: 640px) {{ .hero {{padding:26px 20px; border-radius:18px;}} }}
</style>
<a href="{WHATSAPP_LINK}" target="_blank" class="whatsapp-float" aria-label="Contact Research Base HQ on WhatsApp">✆</a>
""", unsafe_allow_html=True)

for k,v in {"is_admin":False,"show_admin_login":False,"paid":False,"verify_code":None,"tier":None,"last_project_paid":""}.items():
    if k not in st.session_state: st.session_state[k]=v

# ---------------------------------------------------------------- ADMIN
if st.query_params.get("admin")=="1" and not st.session_state.is_admin: st.session_state.show_admin_login=True
if st.session_state.show_admin_login and not st.session_state.is_admin:
    st.markdown("### 🔒 Creator Only"); pwd=st.text_input("Creator Password",type="password")
    if st.button("Login as Creator"):
        if ADMIN_PASSWORD and hmac.compare_digest(pwd, str(ADMIN_PASSWORD)): st.session_state.is_admin=True; st.session_state.show_admin_login=False; st.rerun()
        else: st.error("Denied")
    st.stop()   # show ONLY the login box until the correct password is entered

if st.session_state.is_admin:
    h1, h2, h3 = st.columns([4, 1, 1])
    with h1: st.subheader("📊 CREATOR DASHBOARD")
    with h2:
        if st.button("🔄 Refresh data"): payment_rows.clear(); review_rows.clear(); st.rerun()
    with h3:
        if st.button("Logout Creator"): st.session_state.is_admin=False; st.query_params.clear(); st.rerun()
    st.markdown("#### Funnel performance")
    try:
        lead_count, event_counts, source_events = get_funnel_summary()
        fm1, fm2, fm3, fm4 = st.columns(4)
        fm1.metric("Captured leads", lead_count)
        fm2.metric("Landing views", event_counts.get("landing_page_view", 0))
        fm3.metric("Projects started", event_counts.get("project_started", 0))
        fm4.metric("Checkout starts", event_counts.get("checkout_started", 0))
        st.dataframe(pd.DataFrame([{"Event": k, "Count": v} for k, v in sorted(event_counts.items())]), use_container_width=True, hide_index=True)
        if source_events:
            st.markdown("**Events by marketing source**")
            st.dataframe(pd.DataFrame(source_events), use_container_width=True, hide_index=True)
        leads_df = get_marketing_leads_df()
        if not leads_df.empty:
            st.download_button("Download lead list for consent-aware follow-up (CSV)", leads_df.to_csv(index=False).encode("utf-8"), file_name="research_base_hq_leads.csv", mime="text/csv")
    except Exception as exc:
        st.warning(f"Funnel metrics unavailable: {type(exc).__name__}")
    sh = get_sheet()
    if not sh:
        st.error("Cannot connect to the Google Sheet.")
        st.code(SHEET_ERR["msg"] or "No error message captured", language=None)
        st.info("💡 " + sheet_hint(SHEET_ERR["msg"]))
        try: st.write("**The app is signed in as (share your sheet with this email as Editor):**"); st.code(st.secrets["gcp_service_account"]["client_email"], language=None)
        except Exception: st.warning("client_email not found in [gcp_service_account] in your secrets.")
        try:
            tgt = str(st.secrets["paysheets"]["sheet_url"]).strip()
            m_id = re.search(r"/d/([A-Za-z0-9_-]+)", tgt)
            st.write("**Spreadsheet ID the app is trying to open:**"); st.code(m_id.group(1) if m_id else tgt, language=None)
        except Exception: st.warning("[paysheets] sheet_url not found in your secrets.")
        st.stop()
    pdf = payments_df()
    if pdf is None:
        st.error("Could not read the 'Payments' tab. Check the tab name and that the sheet is shared with the service account."); st.stop()
    ok_df = pdf[pdf["status"] == "success"]
    atab1, atab2, atab3, atab4 = st.tabs(["📈 Overview", "🔎 Search payments", "🛠️ Grant / restore access", "📥 Review inbox"])

    # ---------- Overview
    with atab1:
        if ok_df.empty:
            st.info("No successful payments recorded yet.")
        else:
            now = datetime.datetime.now()
            dates = pd.to_datetime(ok_df["date"].astype(str).str.split(" ").str[0], errors="coerce")
            this_month = ok_df[(dates.dt.year == now.year) & (dates.dt.month == now.month)]
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Successful payments", len(ok_df))
            m2.metric("Total revenue", f"₦{ok_df['amount_naira'].sum():,.0f}")
            m3.metric("This month", f"₦{this_month['amount_naira'].sum():,.0f}", f"{len(this_month)} sales", delta_color="off")
            m4.metric("Manual grants", int(ok_df["manual"].sum()))
            st.markdown("#### Basic vs Standard")
            summary = ok_df.groupby("plan").agg(Sales=("plan", "size"), Revenue=("amount_naira", "sum")).reset_index()
            summary["Revenue"] = summary["Revenue"].apply(lambda x: f"₦{x:,.0f}")
            summary.columns = ["Plan", "Sales", "Revenue"]
            st.dataframe(summary, use_container_width=True, hide_index=True)
            monthly = ok_df.assign(month=dates.dt.strftime("%Y-%m")).groupby("month")["amount_naira"].sum()
            if not monthly.empty:
                st.markdown("#### Revenue by month (₦)")
                st.bar_chart(monthly)
            st.markdown("#### Latest 10 payments")
            st.dataframe(ok_df.iloc[::-1].head(10)[["date", "reg_no", "email", "project_title", "plan", "amount_naira", "access"]], use_container_width=True, hide_index=True)

    # ---------- Search
    with atab2:
        q = st.text_input("Search by reg number, email, project title, verification code or reference", key="adm_search").strip().lower()
        view = pdf.iloc[::-1]
        if q:
            cols = ["reg_no", "email", "project_title", "verify_code", "reference", "uni"]
            mask = view[cols].apply(lambda c: c.astype(str).str.lower().str.contains(q, regex=False)).any(axis=1)
            view = view[mask]
        st.caption(f"{len(view)} result(s)")
        show_cols = ["date", "reg_no", "email", "project_title", "uni", "plan", "amount_naira", "status", "access", "verify_code", "reference"]
        st.dataframe(view[show_cols], use_container_width=True, hide_index=True)
        if len(view):
            st.download_button("⬇️ Download these results (CSV)", view[show_cols].to_csv(index=False).encode("utf-8"), "payments.csv", mime="text/csv")

    # ---------- Grant / restore
    with atab3:
        st.markdown("#### 1️⃣ Student paid on Paystack but access did not work")
        st.caption("Paste the Paystack reference (from the student's receipt or your Paystack dashboard). It is confirmed with Paystack before anything is recorded.")
        rec_ref = st.text_input("Paystack reference", key="rec_ref_in").strip()
        if st.button("🔎 Look up reference") and rec_ref:
            ok_p, d_p = verify_paystack(rec_ref)
            st.session_state["rec_lookup"] = {"ref": rec_ref, "ok": ok_p, "data": d_p}
        lk = st.session_state.get("rec_lookup")
        if lk:
            ref = lk["ref"]; d = lk["data"]
            if not lk["ok"]:
                st.error("Paystack does not show this reference as a successful payment.")
            elif reference_used(ref):
                st.info("This reference is already recorded. Search for it in the 'Search payments' tab.")
            else:
                meta = d.get("metadata") if isinstance(d.get("metadata"), dict) else {}
                st.success(f"Paystack confirms a successful payment of ₦{(d.get('amount') or 0) / 100:,.0f} from {(d.get('customer') or {}).get('email', 'unknown email')}.")
                st.caption("Check or complete the details below, then record it.")
                r_reg = st.text_input("Registration number", value=str(meta.get("reg_no", "")), key=f"rec_reg_{ref}").strip()
                r_title = st.text_input("Project title (exactly as the student typed it)", value=str(meta.get("project_title", "")), key=f"rec_title_{ref}").strip()
                r_email = st.text_input("Email", value=(d.get("customer") or {}).get("email", ""), key=f"rec_email_{ref}").strip()
                r_uni = st.text_input("University", value=str(meta.get("uni", "")), key=f"rec_uni_{ref}").strip()
                r_dept = st.text_input("Department", value=str(meta.get("dept", "")), key=f"rec_dept_{ref}").strip()
                tiers = list(PRICE_TIERS.keys())
                r_tier = st.selectbox("Plan", tiers, index=tiers.index(_tier_for(meta.get("tier", ""), d.get("amount"))), key=f"rec_tier_{ref}")
                if st.button("✅ Record payment and grant access", key=f"rec_save_{ref}"):
                    if not (r_reg and r_title and "@" in r_email):
                        st.error("Fill in the registration number, project title and a valid email.")
                    else:
                        paid_at = str(d.get("paid_at") or "")
                        when = f"{paid_at[:10]} {paid_at[11:16]}" if len(paid_at) >= 16 else datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                        vc = ''.join(random.choices(string.ascii_uppercase + string.digits, k=8))
                        if save_payment([when, r_uni, r_dept, r_title, vc, r_email, str(d.get("amount", 0)), r_tier, r_reg, "success", ref]):
                            payment_rows.clear()
                            st.success(f"Recorded. Verification code: {vc}. The student can now enter the same Reg No, Project Title and Email to restore access.")
                        else:
                            st.error("Could not write to the Payments sheet. Check sharing and the tab name.")
        st.divider()
        st.markdown("#### 2️⃣ Grant or renew access manually")
        st.caption("For bank-transfer payments, goodwill access, or renewing an expired project. Enter the amount received so it counts in your revenue (leave 0 for free access).")
        g1, g2 = st.columns(2)
        with g1:
            m_reg = st.text_input("Registration number", key="mg_reg").strip()
            m_email = st.text_input("Student email", key="mg_email").strip()
            m_uni = st.text_input("University", key="mg_uni").strip()
            m_amount = st.number_input("Amount received (₦)", min_value=0, step=1000, value=0, key="mg_amount")
        with g2:
            m_title = st.text_input("Project title (exactly as the student will type it)", key="mg_title").strip()
            m_dept = st.text_input("Department", key="mg_dept").strip()
            tiers = list(PRICE_TIERS.keys())
            m_tier = st.selectbox("Plan", tiers, index=1, key="mg_tier")
            m_note = st.text_input("Note (optional, e.g. 'bank transfer')", key="mg_note").strip()
        if m_reg and m_title and "@" in m_email:
            existing = find_payment(m_reg, m_title, m_email)
            if existing:
                st.warning(f"This student already has active access (code {existing['code']}). Granting again starts a new 180-day period.")
        if st.button("✅ Grant access", key="mg_go"):
            if not (m_reg and len(m_title) > 5 and "@" in m_email):
                st.error("Fill in the registration number, a project title (more than 5 characters) and a valid email.")
            else:
                vc = ''.join(random.choices(string.ascii_uppercase + string.digits, k=8))
                ts = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
                note = re.sub(r'[^A-Za-z0-9]+', '-', m_note).strip('-')[:30]
                ref_val = f"manual-{ts}" + (f"-{note}" if note else "")
                if save_payment([datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), m_uni, m_dept, m_title, vc, m_email, str(int(m_amount) * 100), m_tier, m_reg, "success", ref_val]):
                    payment_rows.clear()
                    st.success(f"Access granted. Verification code: {vc}. The student enters the same Reg No, Project Title and Email on the site and access restores automatically.")
                else:
                    st.error("Could not write to the Payments sheet. Check sharing and the tab name.")

    # ---------- Review inbox
    with atab4:
        rr = review_rows()
        if rr is None:
            st.error("Could not read the 'Reviews' tab. Check the tab name and sharing.")
        else:
            pending = [(n, r) for n, r in rr if not _rcell(r, "reply")]
            answered = [(n, r) for n, r in rr if _rcell(r, "reply")]
            c_a, c_b = st.columns(2)
            c_a.metric("Waiting for reply", len(pending)); c_b.metric("Answered", len(answered))
            which = st.radio("Show", ["Waiting for reply", "Answered", "All"], horizontal=True, key="rv_filter")
            chosen = pending if which == "Waiting for reply" else (answered if which == "Answered" else rr)
            chosen = list(reversed(chosen))[:30]
            if not chosen: st.info("Nothing here.")
            st.caption("Showing the 30 newest.")
            for n, r in chosen:
                icon = "✅" if _rcell(r, "reply") else "🟡"
                with st.expander(f"{icon} {_rcell(r, 'date')} — {_rcell(r, 'reg')} — {_rcell(r, 'title')[:60]}"):
                    st.write(f"**Email:** {_rcell(r, 'email')}  |  **University:** {_rcell(r, 'uni')}  |  **Dept:** {_rcell(r, 'dept')}  |  **Code:** {_rcell(r, 'code')}")
                    st.text_area("Submitted work", value=_rcell(r, "work"), height=200, disabled=True, key=f"rv_work_{n}")
                    reply = st.text_area("Your reply / corrections", value=_rcell(r, "reply"), height=160, key=f"rv_reply_{n}")
                    if st.button("Save reply", key=f"rv_save_{n}"):
                        if not reply.strip():
                            st.error("Write a reply first.")
                        else:
                            try:
                                sh.worksheet("Reviews").update_cell(n, RCOL["reply"] + 1, _safe_cell(reply.strip())[:4000])
                                review_rows.clear()
                                st.success("Reply saved. The student sees it in their Send to Admin tab.")
                            except Exception as e:
                                st.error(f"Could not save the reply: {e}")
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
                # Fail closed: never unlock paid features when the durable payment record failed.
                st.session_state["pay_notice"]=f"Paystack confirms payment, but our record could not be saved. Keep reference {pay_ref} and contact support; access will be restored after reconciliation."
                track_event('payment_record_failed', email=pay_email, project_key=f"{reg}|{title}")
            else:
                track_event('payment_success', email=pay_email, project_key=f"{reg}|{title}", metadata=tier)
                st.session_state.paid=True; st.session_state.verify_code=vc; st.session_state.tier=tier
                st.session_state.last_project_paid=f"{reg.lower()}|{title.lower()}"
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

# ---------------------------------------------------------------- HEADER / CONVERSION LANDING PAGE
st.markdown(f"<div style='text-align:left;'><span class='app-name'>📘 {APP_NAME} <span style='font-weight:500;color:#526579'>· Research workflow workspace</span></span></div>", unsafe_allow_html=True)

if not st.session_state.paid:
    st.markdown(f"""
<div class="hero">
  <div style="font-size:12px;font-weight:800;letter-spacing:.14em;text-transform:uppercase;color:#BFDBFE;margin-bottom:14px;">LESS SPREADSHEET STRESS. MORE CLARITY.</div>
  <h1>Move from raw research data to results you can understand.</h1>
  <p>Research Base HQ helps students organise datasets, explore supported statistical analyses, interpret findings, develop discussion drafts and prepare project documents for review — in one guided workspace.</p>
  <div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:22px;">
    <span style="background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.22);color:white;padding:8px 12px;border-radius:10px;font-size:13px;">✓ Start with a free data template</span>
    <span style="background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.22);color:white;padding:8px 12px;border-radius:10px;font-size:13px;">✓ Preview the workflow before paying</span>
    <span style="background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.22);color:white;padding:8px 12px;border-radius:10px;font-size:13px;">✓ Pay per project</span>
  </div>
  <div style="margin-top:18px;font-size:12px;color:#DCEBFF;">For students and researchers · {VERIFY_DOMAIN}</div>
</div>
""", unsafe_allow_html=True)

    st.markdown("<div class='section-kicker'>A guided workflow</div><h2 style='margin-top:4px'>The key steps of your research project, in one place</h2><p style='color:#526579;max-width:800px'>Start with your own data or explore the sample below. Check the output against your approved methodology and supervisor's guidance.</p>", unsafe_allow_html=True)
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        st.markdown("<div class='feature-card'><div style='font-size:25px'>📂</div><b>1. Organise data</b><p class='small-muted'>Use a structured template and check column names, missing values and duplicates.</p></div>", unsafe_allow_html=True)
    with f2:
        st.markdown("<div class='feature-card'><div style='font-size:25px'>📊</div><b>2. Explore analysis</b><p class='small-muted'>Review descriptive summaries and select a supported test appropriate to your design.</p></div>", unsafe_allow_html=True)
    with f3:
        st.markdown("<div class='feature-card'><div style='font-size:25px'>🧠</div><b>3. Interpret findings</b><p class='small-muted'>Build a results narrative and discussion draft from your actual output.</p></div>", unsafe_allow_html=True)
    with f4:
        st.markdown("<div class='feature-card'><div style='font-size:25px'>📝</div><b>4. Prepare document</b><p class='small-muted'>Organise and format your project, then check it against current departmental rules.</p></div>", unsafe_allow_html=True)
    with st.expander("See an illustrative example of the kind of results table you may produce", expanded=False):
        st.caption("Illustrative values only — not a real study result and not generated from your data.")
        demo = pd.DataFrame({"Group":["Control","Treatment A","Treatment B"],"n":[8,8,8],"Mean score":[10.2,12.4,14.1],"SD":[1.8,2.0,1.6]})
        st.dataframe(demo, use_container_width=True, hide_index=True)
        st.info("Your actual test choice and interpretation depend on your study design, variable types, assumptions and analysis plan. The sample above is only to illustrate output layout.")
    st.markdown("<div class='trust-note'><b>Research integrity matters.</b> Use the output as research support, verify statistical decisions with your approved methods, check every reference, and review the final work with your supervisor. Do not upload identifiable participant data unless you are authorised to do so.</div>", unsafe_allow_html=True)

# ---------------------------------------------------------------- LEAD MAGNET / TOP-OF-FUNNEL
with st.expander("🎁 Free Chapter Four Survival Kit — checklist, template and sample guidance", expanded=not bool(st.session_state.get('paid'))):
    st.write("A practical starting point for students preparing research data and Chapter Four. No purchase is required to request this resource.")
    kit_text = """RESEARCH BASE HQ — CHAPTER FOUR SURVIVAL CHECKLIST

Before analysis:
[ ] Confirm the approved research objectives and hypotheses.
[ ] Keep an untouched copy of the original dataset.
[ ] Use one row per participant/experimental unit unless the design requires repeated measures.
[ ] Use clear, unique column names and document units.
[ ] Check missing values, duplicates, data-entry errors and out-of-range values.
[ ] Create a variable map identifying outcome variables, predictors, groups and categorical variables.
[ ] Confirm the planned statistical method against the approved methodology and study design.

When reporting results:
[ ] Report the sample size and relevant descriptive statistics.
[ ] Identify the test statistic, degrees of freedom where applicable, exact p-value and effect size when appropriate.
[ ] Do not interpret statistical significance as proof of causation.
[ ] Use post-hoc comparisons only when justified by the analysis plan and test results.
[ ] Label every table and figure and define abbreviations.
[ ] Check all results against the raw output and ask your supervisor/statistician when uncertain.

Discussion:
[ ] State the main finding in relation to each objective.
[ ] Compare findings only with real, verified articles.
[ ] Explain plausible reasons for agreement or disagreement without inventing evidence.
[ ] Discuss limitations and avoid claims not supported by the data.

Formatting:
[ ] Check the latest official departmental guide.
[ ] Verify references, citations, table numbering and headings.
[ ] Keep a backup and review the final document before submission.
"""
    st.download_button("Download free Chapter Four checklist (.txt)", kit_text, file_name="ResearchBaseHQ_Chapter4_Checklist.txt", mime="text/plain")
    lead_email = st.text_input("Email address to receive occasional research tips (optional)", key="lead_email_capture")
    marketing_consent = st.checkbox("I agree to receive occasional Research Base HQ research tips and product updates by email. I can opt out at any time.", key="lead_consent")
    if st.button("Send me the free kit and tips", key="lead_submit"):
        if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", lead_email.strip()):
            st.error("Enter a valid email address, or download the checklist above without providing an email.")
        else:
            try:
                email_sent = save_lead(lead_email, marketing_consent)
                if email_sent:
                    st.success("Your request was saved and the checklist email was sent. You can also download the checklist above.")
                else:
                    st.success("Your request was saved. Email delivery is not configured or failed, so download the checklist above. No email has been confirmed as sent.")
            except Exception as exc:
                st.error("Lead capture is temporarily unavailable. You can still download the checklist above.")

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
    track_event("checkout_started", email=student_email, project_key=f"{reg_no}|{project_title}", metadata=tier_choice)
    url, reference = initialize_paystack(student_email.strip(), PRICE_TIERS[tier_choice]["amount"], meta)
    if url:
        st.session_state["pending_paystack_reference"] = reference
        st.link_button("Continue to secure Paystack checkout →", url, type="primary", use_container_width=True)

def grant(hit, key):
    st.session_state.paid=True; st.session_state.verify_code=hit["code"]
    st.session_state.tier=hit["tier"]; st.session_state.last_project_paid=key

current_project_key = f"{reg_no.strip().lower()}|{project_title.strip().lower()}"
if reg_no.strip() and project_title.strip() and student_email.strip():
    if not st.session_state.get("_project_started_key") == current_project_key:
        track_event("project_started", email=student_email, project_key=current_project_key)
        st.session_state["_project_started_key"] = current_project_key

if reg_no.strip() and project_title.strip():
    st.subheader("Your Research Project Dashboard")
    stages = ["Project details", "Upload data", "Confirm analysis", "Review discussion sources", "Format document", "Download and final review"]
    latest_stage = get_project_stage(current_project_key)
    stage_rank = {"project_started": 0, "dataset_uploaded": 1, "analysis_completed": 2, "discussion_opened": 3, "formatting_opened": 4, "report_downloaded": 5}
    durable_rank = stage_rank.get(latest_stage, -1)
    completed = [bool(student_email.strip()), ('df' in st.session_state or durable_rank >= 1), (bool(st.session_state.get('analysis_confirmed')) or durable_rank >= 2), (bool(st.session_state.get('discussion_sources_reviewed')) or durable_rank >= 3), (bool(st.session_state.get('formatted_doc')) or durable_rank >= 4), (bool(st.session_state.get('last_downloaded')) or durable_rank >= 5)]
    cols = st.columns(3)
    for idx, (stage, done) in enumerate(zip(stages, completed)):
        with cols[idx % 3]:
            st.markdown(("✅ **" if done else "◻️ **") + stage + "**")
    st.caption("Progress reflects this browser session unless your deployment has configured durable project storage. Do not upload identifiable or confidential participant data unless you are authorised to do so.")

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
    st.markdown("<div class='section-kicker'>Simple project-based pricing</div><h2 style='margin-top:4px'>Choose the level of support you need</h2><p style='color:#526579'>Review the included features before you pay. Both plans are priced per project topic under the current offer.</p>", unsafe_allow_html=True)
    p1,p2=st.columns(2, gap="large")
    with p1:
        st.markdown("""
<div class="price-card">
  <div style="color:#526579;font-weight:800;font-size:12px;letter-spacing:.1em">CORE WORKFLOW</div>
  <h2 style="margin:8px 0 0">Basic</h2>
  <div style="font-size:34px;font-weight:850;color:#102A43;margin:4px 0">₦30,000</div>
  <div style="color:#526579;font-size:13px">Per project topic</div><hr>
  <div style="line-height:2.0;font-size:14px">
    ✓ Dataset template and upload workflow<br>
    ✓ Descriptive statistics and frequency tables<br>
    ✓ Supported inferential tests (where appropriate)<br>
    ✓ Graphs and supported export formats<br>
    ✓ Results tables to review for Chapter Four<br>
    ✓ Project formatting tools<br>
  </div>
  <p style="font-size:12px;color:#64748B;margin-top:12px">Availability depends on file type, selected method and current app capabilities. Verify outputs before use.</p>
</div>
""", unsafe_allow_html=True)
    with p2:
        st.markdown("""
<div class="price-card popular">
  <span class="badge">MORE WRITING SUPPORT</span>
  <h2 style="margin:12px 0 0">Standard</h2>
  <div style="font-size:34px;font-weight:850;color:#102A43;margin:4px 0">₦40,000</div>
  <div style="color:#526579;font-size:13px">Per project topic</div><hr>
  <div style="line-height:2.0;font-size:14px">
    ✓ Everything in Basic<br>
    ✓ Writing clarity and style feedback<br>
    ✓ Revision checklist<br>
    ✓ Send-to-admin review workflow, where available<br>
    ✓ Expanded project-development support<br>
    ✓ Support for questions about the workflow<br>
  </div>
  <p style="font-size:12px;color:#64748B;margin-top:12px">Confirm support scope and availability before checkout. No plan guarantees approval, a particular grade or publication.</p>
</div>
""", unsafe_allow_html=True)
    st.markdown("#### Ready to continue?")
    st.write("Enter your project details below, then choose a plan. You can still use the free template and explore the workflow without paying.")
    tier_choice=st.radio("Choose a plan", list(PRICE_TIERS.keys()), index=1, horizontal=False); st.session_state.selected_tier=tier_choice
    if student_email and reg_no and project_title and "@" in student_email:
        if st.button(f"Continue to secure payment · ₦{PRICE_TIERS[tier_choice]['display']:,}", type="primary", use_container_width=True):
            start_checkout(tier_choice)
    else:
        st.info("To purchase, complete your registration number, project title and email fields above. You can explore the free tools first.")
    with st.expander("Frequently asked questions"):
        st.markdown("**Can I try anything for free?** Yes. The data template and sample output are available to explore without payment. Some full project workflows require a paid plan.")
        st.markdown("**Will the app choose the right statistical test automatically?** You must confirm the method against your design and approved methodology. The app supports only the methods implemented in the current version.")
        st.markdown("**Are references automatically guaranteed to be correct?** No. Verify the details of every source, DOI and claim before citing it.")
        st.markdown("**Does the app guarantee supervisor approval or a grade?** No. It is a support tool, not a substitute for your own work, supervisor guidance or statistical review.")
        st.markdown(f"**I have a payment or access problem.** Contact [Research Base HQ support]({WHATSAPP_LINK}). Keep your Paystack transaction reference if you were debited.")

if st.session_state.paid:
    st.success(f"✅ Active for: '{project_title}' — {VERIFY_DOMAIN}/{st.session_state.verify_code} — No watermark — Correct Order")

tab1,tab2,tab3,tab4,tab5,tab6=st.tabs(["1️⃣ Upload (FREE)","2️⃣ Analysis","3️⃣ Discuss with PubMed/Scholar","4️⃣ Format — Choose Uni Format","5️⃣ Send to Admin","6️⃣ Writing Quality Review — Standard"])

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
            upload_fingerprint = hashlib.sha256(up.getvalue()).hexdigest()
            if st.session_state.get("_uploaded_event_fingerprint") != upload_fingerprint:
                track_event('dataset_uploaded', email=st.session_state.get('student_email_in',''), project_key=f"{st.session_state.get('reg_no_in','')}|{st.session_state.get('project_title_in','')}", metadata=f"rows={len(df)},columns={len(df.columns)}")
                st.session_state["_uploaded_event_fingerprint"] = upload_fingerprint
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
        df=st.session_state.df; ch3_info=st.session_state.get('ch3_info',{"methods":[],"dv_text":"","needs_review":True})
        dep_col=st.session_state.get('auto_dv', df.columns[0])
        indep_cols=st.session_state.get('auto_ivs', [])
        group_col=st.session_state.get('group_col', None)
        st.info(f"📌 From YOUR Variable_Map: Primary Dependent = **{dep_col}** | Independent = **{', '.join(indep_cols) if indep_cols else 'None'}** | Group = **{group_col if group_col else 'None'}**")
        st.caption("Correct Chapter 4 Order: descriptive summaries first, then only inferential tests justified by the approved study design. The app will not infer a test from a column's number of unique values.")
        missing_report = pd.DataFrame({"Column": df.columns, "Data type": [str(df[c].dtype) for c in df.columns], "Missing values": [int(df[c].isna().sum()) for c in df.columns], "Unique values": [int(df[c].nunique(dropna=True)) for c in df.columns]})
        with st.expander("Review data quality before analysis", expanded=True):
            st.write(f"**Rows:** {len(df)} | **Columns:** {len(df.columns)} | **Duplicate rows:** {int(df.duplicated().sum())}")
            st.dataframe(missing_report, use_container_width=True, hide_index=True)
            if int(df.duplicated().sum()): st.warning("Duplicate rows detected. Confirm whether they are genuine repeated observations before removing any records.")
        c1,c2,c3=st.columns(3)
        with c1: dep_col=st.selectbox("Dependent (DV) — 1 only", df.columns, index=list(df.columns).index(dep_col) if dep_col in df.columns else 0)
        with c2: indep_cols=st.multiselect("Independent (IVs)", [c for c in df.columns if c!=dep_col], default=[c for c in indep_cols if c in df.columns and c!=dep_col])
        with c3: group_col_sel=st.selectbox("Group Column (for rats, else None)", ["None"]+list(df.columns), index=(["None"]+list(df.columns)).index(group_col) if group_col in df.columns else 0)
        group_col = None if group_col_sel=="None" else group_col_sel
        st.markdown("#### Confirm the analysis plan")
        proposed = ch3_info.get("methods", []) if isinstance(ch3_info, dict) else []
        allowed_methods = ["ANOVA", "T-Test", "Correlation", "Regression", "Tukey"]
        default_methods = [m for m in proposed if m in allowed_methods]
        selected_methods = st.multiselect("Select only methods approved in your methodology/supervisor's guidance", allowed_methods, default=default_methods, key="confirmed_methods")
        if not selected_methods:
            st.info("No inferential method selected. Descriptive summaries can still be reviewed; confirm a method before inferential analysis.")
        if selected_methods and st.checkbox("I have checked that this analysis matches my approved study design.", key="analysis_plan_confirmation"):
            st.session_state["analysis_confirmed"] = True
        else:
            st.session_state["analysis_confirmed"] = False
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
            infer=run_inferential_correct_order(df, dep_col, indep_cols, group_col, selected_methods) if st.session_state.get('analysis_confirmed') else {}
            if st.session_state.get('analysis_confirmed'):
                analysis_event_key = current_project_key + "|" + ",".join(selected_methods)
                if st.session_state.get("_analysis_event_key") != analysis_event_key:
                    track_event('analysis_completed', email=student_email, project_key=current_project_key, metadata=','.join(selected_methods))
                    st.session_state["_analysis_event_key"] = analysis_event_key
            else:
                st.warning('Inferential analysis is paused until you confirm the approved analysis plan. Descriptive tables remain available.')
            st.divider(); st.subheader("4️⃣ Inferential Statistics — Chapter 4.2 — After Descriptive")
            if 'Data_Handling' in infer:
                st.write("**Data included/excluded for numeric-variable analysis**")
                st.json(infer['Data_Handling'])
            if 'ANOVA' in infer:
                st.write(f"**One-Way ANOVA — {infer['ANOVA']['factor']} on {dep_col}**")
                st.json(infer['ANOVA'])
                if infer['ANOVA'].get('Levene p-value', 1) < 0.05:
                    st.warning("Group variances may differ. Standard one-way ANOVA/Tukey results require careful review; consider a design-appropriate robust alternative with your supervisor/statistician.")
            if 'PostHoc_Error' in infer:
                st.warning(infer['PostHoc_Error'])
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
            if st.download_button("📑 8. Chapter 4 All Tables in Correct Order — Excel — No Watermark", data=excel_buf.getvalue(), file_name=f"{reg_no}_Chapter4_CorrectOrder.xlsx", use_container_width=True):
                st.session_state["last_downloaded"] = True
                track_event("report_downloaded", email=student_email, project_key=current_project_key, metadata="Chapter4 Excel")
            for name,fig in graphs:
                buf=io.BytesIO(); fig.savefig(buf,format='png',dpi=300,bbox_inches='tight'); buf.seek(0)
                st.download_button(f"🖼️ {name}.png — No Watermark", data=buf, file_name=f"{name}.png", mime="image/png", key=f"dl_{name}")

# ---------------------------------------------------------------- TAB 3
with tab3:
    if st.session_state.get("_discussion_opened_event_key") != current_project_key:
        track_event("discussion_opened", email=student_email, project_key=current_project_key)
        st.session_state["_discussion_opened_event_key"] = current_project_key
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
        if st.button("Generate Discussion Draft Using Your Supplied Sources", type="primary", use_container_width=True):
            st.session_state["discussion_sources_reviewed"] = bool(similar_articles.strip())
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
                    ch3 = st.session_state.get('ch3_info',{"methods":[]})
                    infer_temp = run_inferential_correct_order(st.session_state.df, dep_col, st.session_state.get('auto_ivs',[]), st.session_state.get('group_col'), ch3['methods'])
                    if 'ANOVA' in infer_temp:
                        anova_line = f"One-way ANOVA showed significant difference in {dep_col} across {infer_temp['ANOVA']['factor']} (F={infer_temp['ANOVA']['F']}, p={infer_temp['ANOVA']['Sig.']}). Review the post-hoc output before describing specific group differences."
            except: pass
            refs_guideline = similar_articles.strip() if similar_articles.strip() else "No article details were supplied. Add verified article titles, authors, years, and DOI/URLs before making literature comparisons."
            draft = f"""CHAPTER FIVE — DISCUSSION — {project_title.upper()}

This {study_type_note.lower()} was conducted at {uni}, Department of {dept}, to investigate {objective if objective else project_title}. {mean_line}

{anova_line} Main finding: {main_finding}.

Comparison with similar articles as guideline (from PubMed, Google Scholar, Consensus.ai):
Source notes supplied for checking: {refs_guideline}. Do not claim agreement or disagreement with prior studies until each source has been opened and its methods and findings verified. Insert only comparisons that are supported by the cited article text.

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
    if st.session_state.get("_formatting_opened_event_key") != current_project_key:
        track_event("formatting_opened", email=student_email, project_key=current_project_key)
        st.session_state["_formatting_opened_event_key"] = current_project_key
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
        work=st.text_area("Paste work for review", height=200)
        st.caption("Only the first 4,000 characters are sent. For longer work, send the file on WhatsApp.")
        if st.button("Submit to Creator"):
            if not work.strip():
                st.warning("Paste your work first")
            elif save_review([datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), uni, dept, project_title, student_email, _safe_cell(work.strip())[:4000], st.session_state.verify_code, "", reg_no]):
                review_rows.clear()
                st.success("Submitted — the admin's reply will appear below.")
            else:
                st.error("Could not submit right now — please send it on WhatsApp instead.")
        st.divider()
        st.subheader("Your submissions & admin replies")
        if st.button("🔄 Check for replies"): review_rows.clear()
        mine = [(n, r) for n, r in (review_rows() or []) if _rcell(r, "code") == str(st.session_state.verify_code) and _rcell(r, "reg").lower() == reg_no.strip().lower()]
        if not mine:
            st.caption("No submissions yet.")
        for n, r in reversed(mine):
            with st.expander(f"{_rcell(r, 'date')} — {_rcell(r, 'work')[:70]}…"):
                st.text(_rcell(r, "work"))
                if _rcell(r, "reply"): st.success(_rcell(r, "reply"))
                else: st.info("⏳ Waiting for the admin's reply")

# ---------------------------------------------------------------- TAB 6 — WRITING QUALITY REVIEW (NOT AN AI DETECTOR)
with tab6:
    if not st.session_state.paid:
        st.markdown("<div class='locked-box'><h3>🔒 Writing Quality Review Locked</h3><p>Included in Standard.</p></div>", unsafe_allow_html=True)
    elif "Standard" not in str(st.session_state.tier):
        st.error("Writing Quality Review is included in the Standard plan.")
    else:
        st.header(f"📝 Writing Quality Review — {project_title}")
        st.info("This tool checks selected writing features. It cannot reliably determine whether text was written by AI, predict an AI-detector score, or guarantee a supervisor's decision.")
        check_text = st.text_area("Paste a section for writing-quality feedback", height=250, placeholder="Paste a paragraph or section you want to review...")
        if st.button("Review writing quality", type="primary", use_container_width=True):
            if not check_text.strip():
                st.warning("Paste text to review.")
            else:
                words = re.findall(r"\b[\w'-]+\b", check_text)
                sentences = [s.strip() for s in re.split(r"[.!?]+", check_text) if s.strip()]
                sentence_lengths = [len(re.findall(r"\b[\w'-]+\b", s)) for s in sentences]
                repeated = sorted(set(re.findall(r"\b(?:in conclusion|furthermore|moreover|additionally|it is important to note that)\b", check_text.lower())))
                feedback = []
                if len(words) < 100: feedback.append("This is a short passage; assess the wider section before drawing conclusions.")
                if sentence_lengths and max(sentence_lengths) > 35: feedback.append("Review very long sentences for clarity and split them where appropriate.")
                if sentence_lengths and sum(sentence_lengths)/len(sentence_lengths) > 25: feedback.append("The average sentence is long. Consider a mix of concise and more detailed sentences.")
                if repeated: feedback.append("Review repeated transition phrases: " + ", ".join(repeated) + ".")
                if not re.search(r"\([^)]*\b(?:19|20)\d{2}\b[^)]*\)", check_text) and not re.search(r"\b(?:19|20)\d{2}\b", check_text): feedback.append("Check whether factual claims need citations; this tool does not verify references.")
                if not feedback: feedback.append("No basic style flags were found by these simple checks. This does not establish originality or determine authorship.")
                st.session_state["writing_review_result"] = {"words":len(words),"sentences":len(sentences),"avg_sentence_words":round(sum(sentence_lengths)/len(sentence_lengths),1) if sentence_lengths else 0,"feedback":feedback,"text":check_text}
                track_event("writing_quality_reviewed", email=student_email, project_key=current_project_key)
        result = st.session_state.get("writing_review_result")
        if result:
            c1,c2,c3 = st.columns(3)
            c1.metric("Word count", result["words"])
            c2.metric("Sentence count", result["sentences"])
            c3.metric("Average words per sentence", result["avg_sentence_words"])
            st.subheader("Suggested review checklist")
            for item in result["feedback"]: st.write("- " + item)
            st.warning("Edit for accuracy, clarity, your own analysis and correct citations. Do not change wording merely to evade AI-detection systems.")
            st.download_button("Download review notes", "WRITING QUALITY REVIEW\n\n" + "\n".join("- " + x for x in result["feedback"]), file_name="ResearchBaseHQ_Writing_Review.txt", mime="text/plain")
# ---------------------------------------------------------------- FOOTER
st.divider()
st.markdown(f"""
<div style='text-align:center; font-size:13px; color:gray;'>
<b>{APP_NAME}</b> — Research workflow support for students. Check analysis outputs, verify references, and confirm formatting requirements before submission.<br>
Helpline: <b>{WHATSAPP_DISPLAY}</b> — <a href="{WHATSAPP_LINK}" target="_blank">Chat on WhatsApp</a><br>
{VERIFY_DOMAIN} — © 2026<br><small>Research support only: verify statistical choices, citations and institutional requirements. Avoid uploading identifiable participant data without authorisation.</small>
</div>
""", unsafe_allow_html=True)
