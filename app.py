
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
