import streamlit as st
import pandas as pd
import hashlib
from io import BytesIO
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import json
from datetime import datetime

st.set_page_config(page_title="BioStats NG - Admin + Student", layout="wide", page_icon="🇳🇬")

# --- ADMIN CONFIG ---
ADMIN_PASSWORD = "admin123"  # CHANGE THIS TO YOUR OWN PASSWORD
if 'paid' not in st.session_state: st.session_state.paid=False
if 'is_admin' not in st.session_state: st.session_state.is_admin=False
if 'payments' not in st.session_state: st.session_state.payments=[]

UNIVERSITIES = ["UNIPORT","RSU","UNILAG","UI","UNN","UNIBEN","DELSU","Other"]
DEPARTMENTS = ["Physiology","Biochemistry","Anatomy","Pharmacology","Microbiology","Medical Lab Science","Nursing","Public Health","Other"]
UNI_RULES = {
"UNIPORT": {"font_name":"Times New Roman","font_size":12,"spacing":2.0},
"RSU": {"font_name":"Times New Roman","font_size":12,"spacing":1.5},
"UNILAG": {"font_name":"Arial","font_size":11,"spacing":1.5},
"UI": {"font_name":"Times New Roman","font_size":12,"spacing":2.0},
"Default": {"font_name":"Times New Roman","font_size":12,"spacing":2.0}
}

# Top bar
col1, col2 = st.columns([4,1])
with col1:
    st.markdown('<div style="font-size:26px;font-weight:800;color:#0A1931;">BioStats NG — Student + Admin 🇳🇬</div>', unsafe_allow_html=True)
with col2:
    if st.button("🔑 Admin Dashboard" if not st.session_state.is_admin else "👨🎓 Back to Student View"):
        if not st.session_state.is_admin:
            st.session_state.show_admin_login=True
        else:
            st.session_state.is_admin=False
            st.rerun()

if st.session_state.get('show_admin_login', False) and not st.session_state.is_admin:
    pwd = st.text_input("Enter Admin Password", type="password")
    if st.button("Login as Admin"):
        if pwd == ADMIN_PASSWORD:
            st.session_state.is_admin=True
            st.session_state.show_admin_login=False
            st.rerun()
        else:
            st.error("Wrong password")

# ================= ADMIN DASHBOARD =================
if st.session_state.is_admin:
    st.divider()
    st.subheader("📊 ADMIN DASHBOARD — Only You See This")
    st.write("This is hidden from students. They cannot see this page.")

    total = len(st.session_state.payments)
    st.metric("Total Paid Students (Test Mode)", total)

    if total==0:
        st.info("No payments yet in test. Click 'Pay ₦30k (Test)' in student view to simulate.")
    else:
        df_pay = pd.DataFrame(st.session_state.payments)
        st.dataframe(df_pay)
        st.download_button("Download All Payments CSV", df_pay.to_csv(index=False), "payments.csv")

    colA, colB = st.columns(2)
    with colA:
        st.write("**Payments by University**")
        if total>0:
            st.bar_chart(pd.DataFrame(st.session_state.payments)['uni'].value_counts())
    with colB:
        st.write("**Payments by Department**")
        if total>0:
            st.bar_chart(pd.DataFrame(st.session_state.payments)['dept'].value_counts())

    st.divider()
    st.write("**Verification Checker**")
    check = st.text_input("Enter verification code to check student")
    if check:
        found = [p for p in st.session_state.payments if p['verify']==check.upper()]
        if found:
            st.success(f"FOUND: {found}")
        else:
            st.error("Not found in test payments")

    st.divider()
    if st.button("Clear All Test Payments (Reset)"):
        st.session_state.payments=[]
        st.rerun()
    st.stop()

# ================= STUDENT VIEW =================
with st.sidebar:
    st.header("Student Details")
    uni = st.selectbox("University Name Only", UNIVERSITIES)
    dept = st.selectbox("Department (Separate)", DEPARTMENTS)
    full_name = st.text_input("Full Name", "John A. Doe")
    matric = st.text_input("Matric No", "U2020/123456")
    topic = st.text_input("Project Title", "Effect of Vernonia on glucose")
    citation_style = st.selectbox("Choose YOUR Citation Style", ["APA 7th (UNIPORT)", "APA 6th", "Vancouver [1],[2]", "MLA 9th", "Chicago Author-Date", "Harvard UNILAG"])
    rule = UNI_RULES.get(uni, UNI_RULES["Default"])
    st.divider()
    if not st.session_state.paid:
        st.markdown('<div style="background:#FFF3CD;padding:10px;border-radius:10px;">🔒 FREE<br>10% + 5 pages</div>', unsafe_allow_html=True)
        if st.button("💳 Pay ₦30k (Test)"):
            if not matric: st.error("Enter Matric")
            else:
                verify = hashlib.sha256(f"{matric}{datetime.now()}".encode()).hexdigest()[:10].upper()
                # Save to admin dashboard
                st.session_state.payments.append({
"date": datetime.now().strftime("%Y-%m-%d %H:%M"),
"full_name": full_name,
"matric": matric,
"uni": uni,
"dept": dept,
"topic": topic,
"citation_style": citation_style,
"verify": verify,
"amount": "₦30k Test"
                })
                st.session_state.paid=True
                st.rerun()
    else:
        st.markdown(f'<div style="background:#D4EDDA;padding:10px;border-radius:10px;">✅ PAID — {matric}</div>', unsafe_allow_html=True)
        if st.button("Back to FREE"):
            st.session_state.paid=False
            st.rerun()

def format_citation(raw, style, idx):
    raw=raw.strip()
    if "APA 7th" in style: return f"{raw} → APA 7th"
    if "Vancouver" in style: return f"[{idx}] {raw}"
    return f"{raw} → {style}"

tab1, tab2, tab3 = st.tabs(["📸 Lab Data","📝 Full Formatter","🧠 Learn to Discuss"])

with tab1:
    st.subheader(f"{uni} — {dept} — Lab Data")
    up = st.file_uploader("Upload lab Excel", type=["xlsx","csv"])
    if up:
        df = pd.read_csv(up) if up.name.endswith(".csv") else pd.read_excel(up)
        total=len(df); preview=max(2,int(total*0.1))
        if not st.session_state.paid:
            st.warning(f"FREE: {total} rows, showing {preview}")
            st.dataframe(df.head(preview))
        else:
            st.success(f"PAID: All {total} rows")
            st.dataframe(df)

with tab2:
    st.subheader(f"Formatter — {uni} + {dept} + {citation_style}")
    doc_up = st.file_uploader("Upload thesis .docx", type=["docx"])
    cite_text = st.text_area("References (one per line)", "Oguntola 2020 diabetes\nAdeola 2021 vernonia")
    if doc_up:
        file_bytes=doc_up.read()
        def build(file_bytes, is_paid):
            src=docx.Document(BytesIO(file_bytes))
            new_doc=docx.Document()
            p=new_doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=p.add_run(f"{topic.upper()}\n\n{full_name.upper()}\n{matric}\n\n{dept}\n{uni}\n"); r.font.name=rule['font_name']
            new_doc.add_page_break()
            new_doc.add_heading("TABLE OF CONTENTS",1)
            for para in src.paragraphs[:10]:
                if para.text.strip(): new_doc.add_paragraph(para.text[:80], style='List Bullet')
            new_doc.add_page_break()
            limit=9999 if is_paid else 50
            for i, para in enumerate(src.paragraphs):
                if i>limit and not is_paid:
                    new_doc.add_paragraph(f"[LOCKED — Pay ₦30k]"); break
                p=new_doc.add_paragraph(); p.paragraph_format.line_spacing=rule['spacing']; p.text=para.text
                for run in p.runs: run.font.name=rule['font_name']; run.font.size=Pt(rule['font_size'])
            cites=[c.strip() for c in cite_text.split("\n") if c.strip()]
            refs=[format_citation(c,citation_style,i+1) for i,c in enumerate(cites)]
            if not is_paid: refs=refs[:2]+[f"[+{len(cites)-2} locked]"]
            new_doc.add_page_break()
            new_doc.add_heading(f"REFERENCES ({citation_style})",1)
            for r in refs: new_doc.add_paragraph(r, style='List Bullet')
            verify=hashlib.sha256(f"{matric}".encode()).hexdigest()[:10].upper()
            new_doc.add_paragraph(f"Verified {matric} {uni} {dept} {verify} {citation_style}")
            out=BytesIO(); new_doc.save(out); return out.getvalue(), refs, verify
        if not st.session_state.paid:
            fb, refs, ver = build(file_bytes, False)
            st.warning(f"FREE: 5 pages + 2 refs {citation_style}")
            st.download_button("Download FREE preview", fb, "FREE.docx")
        else:
            fb, refs, ver = build(file_bytes, True)
            st.success(f"PAID: Full {uni} {dept} {citation_style} {ver}")
            st.download_button("Download FULL", fb, f"{matric}_FULL.docx")

with tab3:
    st.subheader("Learn to Discuss — Teach Mode")
    finding = st.text_input("What did result show?", "400mg reduced glucose")
    if not st.session_state.paid:
        st.info(f"FREE Step 1: You saw {finding}. Write 1 sentence why it matters. Steps 2-4 locked.")
    else:
        st.success("PAID — Full teaching")
        st.markdown(f"""
        **Step 1:** {finding} → Say "mean dropped from X to Y"
        **Step 2:** Compare in {citation_style}
        **Step 3:** Mechanism for {dept}
        **Step 4:** Limitations
        Write yourself, then defend.
""")
        draft = st.text_area("Paste YOUR paragraph")
        if st.button("Grade my paragraph"):
            st.success("Good — now formatted for defense.")
