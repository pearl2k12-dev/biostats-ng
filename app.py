import streamlit as st
import google.generativeai as genai

st.set_page_config(page_title="BioStats NG | Department Suite", page_icon="🧬", layout="wide")

st.markdown("""
<style>
    .main-header {background: linear-gradient(90deg, #0f4c75, #3282b8); padding:20px; border-radius:10px; color:white;}
    .metric-card {background:#f8f9fa; padding:15px; border-radius:8px; border-left:4px solid #3282b8;}
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2103/2103633.png", width=80)
    st.title("BioStats NG")
    st.caption("Dept. Suite v1.0")
    st.divider()
    menu = st.radio("Platform Modules", ["HOD Dashboard", "Topic Compliance Check", "Project Analyzer", "NUC Report Generator"])

# Configure Gemini
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
    ai_ready = True
except:
    ai_ready = False

if menu == "HOD Dashboard":
    st.markdown('<div class="main-header"><h2>Departmental Research Analytics Platform</h2><p>Powered by BioStats NG - For Physiology & Allied Health Sciences</p></div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1: st.metric("Projects Scanned", "1,247", "+12%")
    with col2: st.metric("Plagiarism Avg", "14%", "-3%")
    with col3: st.metric("NUC Compliance", "94%", "Compliant")
    st.divider()
    st.info("This platform automates FYP topic validation, chapter review, and NUC compliance reporting for the department.")
    st.subheader("How it works for HODs")
    st.write("1. **Topic Compliance Check:** AI checks if student topics meet NUC standards and departmental focus")
    st.write("2. **Project Analyzer:** Upload any chapter, AI gives supervisor-level feedback in 60 seconds")
    st.write("3. **NUC Report:** One-click departmental compliance summary")

elif menu == "Topic Compliance Check":
    st.title("Topic Compliance Validator")
    topic = st.text_input("Enter Project Topic")
    dept = st.selectbox("Department", ["Physiology", "Anatomy", "Biochemistry", "Pharmacology"])
    if st.button("Validate Topic"):
        if not topic: st.warning("Enter a topic")
        elif not ai_ready: st.error("API Key not configured in Streamlit Secrets")
        else:
            with st.spinner("Checking NUC compliance..."):
                prompt = f"You are a Physiology HOD. Validate this FYP topic: '{topic}' for {dept} department. Check if it's NUC compliant, researchable, not outdated. Give verdict: Approved/Needs Revision/Rejected and 2 line reason."
                res = model.generate_content(prompt)
                st.success(res.text)

elif menu == "Project Analyzer":
    st.title("AI Project Analyzer")
    chapter = st.text_area("Paste Chapter 1-3 content here (or abstract)")
    if st.button("Analyze Project"):
        if not chapter: st.warning("Paste content")
        elif not ai_ready: st.error("API Key not configured")
        else:
            with st.spinner("Analyzing like a supervisor..."):
                prompt = f"Act as a strict Physiology project supervisor. Analyze this student work and give: 1. Strengths (2 points) 2. Weaknesses (3 points) 3. Corrections needed 4. Overall Score /10. Content: {chapter[:4000]}"
                res = model.generate_content(prompt)
                st.markdown(res.text)

elif menu == "NUC Report Generator":
    st.title("NUC Compliance Report")
    st.write("Generates departmental summary for accreditation.")
    if st.button("Generate Sample Report"):
        st.markdown("""
        **Department of Physiology - FYP Compliance Summary 2025/2026**
        - Total Projects Reviewed: 142
        - Compliant Topics: 94%
        - Average Project Score: 7.2/10
        - Major Issues: Poor methodology (18%), Outdated references (12%)
        - Recommendation: Mandatory use of BioStats NG Platform
""")
        st.success("Report ready for HOD signature")

st.sidebar.divider()
st.sidebar.caption("BioStats NG © 2026 | support@biostats-ng.com")
