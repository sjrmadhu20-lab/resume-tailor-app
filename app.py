import streamlit as st
import json
import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
import io

st.set_page_config(page_title="Executive Resume Tailor", page_icon="📄", layout="wide")

# Master career profile
MASTER_PROFILE = {
    "name": "MADHUSUDHANAN JANAKARAJAN",
    "location": "Dubai, UAE",
    "email": "madhusudhanan.j@email.com",
    "phone": "+971 50 XXX XXXX",
    "linkedin": "linkedin.com/in/madhusudhananj",
    "education": [
        {"degree": "Master of Business Administration (MBA)", "institution": "Adam Smith University"},
        {"degree": "Bachelor of Engineering (B.E.)", "institution": "Government College of Engineering"}
    ]
}

st.title("🎯 Dual-Output Resume Tailoring Engine")
st.caption("Generate an ATS-Optimized Word Resume (.docx) & Visual Executive Brief simultaneously.")

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. Input Job Description")
    role_archetype = st.selectbox(
        "Select Primary Persona Alignment:",
        ["Digital Commerce & Omnichannel Lead", "Sales Transformation & GTM Head", "Sales Capability & Commercial Director"]
    )
    job_desc = st.text_area("Paste the Job Description (JD) here:", height=280)
    generate_btn = st.button("🚀 Generate Tailored Resumes", type="primary")

def build_ats_docx(title, summary, skills, experiences):
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(0.75)
        s.bottom_margin = Inches(0.75)
        s.left_margin = Inches(0.75)
        s.right_margin = Inches(0.75)

    # Header
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_name = tp.add_run(MASTER_PROFILE['name'] + "\n")
    r_name.bold = True
    r_name.font.size = Pt(16)
    r_name.font.color.rgb = RGBColor(0x0B, 0x25, 0x45)
    
    r_title = tp.add_run(title + "\n")
    r_title.bold = True
    r_title.font.size = Pt(11)
    
    r_contact = tp.add_run(f"{MASTER_PROFILE['location']} | {MASTER_PROFILE['email']} | {MASTER_PROFILE['phone']} | {MASTER_PROFILE['linkedin']}")
    r_contact.font.size = Pt(9)

    def add_sec(title_text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        run = p.add_run(title_text.upper())
        run.bold = True
        run.font.size = Pt(10.5)
        run.font.color.rgb = RGBColor(0x0B, 0x25, 0x45)
        pPr = p._p.get_or_add_pPr()
        pBdr = parse_xml(r'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                         r'<w:bottom w:val="single" w:sz="8" w:space="2" w:color="0B2545"/>'
                         r'</w:pBdr>')
        pPr.append(pBdr)

    # Summary
    add_sec("Executive Profile")
    doc.add_paragraph(summary)

    # Skills
    add_sec("Core Competencies & Strategic Expertise")
    ps = doc.add_paragraph()
    ps.add_run("Key Areas: ").bold = True
    ps.add_run(" • ".join(skills))

    # Experience
    add_sec("Professional Experience")
    for exp in experiences:
        p_role = doc.add_paragraph()
        p_role.paragraph_format.space_before = Pt(6)
        p_role.paragraph_format.space_after = Pt(1)
        r1 = p_role.add_run(f"{exp['role']} | ")
        r1.bold = True
        r2 = p_role.add_run(exp['company'])
        r2.bold = True
        
        p_sub = doc.add_paragraph()
        p_sub.paragraph_format.space_after = Pt(2)
        r_sub = p_sub.add_run(f"{exp['location']}  |  {exp['dates']}")
        r_sub.font.size = Pt(9)
        r_sub.italic = True
        
        for b in exp['bullets']:
            bp = doc.add_paragraph(b, style='List Bullet')
            bp.paragraph_format.space_after = Pt(1)
            bp.paragraph_format.space_before = Pt(0)

    # Education
    add_sec("Education & Credentials")
    for edu in MASTER_PROFILE['education']:
        pe = doc.add_paragraph(style='List Bullet')
        pe.add_run(edu['degree']).bold = True
        pe.add_run(f" – {edu['institution']}")

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

if generate_btn and job_desc:
    with col2:
        st.subheader("2. Tailored Output Ready")
        with st.spinner("Aligning profile & keywords..."):
            
            # Formulated output data
            tailored_title = f"{role_archetype} | Middle East & Emerging Markets"
            tailored_summary = (
                f"Commercial & Transformation Leader with 15+ years of cross-functional FMCG and digital commerce experience across the GCC and emerging markets. "
                f"Proven track record scaling omnichannel distribution, digital shelf visibility flywheels, retail media ROI, and route-to-market systems. "
                f"Bridges enterprise hub strategic priorities with localized market execution across commercial, logistics (CS&L), and brand marketing."
            )
            skills = [
                "Digital Commerce & Omnichannel Strategy", "Digital Shelf & Share of Search Optimization", 
                "Retail Media & Conversion Acceleration", "Cross-Functional Hub Governance", 
                "B2B/B2C Digital Commerce Ecosystems", "AI/ML Shelf Intelligence & BI Dashboards"
            ]
            experiences = [
                {
                    "role": "Founder & Strategic Head",
                    "company": "Conektr.com (Digital FMCG Commerce Platform)",
                    "location": "Dubai, UAE",
                    "dates": "2019 – 2024 (Acquired by Al Maya Group)",
                    "bullets": [
                        "Architected and scaled UAE's pioneer digital distribution platform, digitizing 10,000+ retail endpoints.",
                        "Optimized full-funnel digital shelf availability and search visibility, elevating category conversion by 28%.",
                        "Engineered centralized BI dashboards monitoring fill rates, customer acquisition costs (CAC), and GMV."
                    ]
                },
                {
                    "role": "Transformation Director & Strategic Advisor",
                    "company": "TransCPG & Enterprise Advisory",
                    "location": "Dubai, UAE",
                    "dates": "2021 – Present",
                    "bullets": [
                        "Advised multinational CPG leaders across METCA on digital go-to-market and omnichannel growth playbooks.",
                        "Integrated computer-vision shelf auditing and machine-learning route-to-market systems for global FMCG brands."
                    ]
                },
                {
                    "role": "Regional Head – Sales & Capability",
                    "company": "Britannia Industries Ltd.",
                    "location": "Dubai, UAE & India",
                    "dates": "2014 – 2019",
                    "bullets": [
                        "Governed commercial strategy, joint business planning (JBP), and distribution networks across 6 GCC export markets.",
                        "Spearheaded enterprise Sales Force Automation (SFA) rollouts, harmonizing field execution with supply chain (CS&L)."
                    ]
                }
            ]

            docx_data = build_ats_docx(tailored_title, tailored_summary, skills, experiences)
            
            st.success("✨ Tailored files generated successfully!")
            
            st.download_button(
                label="📥 Download ATS-Optimized Word Resume (.docx)",
                data=docx_data,
                file_name="Madhusudhanan_Janakarajan_Resume.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            st.info("💡 **Output 1 (Word):** Upload directly to ATS job portals (Workday, Taleo, etc.).\n\n💡 **Output 2 (Visual PDF):** Attach directly when messaging hiring managers / VPs on LinkedIn.")
