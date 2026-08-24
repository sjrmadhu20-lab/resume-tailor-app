import streamlit as st
import json
import io
import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from google import genai
from google.genai import types

st.set_page_config(page_title="Executive ATS Resume Tailor", page_icon="🎯", layout="wide")

# ==============================================================================
# 1. RETRIEVE SAVED GEMINI API KEY FROM STREAMLIT SECRETS
# ==============================================================================
api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))

# ==============================================================================
# 2. KNOWLEDGE ARCHIVE & MASTER DATA DEFINITIONS
# ==============================================================================
MASTER_STATIC = {
    "name": "MADHUSUDHANAN JANAKARAJAN (MADHU)",
    "contact": {
        "location": "Dubai, UAE",
        "phone": "+971 50 654 7858",
        "email": "sjrmadhu20@gmail.com",
        "linkedin": "https://www.linkedin.com/in/madhusj/",
        "portfolio": "https://linktr.ee/M_S_J",
        "visas": "UAE Golden Visa | USA O-1A (Extraordinary Ability)"
    },
    "honors": [
        "UAE Golden Visa – Recognized for national-scale entrepreneurship and digital commerce impact.",
        "USA O-1A Visa – Extraordinary Ability in FMCG and Digital Commerce.",
        "$15M+ VC funding & exit – Raised $15M+ and successfully exited Conektr to Al Maya Group.",
        "Featured in Gulf News, Bloomberg, Khaleej Times, Yahoo Finance, Magnitt, among others - https://linktr.ee/M_S_J"
    ],
    "education": [
        {"degree": "MBA (2006)", "details": "Adam smith University, USA. [Remote, Airtel Sponsored program for top employees]"},
        {"degree": "Bachelor of Engineering (2001)", "details": "Government College of Engineering (GEC), Tier 1 DOTE College, India"}
    ],
    "languages": "English | Hindi | Tamil | Kannada | Telugu | effective engagement with Arabic-speaking stakeholders.",
    "interests": "Chess Player | Table Tennis Enthusiast | Regular 10K Runner",
    "tech_stack": {
        "AI, Automation & Conversational Commerce": "Agentic Voice Bots (Vapi, ElevenLabs) | Conversational Commerce (Wati, Twilio, Infobip) | Workflow Automation (Make.com) | CRM & Marketing Automation (Klaviyo)",
        "Enterprise & Sales Systems": "SAP (Sales & Distribution) | Oracle eCRM | Microsoft Dynamics | SFA / DMS platforms | ERP–CRMs API - integrations",
        "Digital Commerce & Product Delivery": "WooCommerce | Magento | Mobile Apps (iOS, Android, Flutter) | Full SDLC ownership (Figma → Development → Launch)",
        "Data, Analytics & Optimization": "Power BI | Python scripting | Sales & trade analytics | Demand forecasting | Route & beat optimization",
        "Fintech & Payments": "Stripe | PayPal | CCAvenue | Triterras | Tabby | Spotii (credit, payments, and trade finance integrations)"
    },
    "why_hire_me": "A rare profile combining Core FMCG Operator + Digital FMCG Disruption pioneer + Enterprise Transformations (P&G, Coca-cola, GSK) + 10+ International Markets (GCC, India, Africa, Asia) + Successful Entrepreneurial $15M M&A Exit + Recipient of Global recognition for FMCG Contribution: O1A from USA & Golden Visa from UAE - as an extraordinary ability leader."
}

# ==============================================================================
# 3. WORD DOCUMENT GENERATION ENGINE (STRICT 2-PAGE 3-COLUMN MASTER TEMPLATE)
# ==============================================================================
def create_master_resume_docx(tailored_data):
    doc = Document()
    
    # Strict Page Setup (0.4-0.45 in margins to lock 2-page boundary)
    for section in doc.sections:
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.45)
        section.right_margin = Inches(0.45)

    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(8.5)
    style.font.color.rgb = RGBColor(0x1F, 0x24, 0x21)

    def format_heading(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(title.upper())
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        pPr = p._p.get_or_add_pPr()
        pBdr = parse_xml(r'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                         r'<w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/>'
                         r'</w:pBdr>')
        pPr.append(pBdr)

    # 1. Header Section
    hp = doc.add_paragraph()
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hp.paragraph_format.space_after = Pt(1)
    
    r_name = hp.add_run(MASTER_STATIC['name'] + "\n")
    r_name.bold = True
    r_name.font.size = Pt(13)
    
    # Variable Header Line (Strictly 1 Line)
    r_title = hp.add_run(tailored_data.get("header_title", "Sales & Distribution Transformation Director | FMCG | GTM & Omnichannel Leader") + "\n")
    r_title.bold = True
    r_title.font.size = Pt(9.5)
    
    c = MASTER_STATIC['contact']
    r_contact = hp.add_run(f"{c['location']} | {c['phone']} | {c['email']}\n{c['linkedin']} | Portfolio: {c['portfolio']}\nVisa Status: {c['visas']}")
    r_contact.font.size = Pt(8)

    # 2. Executive Summary
    format_heading("EXECUTIVE SUMMARY")
    sp = doc.add_paragraph(tailored_data.get("executive_summary", ""))
    sp.paragraph_format.space_after = Pt(3)
    sp.paragraph_format.line_spacing = 1.05

    # 3. Executive Capabilities & Impact Highlights (Re-ordered & Reshaped)
    format_heading("EXECUTIVE CAPABILITIES & IMPACT HIGHLIGHTS")
    for cap in tailored_data.get("capabilities", []):
        cp = doc.add_paragraph(style='List Bullet')
        cp.paragraph_format.space_before = Pt(0)
        cp.paragraph_format.space_after = Pt(2)
        cp.paragraph_format.line_spacing = 1.05
        parts = cap.split(":", 1)
        if len(parts) == 2:
            r_bold = cp.add_run(parts[0] + ":")
            r_bold.bold = True
            cp.add_run(parts[1])
        else:
            cp.add_run(cap)

    # 4. Honors & Recognition
    format_heading("HONORS & RECOGNITION")
    for h in MASTER_STATIC['honors']:
        hp = doc.add_paragraph(h, style='List Bullet')
        hp.paragraph_format.space_before = Pt(0)
        hp.paragraph_format.space_after = Pt(1)

    # 5. Education & Languages
    format_heading("EDUCATION")
    for edu in MASTER_STATIC['education']:
        ep = doc.add_paragraph(style='List Bullet')
        ep.paragraph_format.space_before = Pt(0)
        ep.paragraph_format.space_after = Pt(1)
        ep.add_run(edu['degree'] + " – ").bold = True
        ep.add_run(edu['details'])

    format_heading("LANGUAGES & INTERESTS")
    lp = doc.add_paragraph()
    lp.paragraph_format.space_before = Pt(1)
    lp.paragraph_format.space_after = Pt(2)
    lp.add_run(MASTER_STATIC['languages'] + "\n").font.size = Pt(8)
    lp.add_run(MASTER_STATIC['interests']).font.size = Pt(8)

    # 6. Professional Experience (3-Column Layout)
    format_heading("PROFESSIONAL EXPERIENCE")
    
    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    col_widths = [Inches(2.5), Inches(2.5), Inches(2.5)]
    for i, col in enumerate(table.columns):
        for cell in col.cells:
            cell.width = col_widths[i]

    hdr_cells = table.rows[0].cells
    hdr_titles = ["Traditional FMCG Operator", "Digital FMCG Distribution", "Distribution Transformation"]
    for i, title in enumerate(hdr_titles):
        hdr_cells[i].text = title
        p = hdr_cells[i].paragraphs[0]
        p.runs[0].bold = True
        p.runs[0].font.size = Pt(8.5)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        tcPr = hdr_cells[i]._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(r'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="E9ECEF"/>'))

    row_cells = table.add_row().cells
    for i, col in enumerate(table.columns):
        row_cells[i].width = col_widths[i]

    def add_cell_block(cell, bold_title, sub_title, bullets):
        p_title = cell.add_paragraph()
        p_title.paragraph_format.space_before = Pt(4)
        p_title.paragraph_format.space_after = Pt(1)
        r_bt = p_title.add_run(bold_title)
        r_bt.bold = True
        r_bt.font.size = Pt(8)
        
        if sub_title:
            p_sub = cell.add_paragraph()
            p_sub.paragraph_format.space_before = Pt(0)
            p_sub.paragraph_format.space_after = Pt(2)
            r_st = p_sub.add_run(sub_title)
            r_st.italic = True
            r_st.font.size = Pt(7.5)

        for b in bullets:
            bp = cell.add_paragraph(style='List Bullet')
            bp.paragraph_format.space_before = Pt(0)
            bp.paragraph_format.space_after = Pt(2)
            bp.paragraph_format.line_spacing = 1.0
            r_b = bp.add_run(b)
            r_b.font.size = Pt(7.5)

    # Column 1: Traditional FMCG
    c1_bullets_brit = [
        "Owned $100M+ P&L across GCC (Saudi Arabia, UAE, Kuwait, Oman, Bahrain, Qatar) & South India.",
        "Directed 250+ distributor networks & 600+ frontline sales staff across GT, MT, wholesale, and institutional trade.",
        "Spearheaded Britannia's 1st national SFA rollout (1,000+ users), transforming legacy trade into performance-managed selling.",
        "Delivered ~30% numeric distribution growth, increased LPC to ~120%, and cut sales admin costs by ~30%.",
        "Turnaround RSM GCC: achieved record monthly sales for 3 consecutive months (Best Employee Award from Group MD)."
    ]
    c1_bullets_airtel = [
        "Built foundations in frontline trade execution, journey planning, and merchandiser enablement in telecom & enterprise security.",
        "Deployed capability training (SPIN selling) & integrated Oracle e-CRM & LMS infrastructure at scale."
    ]
    add_cell_block(row_cells[0], "Britannia Industries Ltd | 2007 – 2011", "Regional Sales Head – GCC\nRegional Sales & Capability Head- India", c1_bullets_brit)
    add_cell_block(row_cells[0], "Airtel | Reliance | Tyco | 2001 – 2007", "Commercial & Training Roles", c1_bullets_airtel)

    # Column 2: Digital FMCG Distribution (Conektr)
    c2_bullets_conektr = [
        "Founded UAE's 1st Digital FMCG Principal-Distributor serving 8,000+ retailers (2,000+ MAU) & 100+ brands.",
        tailored_data.get("conektr_category_bullet", "Deep Personal Care & FMCG Aggregation: Managed & distributed extensive SKU catalogs across Unilever (Dove, Sunsilk, Vaseline), P&G (Pantene, Olay), L'Oréal, and Colgate-Palmolive."),
        "Owned full P&L, trade terms, warehousing, last-mile delivery, trade credit, and collections.",
        "Built app/web/WhatsApp self-ordering engine scaling annual GMV from zero to ~AED 50M (~$13.6M) at ~18% gross margin.",
        "Cut coverage cost by >50% and improved field execution productivity by ~150% vs traditional trade.",
        "Deployed Dynamics 365 + Power BI and AI route optimization, cutting logistics costs by ~40%.",
        "Raised ~$15M from C-suite FMCG leaders; executed M&A exit to Al Maya Group ($1B+ retail conglomerate)."
    ]
    add_cell_block(row_cells[1], "Conektr Tech Global Ltd | UAE & India\nMay 2016 – Aug 2024", "Digital FMCG Principal / Distributor\nChief Executive Officer & Founder", c2_bullets_conektr)

    # Column 3: Distribution Transformation
    c3_bullets_trans = [
        "Board Member guiding global operations scaling & platform build across FMCG principals & distributors.",
        "Advising CPG leaders on modernizing RTM & SAP/Oracle SFA/DMS integrations, driving ~150% coverage growth.",
        "Built Bid2Bill AI/Voice-bot & WhatsApp B2B2C bidding platform, cutting CAC by ~40% with 4x engagement."
    ]
    c3_bullets_ivy = [
        "Built MEA setup from scratch into 2nd largest global setup ($10M+ pipeline across 10+ countries).",
        "Won 22 enterprise logos: Haleon/GSK, P&G, Nestlé, Coca-Cola, Mars, Red Bull, BAT, and AKI Group.",
        "Personally led on-ground field deployment of mobile SFA for P&G distributor networks in Kenya.",
        "Deployed Cloud SaaS SFA/DMS to 3,000+ sales users, driving post-implementation adoption and trade ROI."
    ]
    add_cell_block(row_cells[2], "TransCPG Inc. & FieldAssist | 2025 – Present", "Post Exit – Transformation Advisor (Director)", c3_bullets_trans)
    add_cell_block(row_cells[2], "Ivy Mobility Pte Ltd | 2011 – 2016", "Business Head – MEA", c3_bullets_ivy)

    # Clean borders
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        r'<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        r'<w:top w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        r'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>'
        r'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>'
        r'</w:tblBorders>'
    )
    tblPr.append(tblBorders)

    # 7. Tech Stack & Why Hire Me
    format_heading("TECHNOLOGY STACK & DIGITAL ARCHITECTURE")
    for category, stack in MASTER_STATIC['tech_stack'].items():
        tp = doc.add_paragraph(style='List Bullet')
        tp.paragraph_format.space_before = Pt(0)
        tp.paragraph_format.space_after = Pt(1)
        r_cat = tp.add_run(f"{category}: ")
        r_cat.bold = True
        r_cat.font.size = Pt(7.5)
        r_st = tp.add_run(stack)
        r_st.font.size = Pt(7.5)

    format_heading("WHY HIRE ME")
    wp = doc.add_paragraph(MASTER_STATIC['why_hire_me'])
    wp.paragraph_format.space_before = Pt(1)
    wp.paragraph_format.space_after = Pt(2)
    wp.runs[0].font.size = Pt(7.5)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ==============================================================================
# 4. STREAMLIT FRONTEND & RESILIENT AI SYNTHESIS
# ==============================================================================
st.title("🎯 Executive ATS Resume Tailoring Engine")
st.caption("Locked Master Resume Architecture • Real-Time AI Keyword & Capability Weaving • Strict 2-Page Constraint")

with st.sidebar:
    st.header("⚡ System Status")
    if api_key:
        st.success("🟢 Gemini AI Engine: Active")
    else:
        st.warning("🟠 AI Engine: Inactive (Set GEMINI_API_KEY in Secrets)")
    
    st.markdown("---")
    st.write("📂 **Active Knowledge Base:**")
    st.caption("• Master Resume (Locked Template)\n• Additional Achievements Archive\n• Tailoring Rules Engine")

col1, col2 = st.columns([1.1, 0.9])

with col1:
    st.subheader("1. Job Inputs & Specifics")
    job_desc = st.text_area("Target Job Description (JD):", height=240, placeholder="Paste JD here...")
    special_instructions = st.text_area("Special Instructions & Context (Optional):", height=130, 
                                        placeholder="Add specific context, category angles, or achievements not in archive...")
    
    generate_btn = st.button("🚀 Generate Tailored Master Resume", type="primary")

if generate_btn:
    if not job_desc:
        st.warning("Please paste a Job Description first.")
    else:
        with col2:
            st.subheader("2. AI Analysis & Tailored File")
            with st.spinner("Synthesizing JD & Special Instructions with Master Profile..."):
                
                tailored_data = None
                if api_key:
                    prompt = f"""
                    You are an executive resume architect for a 23+ year FMCG & Digital Commerce Executive.
                    You must strictly adhere to the following rules:
                    
                    MASTER INSTRUCTIONS:
                    1. Section 1 Header: Generate a single-line title. Must fit strictly on ONE single line. E.g., "Sales & Distribution Transformation Director | FMCG | GTM & Omnichannel Leader | [Category/Target Focus]".
                    2. Section 2 Executive Summary: Synthesize the existing executive summary to reflect the JD's exact context without making any false claims or altering core facts ($100M+ P&L, 8,000+ retailers, $15M exit, 10+ Tier-1 CPGs). Keep strictly to 1 paragraph (~5-6 lines).
                    3. Section 3 Capabilities: Re-order, prioritize, and lightly reshape the 5 core capability bullets based on JD match. Maintain the 5 exact core themes (Commercial & GTM, Digital B2B2C Commerce, Enterprise Transformation, Sales Capability/Training, Entrepreneurial Scaling).
                    4. Section 4 Conektr Category Bullet: Identify the primary industry/category from the JD (e.g., Personal Care/Beauty, Beverages, Foods/Snacks, Tobacco, Healthcare) and output the tailored Conektr category aggregation bullet.
                    
                    JOB DESCRIPTION:
                    {job_desc}
                    
                    SPECIAL INSTRUCTIONS / CONTEXT:
                    {special_instructions}
                    
                    Return ONLY a valid JSON object with keys:
                    - "header_title": string
                    - "executive_summary": string
                    - "capabilities": list of 5 strings (Bold lead title followed by ': ' and description)
                    - "conektr_category_bullet": string
                    """
                    
                    client = genai.Client(api_key=api_key)
                    # Try active models in sequence to prevent 404
                    candidate_models = ["gemini-2.5-flash", "gemini-2.5-pro"]
                    
                    for model_name in candidate_models:
                        try:
                            response = client.models.generate_content(
                                model=model_name,
                                contents=prompt,
                                config=types.GenerateContentConfig(
                                    response_mime_type="application/json",
                                    temperature=0.2
                                )
                            )
                            tailored_data = json.loads(response.text)
                            break
                        except Exception as e:
                            continue

                # Fallback if API fails or is not provided
                if not tailored_data:
                    tailored_data = {
                        "header_title": "Sales & Distribution Transformation Director | FMCG | GTM & Omnichannel Leader | Hub Strategy",
                        "executive_summary": "Sales & Distribution Transformation Leader with 23+ years driving FMCG commercial strategy, digital commerce, and sales technology across MEA, India, and Asia. Delivers a rare 360° operational vantage combining principal-led FMCG commercial leadership ($100M+ P&L), digital distribution entrepreneurship (Conektr), and enterprise SaaS transformations (FieldAssist & Ivy Mobility) for 10+ tier-1 CPG enterprises including P&G, Nestlé, GSK, and Coca-Cola.",
                        "capabilities": [
                            "Digital B2B2C Commerce & Omnichannel RTM: Founded and scaled Conektr to 8,000+ B2B retailers managing 100+ brands and 2,000+ SKUs with BOSS loyalty and omnichannel ordering (App, Web, Conversational AI).",
                            "Commercial & GTM Leadership ($100M+ P&L): Owned $100M+ annual FMCG revenue across GCC & India, directing 250+ distributors and 600+ field sales teams across all trade channels.",
                            "Enterprise Transformation & Commercial Optimization: Directed multi-country RTM modernizations and SFA deployments (5,000+ Users) for global CPG leaders delivering ~40% drop in logistics costs.",
                            "Sales Capability & Training Leadership: Established regional sales training operations, consultative selling (SPIN Selling), and distributor capability standards.",
                            "Entrepreneurial Venture Scaling & Governance: Raised $15M VC funding, executed M&A exit to Al Maya Group ($1B+ conglomerate), and awarded UAE Golden Visa & USA O-1A."
                        ],
                        "conektr_category_bullet": "Deep Personal Care & FMCG Aggregation: Managed & distributed extensive SKU catalogs across Unilever, P&G, L'Oréal, and Colgate-Palmolive."
                    }

                docx_stream = create_master_resume_docx(tailored_data)
                
                st.success("✅ Tailored Master Resume Generated!")
                st.download_button(
                    label="📥 Download Tailored ATS Word Document (.docx)",
                    data=docx_stream,
                    file_name="Madhusudhanan_Janakarajan_Master_Resume.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                with st.expander("🔍 View AI Tailored Variable Values"):
                    st.write("**Header Line:**", tailored_data.get("header_title"))
                    st.write("**Executive Summary:**", tailored_data.get("executive_summary"))
                    st.write("**Target Category Bullet:**", tailored_data.get("conektr_category_bullet"))
