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
# 1. RETRIEVE API KEY
# ==============================================================================
api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))

# ==============================================================================
# 2. MASTER KNOWLEDGE ARCHIVE (EXACT MASTER TEMPLATE & ACHIEVEMENTS)
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
# 3. WORD DOCUMENT GENERATION ENGINE (EXACT TYPOGRAPHY & 2-PAGE LOCK)
# ==============================================================================
def create_master_resume_docx(tailored_data):
    doc = Document()
    
    # Precise Margins matching Master Resume
    for section in doc.sections:
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.45)
        section.right_margin = Inches(0.45)

    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(10)
    style.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    def format_heading(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(5)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(title.upper())
        r.bold = True
        r.font.name = 'Calibri'
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        pPr = p._p.get_or_add_pPr()
        pBdr = parse_xml(r'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                         r'<w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/>'
                         r'</w:pBdr>')
        pPr.append(pBdr)

    # ---------------- PAGE 1 ----------------
    # 1. Header Section
    hp = doc.add_paragraph()
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hp.paragraph_format.space_after = Pt(1)
    
    # Name: 10 pt
    r_name = hp.add_run(MASTER_STATIC['name'] + "\n")
    r_name.bold = True
    r_name.font.name = 'Calibri'
    r_name.font.size = Pt(10)
    
    # Subtitle line: 9 pt
    f1 = tailored_data.get("header_focus_1", "Sales & Distribution Transformation Director")
    f2 = tailored_data.get("header_focus_2", "Beauty & Personal Care Experience")
    full_header_line = f"{f1} | FMCG | GTM & Omnichannel Leader | {f2}"
    
    r_title = hp.add_run(full_header_line + "\n")
    r_title.bold = True
    r_title.font.name = 'Calibri'
    r_title.font.size = Pt(9)
    
    # Remaining 3 Header Lines: 10 pt
    c = MASTER_STATIC['contact']
    r_c1 = hp.add_run(f"{c['location']} | {c['phone']} | {c['email']}\n")
    r_c1.font.name = 'Calibri'
    r_c1.font.size = Pt(10)
    
    r_c2 = hp.add_run(f"{c['linkedin']} | Portfolio: {c['portfolio']}\n")
    r_c2.font.name = 'Calibri'
    r_c2.font.size = Pt(10)
    
    r_c3 = hp.add_run(f"Visa Status: {c['visas']}")
    r_c3.font.name = 'Calibri'
    r_c3.font.size = Pt(10)

    # 2. Executive Summary (10 pt, Locked to 7-8 Lines)
    format_heading("EXECUTIVE SUMMARY")
    sp = doc.add_paragraph(tailored_data.get("executive_summary", ""))
    sp.paragraph_format.space_before = Pt(2)
    sp.paragraph_format.space_after = Pt(3)
    sp.paragraph_format.line_spacing = 1.05
    for r in sp.runs:
        r.font.name = 'Calibri'
        r.font.size = Pt(10)

    # 3. Executive Capabilities & Impact Highlights (10 pt, 5 Points)
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
            r_bold.font.name = 'Calibri'
            r_bold.font.size = Pt(10)
            r_body = cp.add_run(parts[1])
            r_body.font.name = 'Calibri'
            r_body.font.size = Pt(10)
        else:
            r_body = cp.add_run(cap)
            r_body.font.name = 'Calibri'
            r_body.font.size = Pt(10)

    # 4. Honors & Recognition (10 pt)
    format_heading("HONORS & RECOGNITION")
    for h in MASTER_STATIC['honors']:
        hp = doc.add_paragraph(h, style='List Bullet')
        hp.paragraph_format.space_before = Pt(0)
        hp.paragraph_format.space_after = Pt(1)
        for r in hp.runs:
            r.font.name = 'Calibri'
            r.font.size = Pt(10)

    # 5. Education & Languages (10 pt)
    format_heading("EDUCATION")
    for edu in MASTER_STATIC['education']:
        ep = doc.add_paragraph(style='List Bullet')
        ep.paragraph_format.space_before = Pt(0)
        ep.paragraph_format.space_after = Pt(1)
        r_deg = ep.add_run(edu['degree'] + " – ")
        r_deg.bold = True
        r_deg.font.name = 'Calibri'
        r_deg.font.size = Pt(10)
        r_det = ep.add_run(edu['details'])
        r_det.font.name = 'Calibri'
        r_det.font.size = Pt(10)

    format_heading("LANGUAGES & INTERESTS")
    lp = doc.add_paragraph()
    lp.paragraph_format.space_before = Pt(1)
    lp.paragraph_format.space_after = Pt(1)
    r_l1 = lp.add_run(MASTER_STATIC['languages'] + "\n")
    r_l1.font.name = 'Calibri'
    r_l1.font.size = Pt(10)
    r_l2 = lp.add_run(MASTER_STATIC['interests'])
    r_l2.font.name = 'Calibri'
    r_l2.font.size = Pt(10)

    # ---------------- PAGE 2 EXPLICIT BOUNDARY ----------------
    doc.add_page_break()

    # 6. Professional Experience (3-Column Table)
    format_heading("PROFESSIONAL EXPERIENCE")
    
    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    col_widths = [Inches(2.5), Inches(2.5), Inches(2.5)]
    for i, col in enumerate(table.columns):
        for cell in col.cells:
            cell.width = col_widths[i]

    # Header Row with 12 pt Bold titles
    hdr_cells = table.rows[0].cells
    hdr_titles = ["Traditional FMCG Operator", "Digital FMCG Distribution", "Distribution Transformation"]
    for i, title in enumerate(hdr_titles):
        hdr_cells[i].text = title
        p = hdr_cells[i].paragraphs[0]
        p.runs[0].bold = True
        p.runs[0].font.name = 'Calibri'
        p.runs[0].font.size = Pt(12)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        tcPr = hdr_cells[i]._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(r'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="E9ECEF"/>'))

    row_cells = table.add_row().cells
    for i, col in enumerate(table.columns):
        row_cells[i].width = col_widths[i]

    def add_cell_block(cell, bold_title, sub_title, bullets):
        p_title = cell.add_paragraph()
        p_title.paragraph_format.space_before = Pt(3)
        p_title.paragraph_format.space_after = Pt(1)
        r_bt = p_title.add_run(bold_title)
        r_bt.bold = True
        r_bt.font.name = 'Calibri'
        r_bt.font.size = Pt(10)
        
        if sub_title:
            p_sub = cell.add_paragraph()
            p_sub.paragraph_format.space_before = Pt(0)
            p_sub.paragraph_format.space_after = Pt(2)
            r_st = p_sub.add_run(sub_title)
            r_st.italic = True
            r_st.font.name = 'Calibri'
            r_st.font.size = Pt(9.5)

        for b in bullets:
            bp = cell.add_paragraph(style='List Bullet')
            bp.paragraph_format.space_before = Pt(0)
            bp.paragraph_format.space_after = Pt(1.5)
            bp.paragraph_format.line_spacing = 1.0
            r_b = bp.add_run(b)
            r_b.font.name = 'Calibri'
            r_b.font.size = Pt(9.5)

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
    conektr_category = tailored_data.get("conektr_category_bullet", "Deep Category Portfolio Aggregation: Managed & distributed extensive SKU catalogs across Packaged Foods, Snacking, Beverages, and Grocery essentials.")
    c2_bullets_conektr = [
        "Founded UAE's 1st Digital FMCG Principal-Distributor serving 8,000+ retailers (2,000+ MAU) & 100+ brands.",
        conektr_category,
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

    # 7. Tech Stack & Why Hire Me (10 pt)
    format_heading("TECHNOLOGY STACK & DIGITAL ARCHITECTURE")
    for category, stack in MASTER_STATIC['tech_stack'].items():
        tp = doc.add_paragraph(style='List Bullet')
        tp.paragraph_format.space_before = Pt(0)
        tp.paragraph_format.space_after = Pt(1)
        r_cat = tp.add_run(f"{category}: ")
        r_cat.bold = True
        r_cat.font.name = 'Calibri'
        r_cat.font.size = Pt(10)
        r_st = tp.add_run(stack)
        r_st.font.name = 'Calibri'
        r_st.font.size = Pt(10)

    format_heading("WHY HIRE ME")
    wp = doc.add_paragraph(MASTER_STATIC['why_hire_me'])
    wp.paragraph_format.space_before = Pt(1)
    wp.paragraph_format.space_after = Pt(1)
    for r in wp.runs:
        r.font.name = 'Calibri'
        r.font.size = Pt(10)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ==============================================================================
# 4. STREAMLIT FRONTEND & STRICT PROMPT GOVERNANCE
# ==============================================================================
st.title("🎯 Executive ATS Resume Tailoring Engine")
st.caption("Master Knowledge Archive • Dual Header Variables • Exact Calibri Typography • Locked 2-Page Boundary")

with st.sidebar:
    st.header("⚡ System Status")
    if api_key:
        st.success("🟢 Gemini AI Engine: Active")
    else:
        st.warning("🟠 AI Engine: Inactive (Set GEMINI_API_KEY in Secrets)")
    
    st.markdown("---")
    st.write("📂 **Active Knowledge Archive:**")
    st.caption("• Master Resume (Exact Typography)\n• Additional Achievements (All Categories)\n• Dual Header & Category Engine")

col1, col2 = st.columns([1.1, 0.9])

with col1:
    st.subheader("1. Job Inputs & Specifics")
    job_desc = st.text_area("Target Job Description (JD):", height=240, placeholder="Paste JD here...")
    special_instructions = st.text_area("Special Instructions & Context (Optional):", height=130, 
                                        placeholder="E.g., Target company is a Food/Snacks company, emphasize bakery/confectionery at Britannia and grocery aggregation at Conektr...")
    
    generate_btn = st.button("🚀 Generate Tailored Master Resume", type="primary")

if generate_btn:
    if not job_desc:
        st.warning("Please paste a Job Description first.")
    else:
        with col2:
            st.subheader("2. AI Analysis & Tailored File")
            with st.spinner("Synthesizing JD, Special Instructions & Knowledge Archive..."):
                
                tailored_data = None
                if api_key:
                    prompt = f"""
                    You are an executive resume architect for Madhusudhanan Janakarajan (23+ year FMCG & Digital Commerce Executive).
                    
                    STRICT RULES & CONSTRAINTS:
                    1. SECTION 1 (HEADER TITLE DUAL VARIABLES):
                       - Output TWO parts for the header:
                         * "header_focus_1": The best target title matching the role (e.g. "Sales & Distribution Transformation Director", "Digital Commerce Strategy Lead", "Commercial & Capability Director").
                         * "header_focus_2": The category or focus angle matching the JD (e.g. "Packaged Foods & Snacking Leadership", "Beauty & Personal Care Experience", "Beverages & QSR Distribution", "Regional Hub & Omnichannel GTM").
                       - The final combined line: "[header_focus_1] | FMCG | GTM & Omnichannel Leader | [header_focus_2]" MUST fit strictly on ONE single line (9 pt font). Keep each part concise (30-45 chars max).

                    2. SECTION 2 (EXECUTIVE SUMMARY - EXACT 7 TO 8 LINES):
                       - Must be a rich, authoritative, highly comprehensive paragraph of EXACTLY 135 to 150 words (Strictly 7-8 printed lines in 10 pt font, never more, never less).
                       - Must retain ALL core anchor metrics: 23+ years FMCG commercial strategy, rare 360° vantage (principal FMCG operator, digital distribution founder, enterprise SaaS advisor), $100M+ P&L ownership, Founded Conektr (8,000+ retailers, 100+ brands), Enterprise SFA/RTM modernizations (FieldAssist & Ivy Mobility for 10+ Tier-1 CPGs: P&G, Nestlé, GSK, Coca-Cola), delivering ~40% logistics cost optimization and ~20% sales productivity uplifts.
                       - Seamlessly adapt the narrative focus and terminology to match the target JD without altering facts.

                    3. SECTION 3 (EXECUTIVE CAPABILITIES - EXACT 5 BULLETS):
                       - Re-order and prioritize the 5 master capability themes so the top 2 bullets address the highest priority requirements of the JD.
                       - Maintain all 5 themes: (1) Commercial & GTM Leadership ($100M+ P&L), (2) Digital B2B2C Commerce & Omnichannel RTM, (3) Enterprise Transformation & Commercial Optimization, (4) Sales Capability & Training Leadership, (5) Entrepreneurial Venture Scaling & Governance.
                       - Keep length balanced and formatted as 'Bold Header: Detailed metric description'.

                    4. SECTION 4 (CONEKTR CATEGORY BULLET):
                       - Detect the target company/category from JD or Special Instructions:
                         * If Food / Snacks / Bakery: "Deep Packaged Foods & Grocery Aggregation: Managed & distributed extensive SKU catalogs across Britannia, Mondelez, Kellogg's, Nestlé, and Haldiram's."
                         * If Beverages: "Deep Beverage & Energy Drinks Aggregation: Scaled distribution catalogs across Coca-Cola, PepsiCo, Red Bull, and institutional beverage brands."
                         * If Personal Care / Beauty: "Deep Beauty & Personal Care Aggregation: Managed & distributed extensive SKU catalogs across Unilever (Dove, Sunsilk, Vaseline), P&G (Pantene, Olay), L'Oréal, and Colgate-Palmolive."
                         * If Regulated / Tobacco: "Regulated Categories & Multi-Brand Aggregation: Direct commercial distribution across UAE retail for major international brands (Marlboro, Dunhill, Chesterfield) across premium and value tiers."
                         * If Omnichannel / Tech / Healthcare: Output the relevant brand aggregation from the Knowledge Archive.

                    INPUT JOB DESCRIPTION:
                    {job_desc}

                    INPUT SPECIAL INSTRUCTIONS / CONTEXT:
                    {special_instructions}

                    Return ONLY a valid JSON object:
                    {{
                      "header_focus_1": "string",
                      "header_focus_2": "string",
                      "executive_summary": "string",
                      "capabilities": ["string", "string", "string", "string", "string"],
                      "conektr_category_bullet": "string"
                    }}
                    """
                    
                    client = genai.Client(api_key=api_key)
                    for model_candidate in ["gemini-2.5-flash", "gemini-2.5-pro", "gemini-2.0-flash"]:
                        try:
                            response = client.models.generate_content(
                                model=model_candidate,
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

                # Deterministic Fallback if API fails
                if not tailored_data:
                    tailored_data = {
                        "header_focus_1": "Sales & Distribution Transformation Director",
                        "header_focus_2": "Beauty & Personal Care Experience",
                        "executive_summary": "Sales & Distribution Transformation Leader with 23+ years driving FMCG commercial strategy, digital commerce, and sales technology across MEA, India, and Asia. Delivers a rare 360° operational vantage combining principal-led FMCG commercial leadership, digital distribution entrepreneurship, and enterprise SaaS transformations with $100M+ P&L/portfolio ownership. Founded and scaled the UAE's premier digital B2B2C distribution ecosystem (Conektr), operating as a principal-cum-distribution hub aggregating 100+ global brands and 2,000+ Grocery SKUs across Foods, Beverages, Personal Care, Tobacco and Household essentials to 8,000+ retailers. Across enterprise technology advisory (FieldAssist & Ivy Mobility), spearheaded GTM/SFA modernization and Perfect Store automation for 10+ tier-1 CPG enterprises including P&G, Nestlé, GSK, Coca-Cola, PepsiCo and Haldiram's—consistently delivering ~40% logistics cost optimization and ~20% sales productivity uplifts.",
                        "capabilities": [
                            "Commercial & GTM Leadership ($100M+ P&L): Owned $100M+ annual FMCG revenue across GCC & India, directing 250+ distributors and 600+ field sales teams across GT, MT, Wholesale, B2B, and Institutional channels. Spearheaded RTM redesign, distributor governance, trade margin economics, pricing/promotions, and Order-to-Cash optimization.",
                            "Digital B2B2C Commerce & Omnichannel RTM: Founded and scaled Conektr (UAE's first digital FMCG distributor) to 8,000+ B2B retailers, managing 100+ brands and 2,000+ SKUs across Foods, Beverages, and Non-Food categories. Expanded into direct B2C commerce by launching the consumer app and proprietary BOSS loyalty engine (Buying, Operating, Selling & Saving), turning network grocers into fulfillment micro-hubs/dark stores. Built omnichannel ordering (App, Web, Conversational AI) with fintech-enabled payment rails.",
                            "Enterprise Transformation & Commercial Optimization: Directed multi-country RTM modernizations, DMS/ERP integrations and SFA deployments (Over 5000+ Users) for global CPG leaders (P&G, Nestlé, Haleon/GSK, Coca-Cola). Deployed AI route/beat optimization, AI-driven demand forecasting, and automated ordering—delivering a ~40% drop in logistics/admin costs, >30% reduction in outlet coverage costs, ~30% frontline sales productivity uplift, ~150% expansion in numeric distribution growth.",
                            "Sales Capability & Training Leadership: Established and led regional sales training operations managing a team of certified trainers to design and deliver end-to-end sales induction and leadership curricula up to the Regional Sales Manager (RSM) level. Top-performer in consultative selling frameworks (including SPIN Selling), driving frontline execution rigor, distributor capability building, and institutionalized sales performance standards. Managed Train the Trainer, Soft skills and automation training.",
                            "Entrepreneurial Venture Scaling & Governance: Raised $15M in funding from DIFC VC, veteran FMCG executives (ex-Mondelēz President, BAT CFO) validating commercial credibility, and executed a successful strategic M&A exit to Al Maya Group ($1B+ conglomerate). Awarded UAE Golden Visa and USA O-1A (Extraordinary Ability); featured in Bloomberg, Gulf News, and Magnitt."
                        ],
                        "conektr_category_bullet": "Deep Beauty & Personal Care Aggregation: Managed & distributed extensive SKU catalogs across Unilever (Dove, Sunsilk, Vaseline), P&G (Pantene, Olay), L'Oréal, and Colgate-Palmolive."
                    }

                docx_stream = create_master_resume_docx(tailored_data)
                
                st.success("✅ Tailored Master Resume Generated Successfully!")
                st.download_button(
                    label="📥 Download Tailored ATS Word Document (.docx)",
                    data=docx_stream,
                    file_name="Madhusudhanan_Janakarajan_Tailored_Resume.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                with st.expander("🔍 View AI Tailored Variable Breakdown"):
                    st.write("**Header Variable 1:**", tailored_data.get("header_focus_1"))
                    st.write("**Header Variable 2:**", tailored_data.get("header_focus_2"))
                    st.write("**Executive Summary (7-8 Lines):**", tailored_data.get("executive_summary"))
                    st.write("**Conektr Category Bullet:**", tailored_data.get("conektr_category_bullet"))
```[cite: 15]
