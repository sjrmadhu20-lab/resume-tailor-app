import streamlit as st
import json
import io
import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from google import genai
from google.genai import types

st.set_page_config(page_title="Executive ATS Resume Tailor v2", page_icon="🎯", layout="wide")

# ==============================================================================
# 1. RETRIEVE API KEY FROM STREAMLIT SECRETS
# ==============================================================================
api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))

# ==============================================================================
# 2. MASTER KNOWLEDGE ARCHIVE (EXACT DATA & METRICS)
# ==============================================================================
MASTER_STATIC = {
    "name": "MADHUSUDHANAN JANAKARAJAN (MADHU)",
    "contact": {
        "location": "Dubai, UAE",
        "phone": "+971 50 654 7858",
        "email": "sjrmadhu20@gmail.com",
        "email_url": "mailto:sjrmadhu20@gmail.com",
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
    "languages": "English | Hindi | Tamil | Kannada | Telugu |   effective engagement with Arabic-speaking stakeholders.",
    "interests": "Chess Player | Table Tennis Enthusiast | Regular 10K Runner",
    "tech_stack": {
        "AI, Automation & Conversational Commerce": "Agentic Voice Bots (Vapi, ElevenLabs) | Conversational Commerce (Wati, Twilio, Infobip) | Workflow Automation (Make.com) | CRM & Marketing Automation (Klaviyo)",
        "Enterprise & Sales Systems": "SAP (Sales & Distribution) | Oracle eCRM | Microsoft Dynamics | SFA / DMS platforms | ERP–CRMs API - integrations",
        "Digital Commerce & Product Delivery": "WooCommerce | Magento | Mobile Apps (iOS, Android, Flutter) | Full SDLC ownership (Figma → Development → Launch)",
        "Data, Analytics & Optimization": "Power BI | Python scripting | Sales & trade analytics | Demand forecasting | Route & beat optimization",
        "Fintech & Payments": "Stripe | PayPal | CCAvenue | Triterras | Tabby | Spotii (credit, payments, and trade finance integrations)"
    },
    "why_hire_me_parts": [
        ("A rare profile combining Core FMCG Operator ", False),
        ("+", True),
        (" Digital FMCG Disruption pioneer ", False),
        ("+", True),
        (" Enterprise Transformations (P&G, Coca-cola, GSK) ", False),
        ("+", True),
        (" 10+ International Markets (GCC, India, Africa, Asia) ", False),
        ("+", True),
        (" Successful Entrepreneurial $15M M&A Exit ", False),
        ("+", True),
        (" Recipient of Global recognition for FMCG Contribution: O1A from USA & Golden Visa from UAE - as an extraordinary ability leader.", False)
    ]
}

# XML Helper for clickable Word hyperlinks
def add_hyperlink(paragraph, url, text, color_rgb="004B87", underline=True, font_size_pt=10, is_highlighted=False):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = parse_xml(f'<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" r:id="{r_id}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>')
    new_run = parse_xml(f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
    rPr = parse_xml(f'<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
    
    rPr.append(parse_xml(f'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Calibri" w:hAnsi="Calibri"/>'))
    val_sz = int(font_size_pt * 2)
    rPr.append(parse_xml(f'<w:sz xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="{val_sz}"/>'))
    rPr.append(parse_xml(f'<w:color xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="{color_rgb}"/>'))
    if underline:
        rPr.append(parse_xml(f'<w:u xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="single"/>'))
    if is_highlighted:
        rPr.append(parse_xml(r'<w:highlight xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="yellow"/>'))
        
    new_run.append(rPr)
    new_run.append(parse_xml(f'<w:t xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{text}</w:t>'))
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

# ==============================================================================
# 3. WORD DOCUMENT GENERATION ENGINE
# ==============================================================================
def create_master_resume_docx(tailored_data, highlight_changes=False):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.40)
        section.bottom_margin = Inches(0.40)
        section.left_margin = Inches(0.50)
        section.right_margin = Inches(0.50)

    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(10)
    style.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    def add_heading(title, space_before=4, space_after=2, line_border=False):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.0
        
        r = p.add_run(title.upper())
        r.bold = True
        r.font.name = 'Calibri'
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        
        if line_border:
            pPr = p._p.get_or_add_pPr()
            pBdr = parse_xml(r'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                             r'<w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/>'
                             r'</w:pBdr>')
            pPr.append(pBdr)

    # ---------------- PAGE 1 ----------------
    # 1. Header Block with Dual Variables (First & Last)
    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(6)
    p_name.paragraph_format.line_spacing = 1.15
    r_name = p_name.add_run(MASTER_STATIC['name'])
    r_name.bold = True
    r_name.font.name = 'Calibri'
    r_name.font.size = Pt(12)

    f1 = tailored_data.get("header_focus_1", "IT & Digital Transformation Director")
    f2 = tailored_data.get("header_focus_2", "Enterprise Sales Technology Leader")
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(6)
    p_sub.paragraph_format.line_spacing = 1.15
    
    r_f1 = p_sub.add_run(f1)
    r_f1.bold = True
    r_f1.font.name = 'Calibri'
    r_f1.font.size = Pt(9)
    if highlight_changes:
        r_f1.font.highlight_color = docx.enum.text.WD_COLOR_INDEX.YELLOW
        
    r_mid = p_sub.add_run(" | FMCG | GTM & Omnichannel Leader | ")
    r_mid.bold = True
    r_mid.font.name = 'Calibri'
    r_mid.font.size = Pt(9)
    
    r_f2 = p_sub.add_run(f2)
    r_f2.bold = True
    r_f2.font.name = 'Calibri'
    r_f2.font.size = Pt(9)
    if highlight_changes:
        r_f2.font.highlight_color = docx.enum.text.WD_COLOR_INDEX.YELLOW

    c = MASTER_STATIC['contact']
    p_contact = doc.add_paragraph()
    p_contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_contact.paragraph_format.space_before = Pt(0)
    p_contact.paragraph_format.space_after = Pt(6)
    p_contact.paragraph_format.line_spacing = 1.15
    
    r_c1 = p_contact.add_run(f"{c['location']} | {c['phone']} | ")
    r_c1.font.name = 'Calibri'
    r_c1.font.size = Pt(10)
    add_hyperlink(p_contact, c['email_url'], c['email'], color_rgb="004B87", underline=True, font_size_pt=10)
    
    r_br1 = p_contact.add_run("\n")
    r_br1.font.name = 'Calibri'
    r_br1.font.size = Pt(10)
    
    add_hyperlink(p_contact, c['linkedin'], c['linkedin'], color_rgb="004B87", underline=True, font_size_pt=10)
    r_c2_mid = p_contact.add_run(" | Portfolio: ")
    r_c2_mid.font.name = 'Calibri'
    r_c2_mid.font.size = Pt(10)
    add_hyperlink(p_contact, c['portfolio'], c['portfolio'], color_rgb="004B87", underline=True, font_size_pt=10)
    
    r_br2 = p_contact.add_run("\n")
    r_br2.font.name = 'Calibri'
    r_br2.font.size = Pt(10)
    
    r_c3_lbl = p_contact.add_run("Visa Status: ")
    r_c3_lbl.bold = True
    r_c3_lbl.font.name = 'Calibri'
    r_c3_lbl.font.size = Pt(10)
    r_c3_val = p_contact.add_run(c['visas'])
    r_c3_val.font.name = 'Calibri'
    r_c3_val.font.size = Pt(10)

    # 2. Executive Summary
    add_heading("EXECUTIVE SUMMARY", space_before=4, space_after=2, line_border=False)
    sp = doc.add_paragraph()
    sp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    sp.paragraph_format.space_before = Pt(0)
    sp.paragraph_format.space_after = Pt(6)
    sp.paragraph_format.line_spacing = 1.15
    r_sum = sp.add_run(tailored_data.get("executive_summary", ""))
    r_sum.font.name = 'Calibri'
    r_sum.font.size = Pt(10)
    if highlight_changes:
        r_sum.font.highlight_color = docx.enum.text.WD_COLOR_INDEX.YELLOW

    # 3. Executive Capabilities & Impact Highlights
    add_heading("EXECUTIVE CAPABILITIES & IMPACT HIGHLIGHTS", space_before=5, space_after=2, line_border=True)
    for cap in tailored_data.get("capabilities", []):
        cp = doc.add_paragraph(style='List Bullet')
        cp.paragraph_format.space_before = Pt(0)
        cp.paragraph_format.space_after = Pt(6)
        cp.paragraph_format.line_spacing = 1.05
        
        pPr = cp._p.get_or_add_pPr()
        pPr.append(parse_xml(r'<w:contextualSpacing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="0"/>'))
        
        parts = cap.split(":", 1)
        if len(parts) == 2:
            r_bold = cp.add_run(parts[0] + ":")
            r_bold.bold = True
            r_bold.font.name = 'Calibri'
            r_bold.font.size = Pt(10)
            if highlight_changes:
                r_bold.font.highlight_color = docx.enum.text.WD_COLOR_INDEX.YELLOW
            r_body = cp.add_run(parts[1])
            r_body.font.name = 'Calibri'
            r_body.font.size = Pt(10)
            if highlight_changes:
                r_body.font.highlight_color = docx.enum.text.WD_COLOR_INDEX.YELLOW
        else:
            r_body = cp.add_run(cap)
            r_body.font.name = 'Calibri'
            r_body.font.size = Pt(10)
            if highlight_changes:
                r_body.font.highlight_color = docx.enum.text.WD_COLOR_INDEX.YELLOW

    # 4. Honors & Recognition
    add_heading("HONORS & RECOGNITION", space_before=4, space_after=2, line_border=False)
    for h in MASTER_STATIC['honors']:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.line_spacing = 1.0
        r_t = p.add_run(h)
        r_t.font.name = 'Calibri'
        r_t.font.size = Pt(10)

    # 5. Education & Languages
    add_heading("EDUCATION", space_before=4, space_after=2, line_border=False)
    for edu in MASTER_STATIC['education']:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.line_spacing = 1.0
        r_bp = p.add_run(edu['degree'] + " – ")
        r_bp.bold = True
        r_bp.font.name = 'Calibri'
        r_bp.font.size = Pt(10)
        r_t = p.add_run(edu['details'])
        r_t.font.name = 'Calibri'
        r_t.font.size = Pt(10)

    add_heading("LANGUAGES & INTERESTS :", space_before=4, space_after=2, line_border=False)
    p_lang1 = doc.add_paragraph()
    p_lang1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_lang1.paragraph_format.space_before = Pt(0)
    p_lang1.paragraph_format.space_after = Pt(2)
    p_lang1.paragraph_format.line_spacing = 1.0
    r_l1 = p_lang1.add_run(MASTER_STATIC['languages'])
    r_l1.font.name = 'Calibri'
    r_l1.font.size = Pt(10)

    p_lang2 = doc.add_paragraph()
    p_lang2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_lang2.paragraph_format.space_before = Pt(0)
    p_lang2.paragraph_format.space_after = Pt(2)
    p_lang2.paragraph_format.line_spacing = 1.0
    r_l2 = p_lang2.add_run(MASTER_STATIC['interests'])
    r_l2.font.name = 'Calibri'
    r_l2.font.size = Pt(10)

    # ---------------- PAGE 2 EXPLICIT BOUNDARY ----------------
    doc.add_page_break()

    # 6. Professional Experience (3-Column Table)
    add_heading("PROFESSIONAL EXPERIENCE", space_before=0, space_after=2, line_border=False)
    
    table = doc.add_table(rows=2, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    col_widths = [Inches(2.5), Inches(2.5), Inches(2.5)]
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            cell.width = col_widths[i]

    hdr_titles = ["Traditional FMCG Operator", "Digital FMCG Distribution", "Distribution Transformation"]
    for i, title in enumerate(hdr_titles):
        cell = table.rows[0].cells[i]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(title)
        r.bold = True
        r.font.name = 'Calibri'
        r.font.size = Pt(11.5)
        tcPr = cell._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(r'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="E9ECEF"/>'))

    def populate_cell_content(cell, item_list):
        cell.text = ""
        for idx, item in enumerate(item_list):
            p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(0 if idx == 0 else item.get("space_before", 0))
            p.paragraph_format.space_after = Pt(item.get("space_after", 1))
            
            if item.get("is_bullet", False):
                p.style = 'List Bullet'
                p.paragraph_format.space_after = Pt(1.2)
                p.paragraph_format.line_spacing = 1.0
                
            r = p.add_run(item["text"])
            r.bold = item.get("bold", False)
            r.italic = item.get("italic", False)
            r.font.name = 'Calibri'
            r.font.size = Pt(item.get("size", 9.5))
            if highlight_changes and item.get("highlight", False):
                r.font.highlight_color = docx.enum.text.WD_COLOR_INDEX.YELLOW

    c0_items = [
        {"text": "Britannia Industries Ltd | 2007 – 2011", "bold": True, "size": 10, "space_before": 1},
        {"text": "Regional Sales Head – GCC", "italic": True, "size": 9.5},
        {"text": "Regional Sales & Capability Head- India", "italic": True, "size": 9.5, "space_after": 2},
        {"text": "Owned $100M+ P&L across GCC (Saudi Arabia, UAE, Kuwait, Oman, Bahrain, Qatar) & South India.", "is_bullet": True, "size": 9.5},
        {"text": "Directed 250+ distributor networks & 600+ frontline sales staff across GT, MT, wholesale, and institutional trade.", "is_bullet": True, "size": 9.5},
        {"text": "Spearheaded Britannia's 1st national SFA rollout (1,000+ users), transforming legacy trade into performance-managed selling.", "is_bullet": True, "size": 9.5},
        {"text": "Delivered ~30% numeric distribution growth, increased LPC to ~120%, and cut sales admin costs by ~30%.", "is_bullet": True, "size": 9.5},
        {"text": "Turnaround RSM GCC: achieved record monthly sales for 3 consecutive months (Best Employee Award from Group MD).", "is_bullet": True, "size": 9.5, "space_after": 3},
        {"text": "Airtel | Reliance | Tyco | 2001 – 2007", "bold": True, "size": 10, "space_before": 2},
        {"text": "Commercial & Training Roles –", "italic": True, "size": 9.5, "space_after": 2},
        {"text": "Built foundations in frontline trade execution, journey planning, and merchandiser enablement in telecom & enterprise security.", "is_bullet": True, "size": 9.5},
        {"text": "Deployed capability training (SPIN selling) & integrated Oracle e-CRM & LMS infrastructure at scale.", "is_bullet": True, "size": 9.5}
    ]

    conektr_cat = tailored_data.get("conektr_category_bullet", "Deep FMCG Category Aggregation: Scaled multi-category catalogs across ambient, packaged food, and consumer goods portfolios.")
    c1_items = [
        {"text": "Digital FMCG Principal / Distributor", "italic": True, "size": 9.5, "space_before": 1},
        {"text": "Chief Executive Officer & Founder", "bold": True, "size": 10},
        {"text": "Conektr Tech Global Ltd | UAE & India", "bold": True, "size": 10},
        {"text": "May 2016 – Aug 2024", "italic": True, "size": 9.5, "space_after": 2},
        {"text": "Founded UAE’s 1st Digital FMCG Principal-Distributor serving 8,000+ retailers (2,000+ MAU) & 100+ brands.", "is_bullet": True, "size": 9.5},
        {"text": conektr_cat, "is_bullet": True, "size": 9.5, "highlight": True},
        {"text": "Owned full P&L, trade terms, warehousing, last-mile delivery, trade credit, and collections.", "is_bullet": True, "size": 9.5},
        {"text": "Built app/web/WhatsApp self-ordering engine scaling annual GMV from zero to ~AED 50M (~$13.6M) at ~18% gross margin.", "is_bullet": True, "size": 9.5},
        {"text": "Cut coverage cost by >50% and improved field execution productivity by ~150% vs traditional trade.", "is_bullet": True, "size": 9.5},
        {"text": "Deployed Dynamics 365 + Power BI and AI route optimization, cutting logistics costs by ~40%.", "is_bullet": True, "size": 9.5},
        {"text": "Raised ~$15M from C-suite FMCG leaders; executed M&A exit to Al Maya Group ($1B+ retail conglomerate).", "is_bullet": True, "size": 9.5}
    ]

    c2_items = [
        {"text": "Post Exit –", "italic": True, "size": 9.5, "space_before": 1},
        {"text": "Transformation Advisor (Director)", "bold": True, "size": 10},
        {"text": "TransCPG Inc. &", "bold": True, "size": 10},
        {"text": "FieldAssist | 2025 – Present", "bold": True, "size": 10, "space_after": 2},
        {"text": "Board Member guiding global operations scaling & platform build across FMCG principals & distributors.", "is_bullet": True, "size": 9.5},
        {"text": "Advising CPG leaders on modernizing RTM & SAP/Oracle SFA/DMS integrations, driving ~150% coverage growth.", "is_bullet": True, "size": 9.5},
        {"text": "Built Bid2Bill AI/Voice-bot & WhatsApp B2B2C bidding platform, cutting CAC by ~40% with 4x engagement.", "is_bullet": True, "size": 9.5, "space_after": 3},
        {"text": "Business Head – MEA", "bold": True, "size": 10, "space_before": 2},
        {"text": "Ivy Mobility Pte Ltd | 2011 – 2016", "bold": True, "size": 10, "space_after": 2},
        {"text": "Built MEA setup from scratch into 2nd largest global setup ($10M+ pipeline across 10+ countries).", "is_bullet": True, "size": 9.5},
        {"text": "Won 22 enterprise logos: Haleon/GSK, P&G, Nestlé, Coca-Cola, Mars, Red Bull, BAT, and AKI Group.", "is_bullet": True, "size": 9.5},
        {"text": "Personally led on-ground field deployment of mobile SFA for P&G distributor networks in Kenya.", "is_bullet": True, "size": 9.5},
        {"text": "Deployed Cloud SaaS SFA/DMS to 3,000+ sales users, driving post-implementation adoption and trade ROI.", "is_bullet": True, "size": 9.5}
    ]

    # --- INJECT CONTEXTUAL JD BULLETS INTO COLUMN 2 OR 3 (MAX 18-24 WORDS EACH) ---
    c1_extra = tailored_data.get("column_2_extra_bullet", "")
    if c1_extra and c1_extra.strip():
        c1_items.insert(7, {"text": c1_extra.strip(), "is_bullet": True, "size": 9.5, "highlight": True})

    c2_extra = tailored_data.get("column_3_extra_bullet", "")
    if c2_extra and c2_extra.strip():
        c2_items.insert(4, {"text": c2_extra.strip(), "is_bullet": True, "size": 9.5, "highlight": True})

    populate_cell_content(table.rows[1].cells[0], c0_items)
    populate_cell_content(table.rows[1].cells[1], c1_items)
    populate_cell_content(table.rows[1].cells[2], c2_items)

    tblBorders = parse_xml(
        r'<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        r'<w:top w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        r'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>'
        r'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>'
        r'</w:tblBorders>'
    )
    table._tbl.tblPr.append(tblBorders)

    # 7. Tech Stack
    add_heading("TECHNOLOGY STACK & DIGITAL ARCHITECTURE:", space_before=8, space_after=3, line_border=False)
    for category, stack in MASTER_STATIC['tech_stack'].items():
        tp = doc.add_paragraph(style='List Bullet')
        tp.paragraph_format.space_before = Pt(0)
        tp.paragraph_format.space_after = Pt(4.5)
        tp.paragraph_format.line_spacing = 1.05
        
        pPr = tp._p.get_or_add_pPr()
        pPr.append(parse_xml(r'<w:contextualSpacing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="0"/>'))
        
        r_cat = tp.add_run(f"{category}: ")
        r_cat.bold = True
        r_cat.font.name = 'Calibri'
        r_cat.font.size = Pt(10)
        r_st = tp.add_run(stack)
        r_st.font.name = 'Calibri'
        r_st.font.size = Pt(10)

    # 8. Why Hire Me
    p_why = doc.add_paragraph()
    p_why.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_why.paragraph_format.space_before = Pt(6)
    p_why.paragraph_format.space_after = Pt(2)
    p_why.paragraph_format.line_spacing = 1.05
    
    r_wh_lbl = p_why.add_run("WHY HIRE ME: ")
    r_wh_lbl.bold = True
    r_wh_lbl.underline = True
    r_wh_lbl.font.name = 'Calibri'
    r_wh_lbl.font.size = Pt(10)
    
    for text_segment, is_plus in MASTER_STATIC['why_hire_me_parts']:
        r_part = p_why.add_run(text_segment)
        r_part.font.name = 'Calibri'
        r_part.font.size = Pt(10)
        if is_plus:
            r_part.bold = True
            r_part.font.color.rgb = RGBColor(0x00, 0xB0, 0xF0)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ==============================================================================
# 4. STREAMLIT FRONTEND & DUAL DOWNLOAD BUTTON ENGINE
# ==============================================================================
st.title("🎯 Executive ATS Resume Tailoring Engine v2")
st.caption("Dual Header Variables • Strict Column Word Budgets • Multi-Column JD Alignment • Locked 2-Page Boundary")

with st.sidebar:
    st.header("⚡ System Status")
    if api_key:
        st.success("🟢 Gemini AI Engine: Active")
    else:
        st.warning("🟠 AI Engine: Inactive (Set GEMINI_API_KEY in Secrets)")
    
    st.markdown("---")
    st.write("📂 **Active Knowledge Archive:**")
    st.caption("• Master Resume (Exact Layout)\n• Contextual Column Injections (Cols 2 & 3)\n• Strict 18-24 Word Limits")

col1, col2 = st.columns([1.1, 0.9])

with col1:
    st.subheader("1. Job Inputs & Specifics")
    job_desc = st.text_area("Target Job Description (JD):", height=240, placeholder="Paste JD here...")
    special_instructions = st.text_area("Special Instructions & Context (Optional):", height=130, 
                                        placeholder="E.g., Target company is Arla Foods, emphasize IT/Digital Transformation over pure sales...")
    
    generate_btn = st.button("🚀 Generate Tailored Master Resumes", type="primary")

if generate_btn:
    if not job_desc:
        st.warning("Please paste a Job Description first.")
    else:
        with col2:
            st.subheader("2. AI Analysis & Tailored Files Ready")
            with st.spinner("Synthesizing JD, Special Instructions & Knowledge Archive..."):
                
                tailored_data = None
                if api_key:
                    prompt = f"""
                    You are an executive resume architect for Madhusudhanan Janakarajan (23+ year FMCG, Digital Transformation & Enterprise Technology Executive).

                    STRICT RULES & CONSTRAINTS:

                    1. HEADER SUBTITLE DUAL VARIABLES:
                       - The complete line format is: "[header_focus_1] | FMCG | GTM & Omnichannel Leader | [header_focus_2]"
                       - "header_focus_1" (First Variable): Target leadership title (e.g., "IT & Digital Transformation Director", "Sales & Distribution Transformation Director", "Commercial & Enterprise Strategy Director"). Max 36 characters.
                       - "header_focus_2" (Last Variable): Matching domain focus (e.g., "Enterprise Sales Technology Leader", "Dairy & Packaged Foods Leadership", "Omnichannel RTM & Digital Execution"). Max 40 characters.
                       - The full line MUST comfortably fit on 1 single line in 9 pt font.

                    2. EXECUTIVE SUMMARY (EXACT 7 TO 8 PRINTED LINES / 135-150 WORDS):
                       - Write a high-impact, authoritative executive summary of EXACTLY 135 to 150 words.
                       - Tailor the opening narrative dynamically to the role:
                         * If IT / Digital / Enterprise Tech JD: Lead with IT/SaaS modernization, digital product architecture, ERP/DMS integration, and bridging IT with sales execution.
                         * If Commercial / Sales / Category JD: Lead with P&L ownership, RTM strategy, distributor governance, and digital commerce.
                       - Retain core metrics ($100M+ P&L, 8,000+ retailers, 10+ Tier-1 CPG logos: P&G, Nestlé, GSK, Coca-Cola, ~40% logistics optimization, ~20% productivity uplifts).

                    3. EXECUTIVE CAPABILITIES (EXACT 5 BULLETS):
                       - Prioritize the 5 master capability themes so the top 2 bullets address the highest priority requirements of the JD.
                       - Maintain all 5 themes formatted as 'Bold Header: Detailed metric description'.

                    4. CONEKTR CATEGORY BULLET:
                       - Synthesize the single best category aggregation bullet (e.g. Dairy & Packaged Foods, Beverages, Personal Care, or Multi-Category Principal Distribution).

                    5. DYNAMIC EXPERIENCE INJECTIONS (STRICT 18 TO 24 WORDS / MAX 3 LINES):
                       - In a 2.5-inch column, each line fits 7-8 words. Therefore, injected points MUST be strictly 18 to 24 words to never exceed 3 lines.
                       - "column_2_extra_bullet" (Under Conektr / Digital FMCG): One sharp, metric-focused point (18-24 words) relating to digital commerce, marketplace apps, or logistics optimization aligned with the JD. Leave empty "" if not relevant.
                       - "column_3_extra_bullet" (Under TransCPG / Ivy / Enterprise Transformation): One sharp, metric-focused point (18-24 words) relating to enterprise IT/SaaS architecture, DMS/ERP rollout, or AI sales automation aligned with the JD. Leave empty "" if not relevant.

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
                      "conektr_category_bullet": "string",
                      "column_2_extra_bullet": "string",
                      "column_3_extra_bullet": "string"
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
                        "header_focus_1": "IT & Digital Transformation Director",
                        "header_focus_2": "Enterprise Sales Technology Leader",
                        "executive_summary": "IT & Digital Transformation Leader with 23+ years driving enterprise sales technology, FMCG commercial strategy, and digital commerce across MENA, India, and Asia. Combines a rare 360° operational vantage across enterprise IT/SaaS modernization, digital product architecture, and principal-led FMCG commercial leadership with $100M+ P&L ownership. Spearheaded multi-country GTM/SFA modernization, Perfect Store automation, and ERP/DMS integrations for 10+ tier-1 CPG enterprises—including P&G, Nestlé, GSK, Coca-Cola, and PepsiCo—consistently delivering ~40% logistics cost optimization and ~20% sales productivity uplifts. Founded and scaled the UAE's premier B2B2C digital distribution platform (Conektr) to an M&A exit, directing end-to-end digital product design, omnichannel ordering engines (App/Web/WhatsApp), and enterprise integrations serving 8,000+ retailers and 100+ global brands. Proven bridge between enterprise IT architecture, frontline sales execution, and C-suite stakeholders, driving regional digitization agendas with high user adoption.",
                        "capabilities": [
                            "Enterprise Transformation & Commercial Optimization: Directed multi-country RTM modernizations, DMS/ERP integrations and SFA deployments (Over 5000+ Users) for global CPG leaders (P&G, Nestlé, Haleon/GSK, Coca-Cola). Deployed AI route/beat optimization, AI-driven demand forecasting, and automated ordering—delivering a ~40% drop in logistics/admin costs, >30% reduction in outlet coverage costs, ~30% frontline sales productivity uplift, ~150% expansion in numeric distribution growth.",
                            "Digital B2B2C Commerce & Omnichannel RTM: Founded and scaled Conektr (UAE's first digital FMCG distributor) to 8,000+ B2B retailers, managing 100+ brands and 2,000+ SKUs across Foods, Beverages, and Non-Food categories. Expanded into direct B2C commerce by launching the consumer app and proprietary BOSS loyalty engine (Buying, Operating, Selling & Saving), turning network grocers into fulfillment micro-hubs/dark stores. Built omnichannel ordering (App, Web, Conversational AI) with fintech-enabled payment rails.",
                            "Commercial & GTM Leadership ($100M+ P&L): Owned $100M+ annual FMCG revenue across GCC & India, directing 250+ distributors and 600+ field sales teams across GT, MT, Wholesale, B2B, and Institutional channels. Spearheaded RTM redesign, distributor governance, trade margin economics, pricing/promotions, and Order-to-Cash optimization.",
                            "Sales Capability & Training Leadership: Established and led regional sales training operations managing a team of certified trainers to design and deliver end-to-end sales induction and leadership curricula up to the Regional Sales Manager (RSM) level. Top-performer in consultative selling frameworks (including SPIN Selling), driving frontline execution rigor, distributor capability building, and institutionalized sales performance standards. Managed Train the Trainer, Soft skills and automation training.",
                            "Entrepreneurial Venture Scaling & Governance: Raised $15M in funding from DIFC VC, veteran FMCG executives (ex-Mondelēz President, BAT CFO) validating commercial credibility, and executed a successful strategic M&A exit to Al Maya Group ($1B+ conglomerate). Awarded UAE Golden Visa and USA O-1A (Extraordinary Ability); featured in Bloomberg, Gulf News, and Magnitt."
                        ],
                        "conektr_category_bullet": "Deep FMCG Category Aggregation: Scaled multi-category catalogs across ambient, packaged food, and consumer goods portfolios.",
                        "column_2_extra_bullet": "Engineered automated micro-fulfillment dark store workflows for urban retailers, reducing replenishment turnaround by 35%.",
                        "column_3_extra_bullet": "Architected seamless API middleware syncing SAP SD and Dynamics 365 with mobile SFA, ensuring 100% order accuracy."
                    }

                docx_clean = create_master_resume_docx(tailored_data, highlight_changes=False)
                docx_highlighted = create_master_resume_docx(tailored_data, highlight_changes=True)
                
                st.success("✅ Tailored Master Resumes Ready!")
                
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    st.download_button(
                        label="📥 Download Clean ATS Resume (.docx)",
                        data=docx_clean,
                        file_name="Madhusudhanan_Janakarajan_Resume_Clean.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                with col_btn2:
                    st.download_button(
                        label="🟡 Download Highlighted Review Resume (.docx)",
                        data=docx_highlighted,
                        file_name="Madhusudhanan_Janakarajan_Resume_Highlighted.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                
                with st.expander("🔍 View AI Tailored Variable Breakdown"):
                    st.write("**Header Variable 1:**", tailored_data.get("header_focus_1"))
                    st.write("**Header Variable 2:**", tailored_data.get("header_focus_2"))
                    st.write("**Executive Summary (135-150 Words):**", tailored_data.get("executive_summary"))
                    st.write("**Column 2 (Conektr) Injected Bullet:**", tailored_data.get("column_2_extra_bullet"))
                    st.write("**Column 3 (TransCPG/Ivy) Injected Bullet:**", tailored_data.get("column_3_extra_bullet"))
                    st.write("**Conektr Category Bullet:**", tailored_data.get("conektr_category_bullet"))
