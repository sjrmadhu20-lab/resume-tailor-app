import streamlit as st
import json
import io
import os
import re
import zipfile
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from google import genai
from google.genai import types

# ReportLab imports for 2-Page Cover Letter + Match Matrix PDF
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

st.set_page_config(page_title="Executive ATS Application Engine", page_icon="🎯", layout="wide")

# ==============================================================================
# 1. API CONFIGURATION
# ==============================================================================
api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))

# ==============================================================================
# 2. MASTER KNOWLEDGE ARCHIVE (LOCKED RESUME DATA)
# ==============================================================================
MASTER_CAPABILITIES = {
    "commercial": "Commercial & GTM Leadership ($100M+ P&L): Owned $100M+ annual FMCG revenue across GCC & India, directing 250+ distributors and 600+ field sales teams across GT, MT, Wholesale, B2B, and Institutional channels. Spearheaded RTM redesign, distributor governance, trade margin economics, pricing/promotions, and Order-to-Cash optimization.",
    "digital": "Digital B2B2C Commerce & Omnichannel RTM: Founded and scaled Conektr (UAE's first digital FMCG distributor) to 8,000+ B2B retailers, managing 100+ brands and 2,000+ SKUs across Foods, Beverages, and Non-Food categories. Expanded into direct B2C commerce by launching the consumer app and proprietary BOSS loyalty engine (Buying, Operating, Selling & Saving), turning network grocers into fulfillment micro-hubs/dark stores. Built omnichannel ordering (App, Web, Conversational AI) with fintech-enabled payment rails.",
    "transformation": "Enterprise Transformation & Commercial Optimization: Directed multi-country RTM modernizations, DMS/ERP integrations and SFA deployments (Over 5000+ Users) for global CPG leaders (P&G, Nestlé, Haleon/GSK, Coca-Cola). Deployed AI route/beat optimization, AI-driven demand forecasting, and automated ordering—delivering a ~40% drop in logistics/admin costs, >30% reduction in outlet coverage costs, ~30% frontline sales productivity uplift, ~150% expansion in numeric distribution growth.",
    "capability": "Sales Capability & Training Leadership: Established and led regional sales training operations managing a team of certified trainers to design and deliver end-to-end sales induction and leadership curricula up to the Regional Sales Manager (RSM) level. Top-performer in consultative selling frameworks (including SPIN Selling), driving frontline execution rigor, distributor capability building, and institutionalized sales performance standards. Managed Train the Trainer, Soft skills and automation training.",
    "entrepreneurship": "Entrepreneurial Venture Scaling & Governance: Raised $15M in funding from DIFC VC, veteran FMCG executives (ex-Mondelēz President, BAT CFO) validating commercial credibility, and executed a successful strategic M&A exit to Al Maya Group ($1B+ conglomerate). Awarded UAE Golden Visa and USA O-1A (Extraordinary Ability); featured in Bloomberg, Gulf News, and Magnitt."
}

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
# 3. WORD RESUME GENERATION ENGINE (LOCKED & PRESERVED)
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

    # 1. Header
    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(6)
    p_name.paragraph_format.line_spacing = 1.15
    r_name = p_name.add_run(MASTER_STATIC['name'])
    r_name.bold = True
    r_name.font.name = 'Calibri'
    r_name.font.size = Pt(12)

    f1 = tailored_data.get("header_focus_1", "Commercial & Digital Transformation Director")
    f2 = tailored_data.get("header_focus_2", "Enterprise Sales & Strategy Leader")
    
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

    # 3. Capabilities
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
            r_body = cp.add_run(parts[1])
            r_body.font.name = 'Calibri'
            r_body.font.size = Pt(10)
        else:
            r_body = cp.add_run(cap)
            r_body.font.name = 'Calibri'
            r_body.font.size = Pt(10)

    # 4. Honors
    add_heading("HONORS & RECOGNITION", space_before=4, space_after=2, line_border=False)
    for h in MASTER_STATIC['honors']:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.line_spacing = 1.0
        r_t = p.add_run(h)
        r_t.font.name = 'Calibri'
        r_t.font.size = Pt(10)

    # 5. Education
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

    # ---------------- PAGE 2 BOUNDARY ----------------
    doc.add_page_break()

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
    return doc_io.getvalue()

# ==============================================================================
# 4. REPORTLAB PDF BUILDER (UPDATED TYPOGRAPHY, SPACING & DIRECT MATRIX)
# ==============================================================================
def create_cover_letter_match_matrix_pdf(cover_data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )

    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('DocTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=13, leading=16, textColor=colors.HexColor('#002B49'), alignment=1, spaceAfter=10)
    subject_style = ParagraphStyle('Subject', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10.5, leading=14, textColor=colors.HexColor('#111827'), spaceBefore=8, spaceAfter=8)
    salutation_style = ParagraphStyle('Salutation', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leading=13.5, textColor=colors.HexColor('#111827'), spaceAfter=8)
    body_style = ParagraphStyle('Body', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leading=14, textColor=colors.HexColor('#1F2937'), alignment=4, spaceAfter=8)
    
    # Centered and indented bullet style
    bullet_style = ParagraphStyle(
        'Bullet', 
        parent=styles['Normal'], 
        fontName='Helvetica', 
        fontSize=9.5, 
        leading=13.5, 
        textColor=colors.HexColor('#1F2937'), 
        leftIndent=24, 
        rightIndent=24, 
        firstLineIndent=-12, 
        spaceAfter=6
    )
    
    sign_style = ParagraphStyle('Sign', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leading=14, textColor=colors.HexColor('#111827'), spaceBefore=8)
    
    th_style = ParagraphStyle('TH', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9.5, leading=12, textColor=colors.HexColor('#002B49'), alignment=1)
    td_left = ParagraphStyle('TDL', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, leading=11.2, textColor=colors.HexColor('#1F2937'))
    td_right = ParagraphStyle('TDR', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, leading=11.2, textColor=colors.HexColor('#1F2937'))

    story = []

    # ---------------- PAGE 1: COVER LETTER ----------------
    story.append(Paragraph("COVER LETTER", title_style))
    story.append(Spacer(1, 4))
    story.append(Paragraph(f"<b>Subject:</b> {cover_data.get('subject_line', 'Application for Target Role')}", subject_style))
    story.append(Spacer(1, 4))
    story.append(Paragraph("Dear Hiring Team,", salutation_style))
    story.append(Spacer(1, 4))
    story.append(Paragraph(cover_data.get("cover_para_1", ""), body_style))
    story.append(Paragraph(cover_data.get("cover_para_2", ""), body_style))
    story.append(Paragraph("<b>Key highlights of what I bring to this mandate include:</b>", body_style))

    for b in cover_data.get("cover_bullets", []):
        story.append(Paragraph(f"• {b}", bullet_style))

    story.append(Spacer(1, 4))
    story.append(Paragraph(cover_data.get("cover_para_closing", ""), body_style))
    story.append(Paragraph("Sincerely,<br/><b>Madhusudhanan Janakarajan (Madhu)</b><br/>+971 50 654 7858 | sjrmadhu20@gmail.com", sign_style))

    # ---------------- PAGE 2: MATCH MATRIX ----------------
    story.append(PageBreak())
    company_name = cover_data.get('target_company', 'TARGET ORGANIZATION').upper()
    story.append(Paragraph("MATCH MATRIX", title_style))
    story.append(Spacer(1, 6))

    matrix_rows = [[
        Paragraph("<b>Target Job Requirement / Focus Domain</b>", th_style),
        Paragraph("<b>How I Match (Evidence & Track Record)</b>", th_style)
    ]]

    for item in cover_data.get("matrix_items", []):
        matrix_rows.append([
            Paragraph(f"<b>{item.get('requirement_title', '')}</b><br/><font color='#4B5563'>{item.get('requirement_desc', '')}</font>", td_left),
            Paragraph(item.get('match_desc', ''), td_right)
        ])

    matrix_table = Table(matrix_rows, colWidths=[2.5 * inch, 4.9 * inch])
    matrix_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E9ECEF')),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D1D5DB')),
        ('TOPPADDING', (0, 0), (-1, -1), 3.5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3.5),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ]))

    story.append(matrix_table)

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# ==============================================================================
# 5. WORD COVER LETTER & MATCH MATRIX BUILDER (.DOCX - CALIBRI 11 BODY)
# ==============================================================================
def create_cover_letter_match_matrix_docx(cover_data):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)
    style.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)

    # ---------------- PAGE 1: COVER LETTER ----------------
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(10)
    r_t = p_title.add_run("COVER LETTER")
    r_t.bold = True
    r_t.font.name = 'Calibri'
    r_t.font.size = Pt(14)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(6)
    p_sub.paragraph_format.space_after = Pt(8)
    r_sb = p_sub.add_run(f"Subject: {cover_data.get('subject_line', '')}")
    r_sb.bold = True
    r_sb.font.name = 'Calibri'
    r_sb.font.size = Pt(11)

    p_d = doc.add_paragraph("Dear Hiring Team,")
    p_d.paragraph_format.space_before = Pt(6)
    p_d.paragraph_format.space_after = Pt(8)

    p_p1 = doc.add_paragraph(cover_data.get("cover_para_1", ""))
    p_p1.paragraph_format.space_after = Pt(8)
    p_p1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    p_p2 = doc.add_paragraph(cover_data.get("cover_para_2", ""))
    p_p2.paragraph_format.space_after = Pt(8)
    p_p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    p_kh = doc.add_paragraph()
    p_kh.paragraph_format.space_after = Pt(6)
    r_kh = p_kh.add_run("Key highlights of what I bring to this mandate include:")
    r_kh.bold = True
    r_kh.font.name = 'Calibri'
    r_kh.font.size = Pt(11)

    # Indented and centered highlights
    for b in cover_data.get("cover_bullets", []):
        bp = doc.add_paragraph(style='List Bullet')
        bp.paragraph_format.left_indent = Inches(0.4)
        bp.paragraph_format.right_indent = Inches(0.4)
        bp.paragraph_format.space_after = Pt(4)
        bp.paragraph_format.line_spacing = 1.1
        r_b = bp.add_run(b)
        r_b.font.name = 'Calibri'
        r_b.font.size = Pt(10.5)

    p_cl = doc.add_paragraph(cover_data.get("cover_para_closing", ""))
    p_cl.paragraph_format.space_before = Pt(6)
    p_cl.paragraph_format.space_after = Pt(8)
    p_cl.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    p_sign = doc.add_paragraph()
    p_sign.paragraph_format.space_before = Pt(8)
    r_s0 = p_sign.add_run("Sincerely,\n")
    r_s0.font.name = 'Calibri'
    r_s1 = p_sign.add_run("Madhusudhanan Janakarajan (Madhu)\n")
    r_s1.bold = True
    r_s1.font.name = 'Calibri'
    r_s2 = p_sign.add_run("+971 50 654 7858 | sjrmadhu20@gmail.com")
    r_s2.font.name = 'Calibri'

    # ---------------- PAGE 2: MATCH MATRIX ----------------
    doc.add_page_break()
    p_mtitle = doc.add_paragraph()
    p_mtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_mtitle.paragraph_format.space_after = Pt(10)
    r_mt = p_mtitle.add_run("MATCH MATRIX")
    r_mt.bold = True
    r_mt.font.name = 'Calibri'
    r_mt.font.size = Pt(14)

    matrix_items = cover_data.get("matrix_items", [])
    table = doc.add_table(rows=len(matrix_items) + 1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.rows[0].cells[0].width = Inches(2.5)
    table.rows[0].cells[1].width = Inches(5.0)

    cell_0 = table.rows[0].cells[0]
    cell_1 = table.rows[0].cells[1]
    
    p_h0 = cell_0.paragraphs[0]
    p_h0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_h0 = p_h0.add_run("Target Job Requirement / Focus Domain")
    r_h0.bold = True
    r_h0.font.name = 'Calibri'
    r_h0.font.size = Pt(10.5)

    p_h1 = cell_1.paragraphs[0]
    p_h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_h1 = p_h1.add_run("How I Match (Evidence & Track Record)")
    r_h1.bold = True
    r_h1.font.name = 'Calibri'
    r_h1.font.size = Pt(10.5)

    # Clean narrative rows without bold/hyphen repetition
    for idx, item in enumerate(matrix_items):
        row_cells = table.rows[idx + 1].cells
        row_cells[0].width = Inches(2.5)
        row_cells[1].width = Inches(5.0)
        
        p0 = row_cells[0].paragraphs[0]
        r_rt = p0.add_run(item.get('requirement_title', '') + "\n")
        r_rt.bold = True
        r_rt.font.name = 'Calibri'
        r_rt.font.size = Pt(10)
        r_rd = p0.add_run(item.get('requirement_desc', ''))
        r_rd.font.name = 'Calibri'
        r_rd.font.size = Pt(9.5)
        
        p1 = row_cells[1].paragraphs[0]
        r_mt = p1.add_run(item.get('match_desc', ''))
        r_mt.font.name = 'Calibri'
        r_mt.font.size = Pt(10)

    tblBorders = parse_xml(
        r'<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        r'<w:top w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        r'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>'
        r'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>'
        r'</w:tblBorders>'
    )
    table._tbl.tblPr.append(tblBorders)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io.getvalue()

def create_full_application_zip(clean_docx, review_docx, matrix_pdf, matrix_docx):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr("1_Madhusudhanan_Janakarajan_Resume_Clean.docx", clean_docx)
        zip_file.writestr("2_Madhusudhanan_Janakarajan_Resume_Highlighted_Review.docx", review_docx)
        zip_file.writestr("3_Cover_Letter_and_Match_Matrix.pdf", matrix_pdf)
        zip_file.writestr("4_Cover_Letter_and_Match_Matrix.docx", matrix_docx)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

# ==============================================================================
# 6. STREAMLIT FRONTEND & ENGINE CONTROLLER
# ==============================================================================
st.title("🎯 Executive ATS Resume & Application Engine")
st.caption("Real-Time AI Tailoring • Exact Master Resume Typography • 2-Page Cover & Matrix PDF • ATS Scoring")

with st.sidebar:
    st.header("⚡ System Status")
    if api_key:
        st.success("🟢 Gemini AI Engine: Active")
    else:
        st.error("🔴 AI Engine Key Missing (Set GEMINI_API_KEY in Secrets)")
    st.markdown("---")
    if st.button("🔄 Reset / Clear Cached Data"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

col1, col2 = st.columns([1.1, 0.9])

with col1:
    st.subheader("1. Job Inputs & Specifics")
    job_desc = st.text_area("Target Job Description (JD):", height=240, placeholder="Paste target Job Description here...")

    st.markdown("##### Special Instructions & Context (Optional)")
    
    st.components.v1.html(
        """
        <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; margin-bottom: 6px;">
            <button id="micBtn" onclick="toggleDictation()" style="
                background-color: #2563EB;
                color: white;
                border: none;
                padding: 7px 14px;
                font-size: 13px;
                font-weight: 600;
                border-radius: 6px;
                cursor: pointer;
                display: inline-flex;
                align-items: center;
                gap: 6px;
            ">🎙️ Click to Speak Instructions</button>
            <span id="status" style="font-size: 12px; color: #4B5563; margin-left: 10px; font-weight: 500;"></span>
        </div>
        <script>
            var recognizing = false;
            var recognition;
            if ('webkitSpeechRecognition' in window || 'SpeechRecognition' in window) {
                var SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
                recognition = new SpeechRecognition();
                recognition.continuous = true;
                recognition.interimResults = false;
                recognition.lang = 'en-US';

                recognition.onstart = function() {
                    recognizing = true;
                    document.getElementById('micBtn').innerText = '🔴 Recording Voice... (Click to Finish)';
                    document.getElementById('micBtn').style.backgroundColor = '#DC2626';
                    document.getElementById('status').innerText = 'Listening... Speak naturally.';
                };

                recognition.onresult = function(event) {
                    var transcript = '';
                    for (var i = event.resultIndex; i < event.results.length; ++i) {
                        transcript += event.results[i][0].transcript + ' ';
                    }
                    var textAreas = window.parent.document.querySelectorAll('textarea');
                    if (textAreas.length > 1) {
                        var target = textAreas[1];
                        var existing = target.value ? target.value.trim() + ' ' : '';
                        target.value = existing + transcript.trim();
                        target.dispatchEvent(new Event('input', { bubbles: true }));
                    }
                };

                recognition.onerror = function(event) {
                    document.getElementById('status').innerText = 'Mic notice: ' + event.error;
                    stopDictation();
                };

                recognition.onend = function() {
                    stopDictation();
                };
            } else {
                document.getElementById('status').innerText = 'Speech recognition not supported in browser.';
            }

            function toggleDictation() {
                if (recognizing) {
                    recognition.stop();
                    stopDictation();
                } else {
                    if (recognition) recognition.start();
                }
            }

            function stopDictation() {
                recognizing = false;
                document.getElementById('micBtn').innerText = '🎙️ Click to Speak Instructions';
                document.getElementById('micBtn').style.backgroundColor = '#2563EB';
                document.getElementById('status').innerText = 'Notes recorded. You can edit or add more text.';
            }
        </script>
        """,
        height=45
    )

    special_instructions = st.text_area(
        "Voice or Typed Notes:",
        height=90,
        placeholder="Click mic above or type instructions here...",
        label_visibility="collapsed"
    )
    
    generate_btn = st.button("🚀 Generate Full Tailored Application Pack", type="primary")

if generate_btn:
    if not job_desc or not job_desc.strip():
        st.warning("Please paste a target Job Description first.")
    elif not api_key:
        st.error("API Key is missing. Please configure GEMINI_API_KEY in Streamlit Secrets.")
    else:
        with col2:
            with st.spinner("Analyzing JD, extracting company context & generating dynamic documents..."):
                prompt = f"""
                You are an executive resume architect and career strategist for Madhusudhanan Janakarajan (23+ year FMCG, Digital Transformation & Enterprise Technology Executive).

                Analyze the provided target Job Description (JD) and special instructions to extract the company name, target role title, and generate fully customized documents.

                STRICT RULES FOR GENERATING JSON:

                1. IDENTIFY TARGET COMPANY & ROLE:
                   - "target_company": The specific company name from the JD.
                   - "target_role": The specific role title from the JD.

                2. HEADER SUBTITLE DUAL VARIABLES:
                   - Format: "[header_focus_1] | FMCG | GTM & Omnichannel Leader | [header_focus_2]"
                   - "header_focus_1": Target leadership title matching the JD (e.g. "E-Commerce & Commercial Director", "Commercial & Digital Transformation Director", "IT & Sales Transformation Director"). Max 36 chars.
                   - "header_focus_2": Specialized domain focus matching the JD (e.g. "DTC & Marketplace Scaling Leader", "Enterprise Sales Technology Leader", "Omnichannel RTM & Digital Execution"). Max 40 chars.

                3. EXECUTIVE SUMMARY (EXACT 135-150 WORDS / 7-8 LINES):
                   - Write an authoritative, rich executive summary of EXACTLY 135 to 150 words dynamically tailored to the role and company.
                   - Emphasize relevant commercial, digital sales, marketplace scaling, and FMCG capabilities directly addressing the JD requirements.
                   - Retain core metrics ($100M+ P&L, 8,000+ retailers, 10+ Tier-1 CPG logos: P&G, Nestlé, GSK, Coca-Cola, ~40% logistics optimization, ~20% productivity uplifts).

                4. CAPABILITY ORDERING (PRIORITIZATION):
                   - Rank the 5 capability keys based on the JD's highest priorities (place the top 2 matching keys first):
                     Available keys: ["commercial", "digital", "transformation", "capability", "entrepreneurship"]
                   - "capability_order": An array containing all 5 keys in ordered priority (e.g., ["digital", "commercial", "transformation", "capability", "entrepreneurship"]).

                5. CONEKTR CATEGORY BULLET:
                   - Category aggregation bullet strictly tailored to the products/domain of the target company.

                6. DYNAMIC EXPERIENCE INJECTIONS (STRICT 18 TO 24 WORDS EACH):
                   - "column_2_extra_bullet": 18-24 words under Conektr (Digital FMCG) aligned with JD, else "".
                   - "column_3_extra_bullet": 18-24 words under TransCPG/Ivy (Transformation) aligned with JD, else "".

                7. ATS MATCH SCORE (INTEGER 88-97):
                   - "ats_match_score": Integer reflecting alignment with the provided JD.

                8. COVER LETTER & MATCH MATRIX:
                   - "subject_line": "Application for [Target Role] - [Target Company]"
                   - "cover_para_1": Authoritative opening explicitly referencing the company name, role title, and candidate's 23+ year track record.
                   - "cover_para_2": Direct alignment with the target company's specific commercial and digital priorities based on JD and special instructions.
                   - "cover_bullets": 4 high-impact bullets with metrics:
                     1) Enterprise IT & SFA/DMS Deployments (P&G, Nestlé, GSK, Britannia).
                     2) 0-to-1 Digital Architecture & Product Leadership (Conektr scaling, 8,000+ outlets, M&A exit).
                     3) Measurable ROI & Executive Buy-in ($15M raised, ~40% logistics savings, ~200% productivity).
                     4) Bridging Commercial & IT Teams ($100M+ P&L, 250+ distributors, high software adoption).
                   - "cover_para_closing": Forward-looking closing paragraph.
                   - "matrix_items": Array of EXACTLY 7 to 8 comprehensive competency rows mapping every major JD pillar to candidate evidence.
                     IMPORTANT: In "match_desc", write a direct, comprehensive narrative paragraph demonstrating evidence WITHOUT adding a redundant bold title or hyphen at the beginning.

                INPUT JOB DESCRIPTION:
                {job_desc}

                INPUT SPECIAL INSTRUCTIONS / CONTEXT:
                {special_instructions}

                Return ONLY a valid JSON object matching this schema:
                {{
                  "target_company": "string",
                  "target_role": "string",
                  "header_focus_1": "string",
                  "header_focus_2": "string",
                  "executive_summary": "string",
                  "capability_order": ["string", "string", "string", "string", "string"],
                  "conektr_category_bullet": "string",
                  "column_2_extra_bullet": "string",
                  "column_3_extra_bullet": "string",
                  "ats_match_score": 94,
                  "cover_letter_data": {{
                    "target_company": "string",
                    "subject_line": "string",
                    "cover_para_1": "string",
                    "cover_para_2": "string",
                    "cover_bullets": ["string", "string", "string", "string"],
                    "cover_para_closing": "string",
                    "matrix_items": [
                      {{
                        "requirement_title": "string",
                        "requirement_desc": "string",
                        "match_desc": "string"
                      }}
                    ]
                  }}
                }}
                """
                
                tailored_data = None
                cover_data = None
                last_error = ""

                model_candidates = [
                    "gemini-3.6-flash",
                    "gemini-3-flash",
                    "gemini-2.0-flash"
                ]

                client = genai.Client(api_key=api_key)
                for model_candidate in model_candidates:
                    try:
                        response = client.models.generate_content(
                            model=model_candidate,
                            contents=prompt,
                            config=types.GenerateContentConfig(
                                response_mime_type="application/json",
                                temperature=0.2
                            )
                        )
                        raw_text = response.text.strip()
                        if raw_text.startswith("```"):
                            raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text)
                            raw_text = re.sub(r"\s*```$", "", raw_text)

                        parsed_json = json.loads(raw_text)
                        
                        ordered_keys = parsed_json.get("capability_order", ["commercial", "digital", "transformation", "capability", "entrepreneurship"])
                        full_capabilities = []
                        for k in ordered_keys:
                            if k in MASTER_CAPABILITIES:
                                full_capabilities.append(MASTER_CAPABILITIES[k])
                        for k, cap_text in MASTER_CAPABILITIES.items():
                            if cap_text not in full_capabilities:
                                full_capabilities.append(cap_text)
                        
                        parsed_json["capabilities"] = full_capabilities
                        tailored_data = parsed_json
                        cover_data = parsed_json.get("cover_letter_data", {})
                        break
                    except Exception as e:
                        last_error = str(e)
                        continue

                if tailored_data:
                    clean_docx = create_master_resume_docx(tailored_data, highlight_changes=False)
                    review_docx = create_master_resume_docx(tailored_data, highlight_changes=True)
                    matrix_pdf = create_cover_letter_match_matrix_pdf(cover_data)
                    matrix_docx = create_cover_letter_match_matrix_docx(cover_data)
                    zip_pack = create_full_application_zip(clean_docx, review_docx, matrix_pdf, matrix_docx)

                    st.session_state["tailored_data"] = tailored_data
                    st.session_state["cover_data"] = cover_data
                    st.session_state["clean_docx"] = clean_docx
                    st.session_state["review_docx"] = review_docx
                    st.session_state["matrix_pdf"] = matrix_pdf
                    st.session_state["matrix_docx"] = matrix_docx
                    st.session_state["zip_pack"] = zip_pack
                    st.session_state["has_results"] = True
                else:
                    st.error(f"Generation Error: {last_error}. Please check your API key and network permissions.")

# ==============================================================================
# 7. PERSISTENT DISPLAY & DOWNLOAD AREA
# ==============================================================================
if st.session_state.get("has_results", False):
    with col2:
        tailored_data = st.session_state["tailored_data"]
        cover_data = st.session_state.get("cover_data", {})
        score = tailored_data.get("ats_match_score", 94)
        target_co = tailored_data.get("target_company", "Target Organization")
        target_rl = tailored_data.get("target_role", "Executive Position")

        st.subheader(f"2. Application Pack: {target_co}")
        st.caption(f"Role: **{target_rl}**")

        st.markdown(f"""
        <div style="
            display: flex;
            align-items: center;
            gap: 16px;
            padding: 12px 16px;
            background: #F0FDF4;
            border: 1px solid #BBF7D0;
            border-radius: 10px;
            margin-bottom: 15px;
        ">
            <div style="
                width: 56px;
                height: 56px;
                border-radius: 50%;
                background: conic-gradient(#16A34A {score * 3.6}deg, #E5E7EB 0deg);
                display: flex;
                align-items: center;
                justify-content: center;
            ">
                <div style="
                    width: 42px;
                    height: 42px;
                    border-radius: 50%;
                    background: white;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    font-weight: bold;
                    color: #16A34A;
                    font-size: 14px;
                ">{score}%</div>
            </div>
            <div>
                <div style="font-weight: 700; font-size: 14.5px; color: #166534;">Target JD Alignment Score: {score}/100</div>
                <div style="font-size: 12px; color: #15803D;">High Match: Tailored precisely to {target_co}'s role specifications.</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.download_button(
            label=f"📦 Download Complete Application Pack (.ZIP) — {target_co}",
            data=st.session_state["zip_pack"],
            file_name=f"Madhusudhanan_Janakarajan_{target_co.replace(' ', '_')}_Application_Pack.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True
        )

        st.markdown("---")
        st.write("📄 **Individual Document Downloads:**")

        col_b1, col_b2 = st.columns(2)
        with col_b1:
            st.download_button(
                label="📥 Clean ATS Resume (.docx)",
                data=st.session_state["clean_docx"],
                file_name=f"Madhusudhanan_Janakarajan_Resume_{target_co.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        with col_b2:
            st.download_button(
                label="🟡 Highlighted Review Resume (.docx)",
                data=st.session_state["review_docx"],
                file_name=f"Madhusudhanan_Janakarajan_Resume_Review_{target_co.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

        col_b3, col_b4 = st.columns(2)
        with col_b3:
            st.download_button(
                label="📄 Cover & Matrix (2-Page PDF)",
                data=st.session_state["matrix_pdf"],
                file_name=f"Madhusudhanan_Janakarajan_CoverLetter_MatchMatrix_{target_co.replace(' ', '_')}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        with col_b4:
            st.download_button(
                label="📝 Cover & Matrix (Word .docx)",
                data=st.session_state["matrix_docx"],
                file_name=f"Madhusudhanan_Janakarajan_CoverLetter_MatchMatrix_{target_co.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

        with st.expander("🔍 View AI Tailored Dynamic Variables"):
            st.write("**Identified Company:**", target_co)
            st.write("**Identified Role:**", target_rl)
            st.write("**Header Focus 1:**", tailored_data.get("header_focus_1"))
            st.write("**Header Focus 2:**", tailored_data.get("header_focus_2"))
            st.write("**Executive Summary:**", tailored_data.get("executive_summary"))
            st.write("**Injected Bullet (Conektr):**", tailored_data.get("column_2_extra_bullet"))
            st.write("**Injected Bullet (TransCPG/Ivy):**", tailored_data.get("column_3_extra_bullet"))
            st.write("**Match Matrix Rows Generated:**", len(cover_data.get("matrix_items", [])))
