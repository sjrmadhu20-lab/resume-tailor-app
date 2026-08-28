import streamlit as st
import json
import io
import os
import re
import time
import zipfile
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from google import genai
from google.genai import types

# ReportLab imports for Pixel-Perfect PDF Generation
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_CENTER
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

st.set_page_config(page_title="Executive ATS Application Engine", page_icon="🎯", layout="wide")

# ==============================================================================
# FONT ENGINE CONFIGURATION (CALIBRI SYSTEM REGISTRATION)
# ==============================================================================
def setup_calibri_fonts():
    font_candidates = [
        # Windows
        ("C:/Windows/Fonts/calibri.ttf", "C:/Windows/Fonts/calibrib.ttf", "C:/Windows/Fonts/calibrii.ttf", "C:/Windows/Fonts/calibriz.ttf"),
        # Linux / Streamlit Cloud / Debian
        ("/usr/share/fonts/truetype/msttcorefonts/calibri.ttf", "/usr/share/fonts/truetype/msttcorefonts/calibrib.ttf", "/usr/share/fonts/truetype/msttcorefonts/calibrii.ttf", "/usr/share/fonts/truetype/msttcorefonts/calibriz.ttf"),
        ("/usr/share/fonts/truetype/calibri/calibri.ttf", "/usr/share/fonts/truetype/calibri/calibrib.ttf", "/usr/share/fonts/truetype/calibri/calibrii.ttf", "/usr/share/fonts/truetype/calibri/calibriz.ttf"),
        # Current Dir
        ("calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf")
    ]
    
    for regular, bold, italic, bolditalic in font_candidates:
        if os.path.exists(regular) and os.path.exists(bold):
            try:
                pdfmetrics.registerFont(TTFont('Calibri', regular))
                pdfmetrics.registerFont(TTFont('Calibri-Bold', bold))
                if os.path.exists(italic):
                    pdfmetrics.registerFont(TTFont('Calibri-Italic', italic))
                else:
                    pdfmetrics.registerFont(TTFont('Calibri-Italic', regular))
                if os.path.exists(bolditalic):
                    pdfmetrics.registerFont(TTFont('Calibri-BoldItalic', bolditalic))
                else:
                    pdfmetrics.registerFont(TTFont('Calibri-BoldItalic', bold))
                return {
                    "regular": "Calibri",
                    "bold": "Calibri-Bold",
                    "italic": "Calibri-Italic",
                    "bolditalic": "Calibri-BoldItalic"
                }
            except Exception:
                continue

    return {
        "regular": "Helvetica",
        "bold": "Helvetica-Bold",
        "italic": "Helvetica-Oblique",
        "bolditalic": "Helvetica-BoldOblique"
    }

FONTS = setup_calibri_fonts()

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
# 3. WORD RESUME ENGINE (UPDATED TO EXACT FORMATTING SPECIFICATIONS)
# ==============================================================================
def populate_resume_document(doc, tailored_data, highlight_changes=False):
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(10)
    style.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    def add_heading(title, space_before=0, space_after=8, line_border_above=False, line_spacing=Pt(12), is_multiple=False):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        if is_multiple:
            p.paragraph_format.line_spacing = 1.16
        else:
            p.paragraph_format.line_spacing = line_spacing
        
        if line_border_above:
            pPr = p._p.get_or_add_pPr()
            pBdr = parse_xml(r'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                             r'<w:top w:val="single" w:sz="6" w:space="1" w:color="000000"/>'
                             r'</w:pBdr>')
            pPr.append(pBdr)
            
        r = p.add_run(title.upper())
        r.bold = True
        r.font.name = 'Calibri'
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    # 1. Header Block (Alignment: Centered, Before: 0pt, After: 8pt, Multiple: 1.16)
    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(0)
    p_name.paragraph_format.line_spacing = 1.16
    r_name = p_name.add_run(MASTER_STATIC['name'])
    r_name.bold = True
    r_name.font.name = 'Calibri'
    r_name.font.size = Pt(12)

    f1 = tailored_data.get("header_focus_1", "Commercial & Digital Transformation Director")
    f2 = tailored_data.get("header_focus_2", "Enterprise Sales & Strategy Leader")
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(8)
    p_sub.paragraph_format.line_spacing = 1.16
    
    r_f1 = p_sub.add_run(f1)
    r_f1.bold = True
    r_f1.font.name = 'Calibri'
    r_f1.font.size = Pt(9.5)
    if highlight_changes:
        r_f1.font.highlight_color = docx.enum.text.WD_COLOR_INDEX.YELLOW
        
    r_mid = p_sub.add_run(" | FMCG | GTM & Omnichannel Leader | ")
    r_mid.bold = True
    r_mid.font.name = 'Calibri'
    r_mid.font.size = Pt(9.5)
    
    r_f2 = p_sub.add_run(f2)
    r_f2.bold = True
    r_f2.font.name = 'Calibri'
    r_f2.font.size = Pt(9.5)
    if highlight_changes:
        r_f2.font.highlight_color = docx.enum.text.WD_COLOR_INDEX.YELLOW

    c = MASTER_STATIC['contact']
    p_contact = doc.add_paragraph()
    p_contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_contact.paragraph_format.space_before = Pt(0)
    p_contact.paragraph_format.space_after = Pt(8)
    p_contact.paragraph_format.line_spacing = 1.16
    
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
    
    r_c3_lbl = p_contact.add_run("Visa Status: ")
    r_c3_lbl.bold = True
    r_c3_lbl.font.name = 'Calibri'
    r_c3_lbl.font.size = Pt(10)
    r_c3_val = p_contact.add_run(c['visas'])
    r_c3_val.font.name = 'Calibri'
    r_c3_val.font.size = Pt(10)

    # 2. Executive Summary (Alignment: Justified, Before: 0pt, After: 8pt, Multiple: 1.16)
    add_heading("EXECUTIVE SUMMARY", space_before=0, space_after=0, line_border_above=False, is_multiple=True)
    sp = doc.add_paragraph()
    sp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    sp.paragraph_format.space_before = Pt(0)
    sp.paragraph_format.space_after = Pt(8)
    sp.paragraph_format.line_spacing = 1.16
    r_sum = sp.add_run(tailored_data.get("executive_summary", ""))
    r_sum.font.name = 'Calibri'
    r_sum.font.size = Pt(10)
    if highlight_changes:
        r_sum.font.highlight_color = docx.enum.text.WD_COLOR_INDEX.YELLOW

    # 3. Capabilities (Line border above, Single line spacing, Left: -0.06", Hanging: 0.25", After: 6pt)
    add_heading("EXECUTIVE CAPABILITIES & IMPACT HIGHLIGHTS", space_before=0, space_after=8, line_border_above=True, line_spacing=Pt(12))
    
    for cap in tailored_data.get("capabilities", []):
        cp = doc.add_paragraph()
        cp.paragraph_format.left_indent = Inches(-0.06)
        cp.paragraph_format.first_line_indent = Inches(-0.25)
        cp.paragraph_format.space_before = Pt(0)
        cp.paragraph_format.space_after = Pt(6)
        cp.paragraph_format.line_spacing = Pt(12)
        cp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        r_bullet = cp.add_run("•\t")
        r_bullet.font.name = 'Calibri'
        r_bullet.font.size = Pt(10)
        
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

    # 4. Honors, Education, Languages (Line border above HONORS, Left: 0.25", Hanging: 0.25", After: 8pt, Single)
    add_heading("HONORS & RECOGNITION", space_before=0, space_after=8, line_border_above=True, line_spacing=Pt(12))
    for h in MASTER_STATIC['honors']:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.first_line_indent = Inches(-0.25)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.line_spacing = Pt(12)
        
        r_t = p.add_run(h)
        r_t.font.name = 'Calibri'
        r_t.font.size = Pt(10)

    add_heading("EDUCATION", space_before=0, space_after=8, line_border_above=False, line_spacing=Pt(12))
    for edu in MASTER_STATIC['education']:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.first_line_indent = Inches(-0.25)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.line_spacing = Pt(12)
        
        r_bp = p.add_run(edu['degree'] + " – ")
        r_bp.bold = True
        r_bp.font.name = 'Calibri'
        r_bp.font.size = Pt(10)
        r_t = p.add_run(edu['details'])
        r_t.font.name = 'Calibri'
        r_t.font.size = Pt(10)

    add_heading("LANGUAGES & INTERESTS :", space_before=0, space_after=8, line_border_above=False, line_spacing=Pt(12))
    p_lang1 = doc.add_paragraph()
    p_lang1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_lang1.paragraph_format.left_indent = Inches(0.25)
    p_lang1.paragraph_format.space_before = Pt(0)
    p_lang1.paragraph_format.space_after = Pt(8)
    p_lang1.paragraph_format.line_spacing = Pt(12)
    r_l1 = p_lang1.add_run(MASTER_STATIC['languages'])
    r_l1.font.name = 'Calibri'
    r_l1.font.size = Pt(10)

    p_lang2 = doc.add_paragraph()
    p_lang2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_lang2.paragraph_format.left_indent = Inches(0.25)
    p_lang2.paragraph_format.space_before = Pt(0)
    p_lang2.paragraph_format.space_after = Pt(8)
    p_lang2.paragraph_format.line_spacing = Pt(12)
    r_l2 = p_lang2.add_run(MASTER_STATIC['interests'])
    r_l2.font.name = 'Calibri'
    r_l2.font.size = Pt(10)

    # ---------------- PAGE 2 BOUNDARY ----------------
    doc.add_page_break()

    # Section Heading: Alignment Justified, Before 0pt, After 8pt, Single
    add_heading("PROFESSIONAL EXPERIENCE", space_before=0, space_after=8, line_border_above=False, line_spacing=Pt(12))
    
    # 3-Column Experience Table
    # Preferred width: 8.0", Indent from left: -0.19", Column width: 2.63" each, Vertical alignment: Top
    table = doc.add_table(rows=2, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    
    tblPr = table._tbl.tblPr
    tblpPr = parse_xml(r'<w:tblpPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:tblpX="-274" w:tblpY="0"/>')
    tblPr.append(tblpPr)
    
    col_widths = [Inches(2.63), Inches(2.63), Inches(2.63)]
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            cell.width = col_widths[i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    hdr_titles = ["Traditional FMCG Operator", "Digital FMCG Distribution", "Distribution Transformation"]
    for i, title in enumerate(hdr_titles):
        cell = table.rows[0].cells[i]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = Pt(11)
        r = p.add_run(title)
        r.bold = True
        r.font.name = 'Calibri'
        r.font.size = Pt(10)
        tcPr = cell._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(r'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="E9ECEF"/>'))

    def populate_cell_content(cell, item_list):
        cell.text = ""
        for idx, item in enumerate(item_list):
            p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(item.get("space_before", 0))
            p.paragraph_format.space_after = Pt(item.get("space_after", 1.5))
            p.paragraph_format.line_spacing = Pt(11)
            
            if item.get("is_bullet", False):
                p.paragraph_format.left_indent = Inches(0.15)
                p.paragraph_format.first_line_indent = Inches(-0.15)
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r_b = p.add_run("• ")
                r_b.font.name = 'Calibri'
                r_b.font.size = Pt(item.get("size", 9.5))
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                
            r = p.add_run(item["text"])
            r.bold = item.get("bold", False)
            r.italic = item.get("italic", False)
            r.font.name = 'Calibri'
            r.font.size = Pt(item.get("size", 9.5))
            if highlight_changes and item.get("highlight", False):
                r.font.highlight_color = docx.enum.text.WD_COLOR_INDEX.YELLOW

    c0_items = [
        {"text": "Britannia Industries Ltd | 2007 – 2011", "bold": True, "size": 9.5, "space_before": 2},
        {"text": "Regional Sales Head – GCC", "italic": True, "size": 9.5},
        {"text": "Regional Sales & Capability Head- India", "italic": True, "size": 9.5, "space_after": 3},
        {"text": "Owned $100M+ P&L across GCC (Saudi Arabia, UAE, Kuwait, Oman, Bahrain, Qatar) & South India.", "is_bullet": True, "size": 9.5},
        {"text": "Directed 250+ distributor networks & 600+ frontline sales staff across GT, MT, wholesale, and institutional trade.", "is_bullet": True, "size": 9.5},
        {"text": "Spearheaded Britannia's 1st national SFA rollout (1,000+ users), transforming legacy trade into performance-managed selling.", "is_bullet": True, "size": 9.5},
        {"text": "Delivered ~30% numeric distribution growth, increased LPC to ~120%, and cut sales admin costs by ~30%.", "is_bullet": True, "size": 9.5},
        {"text": "Turnaround RSM GCC: achieved record monthly sales for 3 consecutive months (Best Employee Award from Group MD).", "is_bullet": True, "size": 9.5, "space_after": 4},
        {"text": "Airtel | Reliance | Tyco | 2001 – 2007", "bold": True, "size": 9.5, "space_before": 3},
        {"text": "Commercial & Training Roles –", "italic": True, "size": 9.5, "space_after": 2},
        {"text": "Built foundations in frontline trade execution, journey planning, and merchandiser enablement in telecom & enterprise security.", "is_bullet": True, "size": 9.5},
        {"text": "Deployed capability training (SPIN selling) & integrated Oracle e-CRM & LMS infrastructure at scale.", "is_bullet": True, "size": 9.5}
    ]

    # Clean multi-line structure for Column 2
    conektr_cat = tailored_data.get("conektr_category_bullet", "Deep FMCG Category Aggregation: Scaled multi-category catalogs across ambient, packaged food, and consumer goods portfolios.")
    c1_items = [
        {"text": "Digital FMCG Principal / Distributor", "italic": True, "size": 9.5, "space_before": 2},
        {"text": "Chief Executive Officer & Founder", "bold": True, "size": 9.5},
        {"text": "Conektr Tech Global Ltd | UAE & India", "bold": True, "size": 9.5},
        {"text": "May 2016 – Aug 2024", "italic": True, "size": 9.5, "space_after": 3},
        {"text": "Founded UAE’s 1st Digital FMCG Principal-Distributor serving 8,000+ retailers (2,000+ MAU) & 100+ brands.", "is_bullet": True, "size": 9.5},
        {"text": conektr_cat, "is_bullet": True, "size": 9.5, "highlight": True},
        {"text": "Owned full P&L, trade terms, warehousing, last-mile delivery, trade credit, and collections.", "is_bullet": True, "size": 9.5},
        {"text": "Built app/web/WhatsApp self-ordering engine scaling annual GMV from zero to ~AED 50M (~$13.6M) at ~18% gross margin.", "is_bullet": True, "size": 9.5},
        {"text": "Cut coverage cost by >50% and improved field execution productivity by ~150% vs traditional trade.", "is_bullet": True, "size": 9.5},
        {"text": "Deployed Dynamics 365 + Power BI and AI route optimization, cutting logistics costs by ~40%.", "is_bullet": True, "size": 9.5},
        {"text": "Raised ~$15M from C-suite FMCG leaders; executed M&A exit to Al Maya Group ($1B+ retail conglomerate).", "is_bullet": True, "size": 9.5}
    ]

    # Clean multi-line structure for Column 3
    c2_items = [
        {"text": "Post Exit –", "italic": True, "size": 9.5, "space_before": 2},
        {"text": "Transformation Advisor (Director)", "bold": True, "size": 9.5},
        {"text": "TransCPG Inc. &", "bold": True, "size": 9.5},
        {"text": "FieldAssist | 2025 – Present", "bold": True, "size": 9.5, "space_after": 3},
        {"text": "Board Member guiding global operations scaling & platform build across FMCG principals & distributors.", "is_bullet": True, "size": 9.5},
        {"text": "Advising CPG leaders on modernizing RTM & SAP/Oracle SFA/DMS integrations, driving ~150% coverage growth.", "is_bullet": True, "size": 9.5},
        {"text": "Built Bid2Bill AI/Voice-bot & WhatsApp B2B2C bidding platform, cutting CAC by ~40% with 4x engagement.", "is_bullet": True, "size": 9.5, "space_after": 4},
        {"text": "Business Head – MEA", "bold": True, "size": 9.5, "space_before": 3},
        {"text": "Ivy Mobility Pte Ltd | 2011 – 2016", "bold": True, "size": 9.5, "space_after": 2},
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

    # 7. Tech Stack (Heading: Justified, After: 8pt, Single; Items: Left: -0.06", Hanging: 0.25", After: 6pt, Single)
    add_heading("TECHNOLOGY STACK & DIGITAL ARCHITECTURE:", space_before=0, space_after=8, line_border_above=False, line_spacing=Pt(12))
    for category, stack in MASTER_STATIC['tech_stack'].items():
        tp = doc.add_paragraph()
        tp.paragraph_format.left_indent = Inches(-0.06)
        tp.paragraph_format.first_line_indent = Inches(-0.25)
        tp.paragraph_format.space_before = Pt(0)
        tp.paragraph_format.space_after = Pt(6)
        tp.paragraph_format.line_spacing = Pt(12)
        tp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        r_b = tp.add_run("•\t")
        r_b.font.name = 'Calibri'
        r_b.font.size = Pt(10)
        
        r_cat = tp.add_run(f"{category}: ")
        r_cat.bold = True
        r_cat.font.name = 'Calibri'
        r_cat.font.size = Pt(10)
        
        r_st = tp.add_run(stack)
        r_st.font.name = 'Calibri'
        r_st.font.size = Pt(10)

    # 8. Why Hire Me (Heading & Content Split, Justified, After 8pt, Single, '+' in Blue)
    add_heading("WHY HIRE ME", space_before=0, space_after=2, line_border_above=False, line_spacing=Pt(12))
    
    p_why = doc.add_paragraph()
    p_why.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_why.paragraph_format.space_before = Pt(0)
    p_why.paragraph_format.space_after = Pt(8)
    p_why.paragraph_format.line_spacing = Pt(12)
    
    for text_segment, is_plus in MASTER_STATIC['why_hire_me_parts']:
        r_part = p_why.add_run(text_segment)
        r_part.font.name = 'Calibri'
        r_part.font.size = Pt(10)
        if is_plus:
            r_part.bold = True
            r_part.font.color.rgb = RGBColor(0x00, 0xB0, 0xF0)

def create_master_resume_docx(tailored_data, highlight_changes=False):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.40)
        section.bottom_margin = Inches(0.40)
        section.left_margin = Inches(0.50)
        section.right_margin = Inches(0.50)
    populate_resume_document(doc, tailored_data, highlight_changes)
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io.getvalue()

# ==============================================================================
# 4. REPORTLAB ELEMENTS GENERATOR (PAGE-FILLING CALIBRI RESUME & SUITE)
# ==============================================================================
def get_pdf_styles():
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('DocTitle', parent=styles['Normal'], fontName=FONTS['bold'], fontSize=13, leading=16, textColor=colors.HexColor('#002B49'), alignment=TA_CENTER, spaceAfter=8)
    subject_style = ParagraphStyle('Subject', parent=styles['Normal'], fontName=FONTS['bold'], fontSize=10, leading=14, textColor=colors.HexColor('#111827'), spaceBefore=5, spaceAfter=5)
    salutation_style = ParagraphStyle('Salutation', parent=styles['Normal'], fontName=FONTS['regular'], fontSize=10, leading=14, textColor=colors.HexColor('#111827'), spaceAfter=5)
    body_style = ParagraphStyle('Body', parent=styles['Normal'], fontName=FONTS['regular'], fontSize=10, leading=14, textColor=colors.HexColor('#1F2937'), alignment=TA_JUSTIFY, spaceAfter=6)
    bullet_style = ParagraphStyle('Bullet', parent=styles['Normal'], fontName=FONTS['regular'], fontSize=10, leading=14, textColor=colors.HexColor('#1F2937'), alignment=TA_JUSTIFY, leftIndent=18, rightIndent=6, firstLineIndent=-12, spaceAfter=5)
    sign_style = ParagraphStyle('Sign', parent=styles['Normal'], fontName=FONTS['regular'], fontSize=10, leading=14, textColor=colors.HexColor('#111827'), spaceBefore=6)
    
    th_style = ParagraphStyle('TH', parent=styles['Normal'], fontName=FONTS['bold'], fontSize=10, leading=12.5, textColor=colors.HexColor('#002B49'), alignment=TA_CENTER)
    td_left = ParagraphStyle('TDL', parent=styles['Normal'], fontName=FONTS['bold'], fontSize=9.5, leading=12.5, textColor=colors.HexColor('#111827'), alignment=TA_LEFT)
    td_right = ParagraphStyle('TDR', parent=styles['Normal'], fontName=FONTS['regular'], fontSize=9.5, leading=12.5, textColor=colors.HexColor('#1F2937'), alignment=TA_JUSTIFY)
    
    # 1:1 Matched Full-Page Resume Typography
    r_name_style = ParagraphStyle('RName', parent=styles['Normal'], fontName=FONTS['bold'], fontSize=12, leading=14, alignment=TA_CENTER, spaceAfter=2)
    r_sub_style = ParagraphStyle('RSub', parent=styles['Normal'], fontName=FONTS['bold'], fontSize=9.5, leading=12, alignment=TA_CENTER, spaceAfter=4, textColor=colors.HexColor('#111827'))
    r_contact_style = ParagraphStyle('RCont', parent=styles['Normal'], fontName=FONTS['regular'], fontSize=10, leading=13, alignment=TA_CENTER, spaceAfter=6, textColor=colors.HexColor('#374151'))
    
    r_h_style = ParagraphStyle('RH', parent=styles['Normal'], fontName=FONTS['bold'], fontSize=10, leading=12.5, spaceBefore=4, spaceAfter=3, textColor=colors.HexColor('#000000'), alignment=TA_JUSTIFY)
    r_body_style = ParagraphStyle('RBody', parent=styles['Normal'], fontName=FONTS['regular'], fontSize=10, leading=13, alignment=TA_JUSTIFY, spaceAfter=5, textColor=colors.HexColor('#1F2937'))
    r_bullet_style = ParagraphStyle('RBul', parent=styles['Normal'], fontName=FONTS['regular'], fontSize=10, leading=12.5, alignment=TA_LEFT, leftIndent=12, firstLineIndent=-12, spaceAfter=4, textColor=colors.HexColor('#1F2937'))
    
    col_hdr_style = ParagraphStyle('ColHdr', parent=styles['Normal'], fontName=FONTS['bold'], fontSize=10, leading=12, alignment=TA_LEFT, textColor=colors.HexColor('#002B49'))
    col_cell_style = ParagraphStyle('ColCell', parent=styles['Normal'], fontName=FONTS['regular'], fontSize=9.5, leading=11.5, alignment=TA_LEFT, textColor=colors.HexColor('#111827'))

    return {
        "title": title_style, "subject": subject_style, "salutation": salutation_style,
        "body": body_style, "bullet": bullet_style, "sign": sign_style,
        "th": th_style, "td_left": td_left, "td_right": td_right,
        "r_name": r_name_style, "r_sub": r_sub_style, "r_contact": r_contact_style,
        "r_h": r_h_style, "r_body": r_body_style, "r_bullet": r_bullet_style,
        "col_hdr": col_hdr_style, "col_cell": col_cell_style
    }

def get_cover_letter_story(cover_data, st_dict):
    story = [
        Paragraph("COVER LETTER", st_dict["title"]),
        Spacer(1, 4),
        Paragraph(f"<b>Subject:</b> {cover_data.get('subject_line', 'Application for Target Role')}", st_dict["subject"]),
        Spacer(1, 4),
        Paragraph("Dear Hiring Team,", st_dict["salutation"]),
        Spacer(1, 4),
        Paragraph(cover_data.get("cover_para_1", ""), st_dict["body"]),
        Paragraph(cover_data.get("cover_para_2", ""), st_dict["body"]),
        Paragraph("<b>Key highlights of what I bring to this mandate include:</b>", st_dict["body"])
    ]
    for b in cover_data.get("cover_bullets", []):
        story.append(Paragraph(f"• &nbsp; {b}", st_dict["bullet"]))
    story.append(Spacer(1, 4))
    story.append(Paragraph(cover_data.get("cover_para_closing", ""), st_dict["body"]))
    story.append(Paragraph("Sincerely,<br/><b>Madhusudhanan Janakarajan (Madhu)</b><br/>+971 50 654 7858 | sjrmadhu20@gmail.com", st_dict["sign"]))
    return story

def get_match_matrix_story(cover_data, st_dict):
    story = [
        Paragraph("MATCH MATRIX", st_dict["title"]),
        Spacer(1, 4)
    ]
    matrix_rows = [[
        Paragraph("<b>Target Job Requirement / Focus Domain</b>", st_dict["th"]),
        Paragraph("<b>How I Match (Evidence & Track Record)</b>", st_dict["th"])
    ]]
    for item in cover_data.get("matrix_items", []):
        matrix_rows.append([
            Paragraph(item.get('requirement_title', ''), st_dict["td_left"]),
            Paragraph(item.get('match_desc', ''), st_dict["td_right"])
        ])
    matrix_table = Table(matrix_rows, colWidths=[2.5 * inch, 5.0 * inch])
    matrix_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E9ECEF')),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D1D5DB')),
        ('TOPPADDING', (0, 0), (-1, -1), 3.5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3.5),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
    ]))
    story.append(matrix_table)
    return story

def get_resume_story(tailored_data, st_dict):
    story = []
    
    # ---------------- PAGE 1 ----------------
    story.append(Paragraph(MASTER_STATIC['name'], st_dict["r_name"]))
    f1 = tailored_data.get("header_focus_1", "Commercial & Digital Transformation Director")
    f2 = tailored_data.get("header_focus_2", "Enterprise Sales & Strategy Leader")
    sub_line = f"{f1} | FMCG | GTM & Omnichannel Leader | {f2}"
    story.append(Paragraph(sub_line, st_dict["r_sub"]))
    c = MASTER_STATIC['contact']
    contact_line = f"{c['location']} | {c['phone']} | <a href='{c['email_url']}'><font color='#004B87'><u>{c['email']}</u></font></a><br/><a href='{c['linkedin']}'><font color='#004B87'><u>{c['linkedin']}</u></font></a> | Portfolio: <a href='{c['portfolio']}'><font color='#004B87'><u>{c['portfolio']}</u></font></a><br/><b>Visa Status:</b> {c['visas']}"
    story.append(Paragraph(contact_line, st_dict["r_contact"]))
    
    story.append(Paragraph("<b>EXECUTIVE SUMMARY</b>", st_dict["r_h"]))
    story.append(Paragraph(tailored_data.get("executive_summary", ""), st_dict["r_body"]))
    
    story.append(Spacer(1, 2))
    story.append(HRFlowable(width="100%", thickness=0.8, color=colors.HexColor('#000000'), spaceBefore=2, spaceAfter=3))
    story.append(Paragraph("<b>EXECUTIVE CAPABILITIES & IMPACT HIGHLIGHTS</b>", st_dict["r_h"]))
    
    for cap in tailored_data.get("capabilities", []):
        parts = cap.split(":", 1)
        if len(parts) == 2:
            story.append(Paragraph(f"• <b>{parts[0]}:</b>{parts[1]}", st_dict["r_bullet"]))
        else:
            story.append(Paragraph(f"• {cap}", st_dict["r_bullet"]))
            
    story.append(Spacer(1, 2))
    story.append(HRFlowable(width="100%", thickness=0.8, color=colors.HexColor('#000000'), spaceBefore=2, spaceAfter=3))
    story.append(Paragraph("<b>HONORS & RECOGNITION</b>", st_dict["r_h"]))
    for h in MASTER_STATIC['honors']:
        if "https://" in h:
            parts = h.split(" - ")
            url = parts[1]
            story.append(Paragraph(f"• {parts[0]} - <a href='{url}'><font color='#004B87'><u>{url}</u></font></a>", st_dict["r_bullet"]))
        else:
            story.append(Paragraph(f"• {h}", st_dict["r_bullet"]))
        
    story.append(Spacer(1, 2))
    story.append(Paragraph("<b>EDUCATION</b>", st_dict["r_h"]))
    for edu in MASTER_STATIC['education']:
        story.append(Paragraph(f"• <b>{edu['degree']}</b> – {edu['details']}", st_dict["r_bullet"]))
        
    story.append(Spacer(1, 2))
    story.append(Paragraph("<b>LANGUAGES & INTERESTS :</b>", st_dict["r_h"]))
    story.append(Paragraph(MASTER_STATIC['languages'], st_dict["r_body"]))
    story.append(Paragraph(MASTER_STATIC['interests'], st_dict["r_body"]))
    
    # ---------------- PAGE 2 BOUNDARY ----------------
    story.append(PageBreak())
    story.append(Paragraph("<b>PROFESSIONAL EXPERIENCE</b>", st_dict["r_h"]))
    story.append(Spacer(1, 2))
    
    c1_txt = (
        "<b>Britannia Industries Ltd | 2007 – 2011</b><br/>"
        "<i>Regional Sales Head – GCC</i><br/>"
        "<i>Regional Sales & Capability Head- India</i><br/>"
        "• Owned $100M+ P&L across GCC (Saudi Arabia, UAE, Kuwait, Oman, Bahrain, Qatar) & South India.<br/>"
        "• Directed 250+ distributor networks & 600+ frontline sales staff across GT, MT, wholesale, and institutional trade.<br/>"
        "• Spearheaded Britannia's 1st national SFA rollout (1,000+ users), transforming legacy trade into performance-managed selling.<br/>"
        "• Delivered ~30% numeric distribution growth, increased LPC to ~120%, and cut sales admin costs by ~30%.<br/>"
        "• Turnaround RSM GCC: achieved record monthly sales for 3 consecutive months (Best Employee Award from Group MD).<br/><br/>"
        "<b>Airtel | Reliance | Tyco | 2001 – 2007</b><br/>"
        "<i>Commercial & Training Roles –</i><br/>"
        "• Built foundations in frontline trade execution, journey planning, and merchandiser enablement in telecom & enterprise security.<br/>"
        "• Deployed capability training (SPIN selling) & integrated Oracle e-CRM & LMS infrastructure at scale."
    )
    
    conektr_cat = tailored_data.get("conektr_category_bullet", "Deep FMCG Category Aggregation: Scaled multi-category catalogs across ambient, packaged food, and consumer goods portfolios.")
    c1_dyn = tailored_data.get("column_2_extra_bullet", "")
    c2_txt = (
        "<i>Digital FMCG Principal / Distributor</i><br/>"
        "<b>Chief Executive Officer & Founder</b><br/>"
        "<b>Conektr Tech Global Ltd | UAE & India</b><br/>"
        "<i>May 2016 – Aug 2024</i><br/>"
        "• Founded UAE's 1st Digital FMCG Principal-Distributor serving 8,000+ retailers (2,000+ MAU) & 100+ brands.<br/>"
        f"• {conektr_cat}<br/>"
        "• Owned full P&L, trade terms, warehousing, last-mile delivery, trade credit, and collections.<br/>"
        "• Built app/web/WhatsApp self-ordering engine scaling annual GMV from zero to ~AED 50M (~$13.6M) at ~18% gross margin.<br/>"
    )
    if c1_dyn:
        c2_txt += f"• {c1_dyn}<br/>"
    c2_txt += (
        "• Cut coverage cost by >50% and improved field execution productivity by ~150% vs traditional trade.<br/>"
        "• Deployed Dynamics 365 + Power BI and AI route optimization, cutting logistics costs by ~40%.<br/>"
        "• Raised ~$15M from C-suite FMCG leaders; executed M&A exit to Al Maya Group ($1B+ retail conglomerate)."
    )
    
    c2_dyn = tailored_data.get("column_3_extra_bullet", "")
    c3_txt = (
        "<i>Post Exit –</i><br/>"
        "<b>Transformation Advisor (Director)</b><br/>"
        "<b>TransCPG Inc. & FieldAssist</b><br/>"
        "<b>2025 – Present</b><br/>"
        "• Board Member guiding global operations scaling & platform build across FMCG principals & distributors.<br/>"
        "• Advising CPG leaders on modernizing RTM & SAP/Oracle SFA/DMS integrations, driving ~150% coverage growth.<br/>"
    )
    if c2_dyn:
        c3_txt += f"• {c2_dyn}<br/>"
    c3_txt += (
        "• Built Bid2Bill AI/Voice-bot & WhatsApp B2B2C bidding platform, cutting CAC by ~40% with 4x engagement.<br/><br/>"
        "<b>Business Head – MEA</b><br/>"
        "<b>Ivy Mobility Pte Ltd | 2011 – 2016</b><br/>"
        "• Built MEA setup from scratch into 2nd largest global setup ($10M+ pipeline across 10+ countries).<br/>"
        "• Won 22 enterprise logos: Haleon/GSK, P&G, Nestlé, Coca-Cola, Mars, Red Bull, BAT, and AKI Group.<br/>"
        "• Personally led on-ground field deployment of mobile SFA for P&G distributor networks in Kenya.<br/>"
        "• Deployed Cloud SaaS SFA/DMS to 3,000+ sales users, driving post-implementation adoption and trade ROI."
    )
    
    exp_table_data = [
        [
            Paragraph("<b>Traditional FMCG Operator</b>", st_dict["col_hdr"]),
            Paragraph("<b>Digital FMCG Distribution</b>", st_dict["col_hdr"]),
            Paragraph("<b>Distribution Transformation</b>", st_dict["col_hdr"])
        ],
        [
            Paragraph(c1_txt, st_dict["col_cell"]),
            Paragraph(c2_txt, st_dict["col_cell"]),
            Paragraph(c3_txt, st_dict["col_cell"])
        ]
    ]
    
    exp_table = Table(exp_table_data, colWidths=[2.63 * inch, 2.63 * inch, 2.63 * inch])
    exp_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E9ECEF')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D1D5DB')),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(exp_table)
    
    story.append(Spacer(1, 3))
    story.append(Paragraph("<b>TECHNOLOGY STACK & DIGITAL ARCHITECTURE:</b>", st_dict["r_h"]))
    for category, stack in MASTER_STATIC['tech_stack'].items():
        story.append(Paragraph(f"• <b>{category}:</b> {stack}", st_dict["r_bullet"]))
        
    story.append(Spacer(1, 2.5))
    story.append(Paragraph("<b>WHY HIRE ME</b>", st_dict["r_h"]))
    why_text = (
        "A rare profile combining Core FMCG Operator <font color='#00B0F0'><b>+</b></font> "
        "Digital FMCG Disruption pioneer <font color='#00B0F0'><b>+</b></font> "
        "Enterprise Transformations (P&G, Coca-cola, GSK) <font color='#00B0F0'><b>+</b></font> "
        "10+ International Markets (GCC, India, Africa, Asia) <font color='#00B0F0'><b>+</b></font> "
        "Successful Entrepreneurial $15M M&A Exit <font color='#00B0F0'><b>+</b></font> "
        "Recipient of Global recognition for FMCG Contribution: O1A from USA & Golden Visa from UAE - as an extraordinary ability leader."
    )
    story.append(Paragraph(why_text, st_dict["r_body"]))
    return story

def create_cover_letter_match_matrix_pdf(cover_data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=28, bottomMargin=28)
    st_dict = get_pdf_styles()
    story = get_cover_letter_story(cover_data, st_dict) + [PageBreak()] + get_match_matrix_story(cover_data, st_dict)
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

def create_resume_pdf(tailored_data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=22, bottomMargin=22)
    st_dict = get_pdf_styles()
    story = get_resume_story(tailored_data, st_dict)
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

def create_combined_application_pdf(cover_data, tailored_data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=22, bottomMargin=22)
    st_dict = get_pdf_styles()
    story = (
        get_cover_letter_story(cover_data, st_dict) +
        [PageBreak()] +
        get_resume_story(tailored_data, st_dict) +
        [PageBreak()] +
        get_match_matrix_story(cover_data, st_dict)
    )
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# ==============================================================================
# 5. WORD COVER, MATRIX & COMBINED PACK BUILDER (.DOCX)
# ==============================================================================
def populate_cover_letter_docx_page(doc, cover_data):
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

    for b in cover_data.get("cover_bullets", []):
        bp = doc.add_paragraph(style='List Bullet')
        bp.paragraph_format.left_indent = Inches(0.5)
        bp.paragraph_format.right_indent = Inches(0.2)
        bp.paragraph_format.space_after = Pt(6)
        bp.paragraph_format.line_spacing = 1.15
        bp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        parts = b.split(":", 1)
        if len(parts) == 2:
            r_head = bp.add_run(parts[0] + ": ")
            r_head.bold = True
            r_head.font.name = 'Calibri'
            r_head.font.size = Pt(10.5)
            r_tail = bp.add_run(parts[1].strip())
            r_tail.font.name = 'Calibri'
            r_tail.font.size = Pt(10.5)
        else:
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

def populate_match_matrix_docx_page(doc, cover_data):
    p_mtitle = doc.add_paragraph()
    p_mtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_mtitle.paragraph_format.space_after = Pt(8)
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

    for idx, item in enumerate(matrix_items):
        row_cells = table.rows[idx + 1].cells
        row_cells[0].width = Inches(2.5)
        row_cells[1].width = Inches(5.0)
        
        p0 = row_cells[0].paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p0.paragraph_format.space_before = Pt(3)
        p0.paragraph_format.space_after = Pt(3)
        r_rt = p0.add_run(item.get('requirement_title', ''))
        r_rt.bold = True
        r_rt.font.name = 'Calibri'
        r_rt.font.size = Pt(9.5)
        
        p1 = row_cells[1].paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p1.paragraph_format.space_before = Pt(3)
        p1.paragraph_format.space_after = Pt(3)
        r_mt = p1.add_run(item.get('match_desc', ''))
        r_mt.font.name = 'Calibri'
        r_mt.font.size = Pt(9.5)

    tblBorders = parse_xml(
        r'<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        r'<w:top w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        r'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>'
        r'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>'
        r'</w:tblBorders>'
    )
    table._tbl.tblPr.append(tblBorders)

def create_cover_letter_match_matrix_docx(cover_data):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.45)
        section.bottom_margin = Inches(0.45)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
    populate_cover_letter_docx_page(doc, cover_data)
    doc.add_page_break()
    populate_match_matrix_docx_page(doc, cover_data)
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io.getvalue()

def create_combined_application_docx(cover_data, tailored_data):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.40)
        section.bottom_margin = Inches(0.40)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
    
    populate_cover_letter_docx_page(doc, cover_data)
    doc.add_page_break()
    populate_resume_document(doc, tailored_data, highlight_changes=False)
    doc.add_page_break()
    populate_match_matrix_docx_page(doc, cover_data)
    
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io.getvalue()

def create_master_application_zip(comb_pdf, comb_docx, review_docx, resume_pdf, cover_matrix_pdf):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr("1_Complete_Application_Set_Cover_Resume_Matrix.pdf", comb_pdf)
        zip_file.writestr("2_Complete_Application_Set_Cover_Resume_Matrix.docx", comb_docx)
        zip_file.writestr("3_Madhusudhanan_Janakarajan_Resume_Highlighted_Review.docx", review_docx)
        zip_file.writestr("4_Madhusudhanan_Janakarajan_Resume_Clean.pdf", resume_pdf)
        zip_file.writestr("5_Cover_Letter_and_Match_Matrix.pdf", cover_matrix_pdf)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

# ==============================================================================
# 6. STREAMLIT FRONTEND & ENGINE CONTROLLER
# ==============================================================================
st.title("🎯 Executive ATS Resume & Application Engine")
st.caption("Real-Time AI Tailoring • Multi-Format Suite • 5-Asset Master Bundle • ATS Scoring")

with st.sidebar:
    st.header("⚡ System Status")
    if api_key:
        st.success("🟢 Gemini AI Engine: Active (Optimized ⚡)")
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
            with st.spinner("⚡ High-Speed Synthesis: Generating complete tailored application pack..."):
                prompt = f"""
                You are an executive resume architect and career strategist for Madhusudhanan Janakarajan (23+ year FMCG, Digital Transformation & Enterprise Technology Executive).

                Analyze the provided target Job Description (JD) and special instructions to extract the company name, target role title, and generate fully customized documents.

                STRICT RULES FOR GENERATING JSON:

                1. IDENTIFY TARGET COMPANY & ROLE:
                   - "target_company": The specific company name from the JD.
                   - "target_role": The specific role title from the JD.

                2. HEADER SUBTITLE DUAL VARIABLES:
                   - Format: "[header_focus_1] | FMCG | GTM & Omnichannel Leader | [header_focus_2]"
                   - "header_focus_1": Target leadership title matching the JD (e.g. "Commercial & Digital Transformation Director", "E-Commerce & Commercial Director", "Global Distributor Management Director"). Max 36 chars.
                   - "header_focus_2": Specialized domain focus matching the JD (e.g. "Enterprise Sales Technology Leader", "Global Distributor Governance Leader", "Omnichannel RTM & Digital Execution"). Max 40 chars.

                3. EXECUTIVE SUMMARY (EXACT 135-150 WORDS / 7-8 LINES):
                   - Write an authoritative, rich executive summary of EXACTLY 135 to 150 words dynamically tailored to the role and company.
                   - Emphasize relevant commercial, distributor management, and transformation capabilities directly addressing the JD requirements.
                   - Retain core metrics ($100M+ P&L, 8,000+ retailers, 10+ Tier-1 CPG logos: P&G, Nestlé, GSK, Coca-Cola, ~40% logistics optimization, ~20% productivity uplifts).

                4. CAPABILITY ORDERING (PRIORITIZATION):
                   - Rank the 5 capability keys based on the JD's highest priorities (place the top 2 matching keys first):
                     Available keys: ["commercial", "digital", "transformation", "capability", "entrepreneurship"]
                   - "capability_order": An array containing all 5 keys in ordered priority.

                5. CONEKTR CATEGORY BULLET:
                   - Category aggregation bullet strictly tailored to the products/domain of the target company.

                6. DYNAMIC EXPERIENCE INJECTIONS (STRICT 18 TO 24 WORDS EACH):
                   - "column_2_extra_bullet": 18-24 words under Conektr (Digital FMCG) aligned with JD, else "".
                   - "column_3_extra_bullet": 18-24 words under TransCPG/Ivy (Transformation) aligned with JD, else "".

                7. ATS MATCH SCORE (INTEGER 88-97):
                   - "ats_match_score": Integer reflecting alignment with the provided JD.

                8. COVER LETTER & MATCH MATRIX (STRICTLY ONE PAGE EACH):
                   - "subject_line": "Application for [Target Role] - [Target Company]"
                   - "cover_para_1": Authoritative opening explicitly referencing the company name, role title, and candidate's 23+ year track record.
                   - "cover_para_2": Direct alignment with the target company's specific commercial and digital priorities based on JD and special instructions.
                   - "cover_bullets": 4 high-impact bullets formatted as "Bold Category: Detailed metric description" matching the visual style:
                     1) Global Distributor Management & Commercial Governance: ...
                     2) Enterprise Digital Architecture & SFA Systems: ...
                     3) Measurable P&L & Operational ROI: ...
                     4) Cross-Functional Leadership & Partner Strategy: ...
                   - "cover_para_closing": Forward-looking closing paragraph.
                   - "matrix_items": Array of EXACTLY 6 concise competency rows mapping JD pillars to candidate evidence.
                     IMPORTANT FOR MATCH MATRIX:
                     * "requirement_title": Concise single statement of the requirement only (no duplicate descriptions).
                     * "match_desc": Direct, concise narrative paragraph demonstrating evidence WITHOUT redundant bold prefixes.

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
                    "gemini-2.5-flash",
                    "gemini-3.5-flash-lite",
                    "gemini-2.0-flash"
                ]

                client = genai.Client(api_key=api_key)
                
                for model_candidate in model_candidates:
                    for attempt in range(2):
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
                            time.sleep(1.5 * (attempt + 1))
                    if tailored_data:
                        break

                if tailored_data:
                    clean_docx = create_master_resume_docx(tailored_data, highlight_changes=False)
                    review_docx = create_master_resume_docx(tailored_data, highlight_changes=True)
                    resume_pdf = create_resume_pdf(tailored_data)
                    
                    cover_matrix_pdf = create_cover_letter_match_matrix_pdf(cover_data)
                    cover_matrix_docx = create_cover_letter_match_matrix_docx(cover_data)
                    
                    comb_pdf = create_combined_application_pdf(cover_data, tailored_data)
                    comb_docx = create_combined_application_docx(cover_data, tailored_data)
                    
                    master_zip = create_master_application_zip(comb_pdf, comb_docx, review_docx, resume_pdf, cover_matrix_pdf)

                    st.session_state["tailored_data"] = tailored_data
                    st.session_state["cover_data"] = cover_data
                    st.session_state["comb_pdf"] = comb_pdf
                    st.session_state["comb_docx"] = comb_docx
                    st.session_state["review_docx"] = review_docx
                    st.session_state["resume_pdf"] = resume_pdf
                    st.session_state["cover_matrix_pdf"] = cover_matrix_pdf
                    st.session_state["master_zip"] = master_zip
                    st.session_state["has_results"] = True
                else:
                    st.error(f"Generation Error: {last_error}. Please check your API key and network permissions.")

# ==============================================================================
# 7. PERSISTENT DISPLAY & MULTI-FILE DOWNLOAD AREA
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
            label=f"📦 Download Complete Application Bundle (.ZIP) — 5 Files",
            data=st.session_state["master_zip"],
            file_name=f"Madhusudhanan_Janakarajan_{target_co.replace(' ', '_')}_Full_Application_Bundle.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True
        )

        st.markdown("---")
        st.write("📄 **Individual Application Files:**")

        col_b1, col_b2 = st.columns(2)
        with col_b1:
            st.download_button(
                label="📑 1. Combined Application Set (.pdf)",
                data=st.session_state["comb_pdf"],
                file_name=f"1_Complete_Application_Set_Cover_Resume_Matrix.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        with col_b2:
            st.download_button(
                label="📝 2. Combined Application Set (.docx)",
                data=st.session_state["comb_docx"],
                file_name=f"2_Complete_Application_Set_Cover_Resume_Matrix.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

        col_b3, col_b4 = st.columns(2)
        with col_b3:
            st.download_button(
                label="🟡 3. Highlighted Review Resume (.docx)",
                data=st.session_state["review_docx"],
                file_name=f"3_Madhusudhanan_Janakarajan_Resume_Review.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        with col_b4:
            st.download_button(
                label="📄 4. Clean ATS Resume (.pdf)",
                data=st.session_state["resume_pdf"],
                file_name=f"4_Madhusudhanan_Janakarajan_Resume_Clean.pdf",
                mime="application/pdf",
                use_container_width=True
            )

        st.download_button(
            label="📊 5. Cover Letter & Match Matrix (.pdf)",
            data=st.session_state["cover_matrix_pdf"],
            file_name=f"5_Cover_Letter_and_Match_Matrix.pdf",
            mime="application/pdf",
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
