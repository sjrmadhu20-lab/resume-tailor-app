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
from docx.oxml import parse_xml
from google import genai
from google.genai import types

st.set_page_config(page_title="Executive ATS Application Engine", page_icon="🎯", layout="wide")

# ==============================================================================
# 1. API CONFIGURATION
# ==============================================================================
api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))[cite: 1]

# ==============================================================================
# 2. MASTER KNOWLEDGE ARCHIVE (LOCKED RESUME DATA)
# ==============================================================================
MASTER_CAPABILITIES = {
    "commercial": "Commercial & GTM Leadership ($100M+ P&L): Owned $100M+ annual FMCG revenue across GCC & India, directing 250+ distributors and 600+ field sales teams across GT, MT, Wholesale, B2B, and Institutional channels. Spearheaded RTM redesign, distributor governance, trade margin economics, pricing/promotions, and Order-to-Cash optimization.",
    "digital": "Digital B2B2C Commerce & Omnichannel RTM: Founded and scaled Conektr (UAE's first digital FMCG distributor) to 8,000+ B2B retailers, managing 100+ brands and 2,000+ SKUs across Foods, Beverages, and Non-Food categories. Expanded into direct B2C commerce by launching the consumer app and proprietary BOSS loyalty engine (Buying, Operating, Selling & Saving), turning network grocers into fulfillment micro-hubs/dark stores. Built omnichannel ordering (App, Web, Conversational AI) with fintech-enabled payment rails.",
    "transformation": "Enterprise Transformation & Commercial Optimization: Directed multi-country RTM modernizations, DMS/ERP integrations and SFA deployments (Over 5000+ Users) for global CPG leaders (P&G, Nestlé, Haleon/GSK, Coca-Cola). Deployed AI route/beat optimization, AI-driven demand forecasting, and automated ordering—delivering a ~40% drop in logistics/admin costs, >30% reduction in outlet coverage costs, ~30% frontline sales productivity uplift, ~150% expansion in numeric distribution growth.",
    "capability": "Sales Capability & Training Leadership: Established and led regional sales training operations managing a team of certified trainers to design and deliver end-to-end sales induction and leadership curricula up to the Regional Sales Manager (RSM) level. Top-performer in consultative selling frameworks (including SPIN Selling), driving frontline execution rigor, distributor capability building, and institutionalized sales performance standards. Managed Train the Trainer, Soft skills and automation training.",
    "entrepreneurship": "Entrepreneurial Venture Scaling & Governance: Raised $15M in funding from DIFC VC, veteran FMCG executives (ex-Mondelēz President, BAT CFO) validating commercial credibility, and executed a successful strategic M&A exit to Al Maya Group ($1B+ conglomerate). Awarded UAE Golden Visa and USA O-1A (Extraordinary Ability); featured in Bloomberg, Gulf News, and Magnitt."
}[cite: 1]

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
}[cite: 1]

def add_hyperlink(paragraph, url, text, color_rgb="004B87", underline=True, font_size_pt=10, is_highlighted=False):
    part = paragraph.part[cite: 1]
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)[cite: 1]

    hyperlink = parse_xml(f'<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" r:id="{r_id}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>')[cite: 1]
    new_run = parse_xml(f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')[cite: 1]
    rPr = parse_xml(f'<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')[cite: 1]
    
    rPr.append(parse_xml(f'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Calibri" w:hAnsi="Calibri"/>'))[cite: 1]
    val_sz = int(font_size_pt * 2)[cite: 1]
    rPr.append(parse_xml(f'<w:sz xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="{val_sz}"/>'))[cite: 1]
    rPr.append(parse_xml(f'<w:color xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="{color_rgb}"/>'))[cite: 1]
    if underline:
        rPr.append(parse_xml(f'<w:u xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="single"/>'))[cite: 1]
    if is_highlighted:
        rPr.append(parse_xml(r'<w:highlight xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="yellow"/>'))[cite: 1]
        
    new_run.append(rPr)[cite: 1]
    new_run.append(parse_xml(f'<w:t xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{text}</w:t>'))[cite: 1]
    hyperlink.append(new_run)[cite: 1]
    paragraph._p.append(hyperlink)[cite: 1]

# ==============================================================================
# 3. WORD RESUME ENGINE (LOCKED V1 BUILD)
# ==============================================================================
def populate_resume_document(doc, tailored_data, highlight_changes=False):
    style = doc.styles['Normal'][cite: 1]
    style.font.name = 'Calibri'[cite: 1]
    style.font.size = Pt(10)[cite: 1]
    style.font.color.rgb = RGBColor(0x00, 0x00, 0x00)[cite: 1]

    def apply_xml_spacing(p, before_pt=0, after_pt=8, line_twips=278):
        pPr = p._p.get_or_add_pPr()[cite: 1]
        spPr = parse_xml(f'<w:spacing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:before="{int(before_pt*20)}" w:after="{int(after_pt*20)}" w:line="{line_twips}" w:lineRule="auto"/>')
        pPr.append(spPr)[cite: 1]

    def add_heading(title, space_before=0, space_after=8, line_border_above=False, is_multiple=False, is_underline=False):
        p = doc.add_paragraph()[cite: 1]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT[cite: 1]
        if is_multiple:
            apply_xml_spacing(p, before_pt=space_before, after_pt=space_after, line_twips=278)
        else:
            apply_xml_spacing(p, before_pt=space_before, after_pt=space_after, line_twips=240)
        
        if line_border_above:
            pPr = p._p.get_or_add_pPr()[cite: 1]
            pBdr = parse_xml(r'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                             r'<w:top w:val="single" w:sz="6" w:space="4" w:color="000000"/>'
                             r'</w:pBdr>')[cite: 1]
            pPr.append(pBdr)[cite: 1]
            
        r = p.add_run(title.upper() if title != "LANGUAGES & INTERESTS :" else title)
        r.bold = True[cite: 1]
        r.underline = is_underline
        r.font.name = 'Calibri'[cite: 1]
        r.font.size = Pt(10)[cite: 1]
        r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)[cite: 1]

    # ---------------- PAGE 1 (LOCKED V1) ----------------
    p_name = doc.add_paragraph()[cite: 1]
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER[cite: 1]
    apply_xml_spacing(p_name, before_pt=0, after_pt=0, line_twips=278)
    r_name = p_name.add_run(MASTER_STATIC['name'])[cite: 1]
    r_name.bold = True[cite: 1]
    r_name.font.name = 'Calibri'[cite: 1]
    r_name.font.size = Pt(12)[cite: 1]

    f1 = tailored_data.get("header_focus_1", "Sales & Distribution Transformation Director")[cite: 1]
    f2 = tailored_data.get("header_focus_2", "Beauty & Personal Care Experience")[cite: 1]
    
    p_sub = doc.add_paragraph()[cite: 1]
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER[cite: 1]
    apply_xml_spacing(p_sub, before_pt=0, after_pt=8, line_twips=278)
    
    r_f1 = p_sub.add_run(f1)[cite: 1]
    r_f1.bold = True[cite: 1]
    r_f1.font.name = 'Calibri'[cite: 1]
    r_f1.font.size = Pt(9)[cite: 1]
    if highlight_changes:
        r_f1.font.highlight_color = docx.enum.text.WD_COLOR_INDEX.YELLOW[cite: 1]
        
    r_mid = p_sub.add_run(" | FMCG | GTM & Omnichannel Leader | ")[cite: 1]
    r_mid.bold = True[cite: 1]
    r_mid.font.name = 'Calibri'[cite: 1]
    r_mid.font.size = Pt(9)[cite: 1]
    
    r_f2 = p_sub.add_run(f2)[cite: 1]
    r_f2.bold = True[cite: 1]
    r_f2.font.name = 'Calibri'[cite: 1]
    r_f2.font.size = Pt(9)[cite: 1]
    if highlight_changes:
        r_f2.font.highlight_color = docx.enum.text.WD_COLOR_INDEX.YELLOW[cite: 1]

    c = MASTER_STATIC['contact'][cite: 1]
    p_contact = doc.add_paragraph()[cite: 1]
    p_contact.alignment = WD_ALIGN_PARAGRAPH.CENTER[cite: 1]
    apply_xml_spacing(p_contact, before_pt=0, after_pt=8, line_twips=278)
    
    r_c1 = p_contact.add_run(f"{c['location']} | {c['phone']} | ")[cite: 1]
    r_c1.font.name = 'Calibri'[cite: 1]
    r_c1.font.size = Pt(10)[cite: 1]
    add_hyperlink(p_contact, c['email_url'], c['email'], color_rgb="004B87", underline=True, font_size_pt=10)[cite: 1]
    
    r_br1 = p_contact.add_run("\n")[cite: 1]
    r_br1.font.name = 'Calibri'[cite: 1]
    r_br1.font.size = Pt(10)[cite: 1]
    
    add_hyperlink(p_contact, c['linkedin'], c['linkedin'], color_rgb="004B87", underline=True, font_size_pt=10)[cite: 1]
    r_c2_mid = p_contact.add_run(" | Portfolio: ")[cite: 1]
    r_c2_mid.font.name = 'Calibri'[cite: 1]
    r_c2_mid.font.size = Pt(10)[cite: 1]
    add_hyperlink(p_contact, c['portfolio'], c['portfolio'], color_rgb="004B87", underline=True, font_size_pt=10)[cite: 1]
    
    r_br2 = p_contact.add_run("\n")[cite: 1]
    r_br2.font.name = 'Calibri'[cite: 1]
    r_br2.font.size = Pt(10)[cite: 1]
    
    r_c3_lbl = p_contact.add_run("Visa Status: ")[cite: 1]
    r_c3_lbl.bold = True[cite: 1]
    r_c3_lbl.font.name = 'Calibri'[cite: 1]
    r_c3_lbl.font.size = Pt(10)[cite: 1]
    r_c3_val = p_contact.add_run(c['visas'])[cite: 1]
    r_c3_val.font.name = 'Calibri'[cite: 1]
    r_c3_val.font.size = Pt(10)[cite: 1]

    add_heading("EXECUTIVE SUMMARY", space_before=0, space_after=8, line_border_above=False, is_multiple=True)
    sp = doc.add_paragraph()[cite: 1]
    sp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY[cite: 1]
    apply_xml_spacing(sp, before_pt=0, after_pt=6, line_twips=278)
    r_sum = sp.add_run(tailored_data.get("executive_summary", ""))[cite: 1]
    r_sum.font.name = 'Calibri'[cite: 1]
    r_sum.font.size = Pt(10)[cite: 1]
    if highlight_changes:
        r_sum.font.highlight_color = docx.enum.text.WD_COLOR_INDEX.YELLOW[cite: 1]

    add_heading("EXECUTIVE CAPABILITIES & IMPACT HIGHLIGHTS", space_before=2, space_after=8, line_border_above=True, is_multiple=False)
    for cap in tailored_data.get("capabilities", [])
        cp = doc.add_paragraph()[cite: 1]
        cp.paragraph_format.left_indent = Inches(0.20)
        cp.paragraph_format.first_line_indent = Inches(-0.25)
        cp.alignment = WD_ALIGN_PARAGRAPH.LEFT[cite: 1]
        apply_xml_spacing(cp, before_pt=0, after_pt=4.5, line_twips=240)
        
        r_bullet = cp.add_run("•\t")
        r_bullet.font.name = 'Calibri'[cite: 1]
        r_bullet.font.size = Pt(10)[cite: 1]
        
        parts = cap.split(":", 1)[cite: 1]
        if len(parts) == 2:[cite: 1]
            r_bold = cp.add_run(parts[0] + ":")[cite: 1]
            r_bold.bold = True[cite: 1]
            r_bold.font.name = 'Calibri'[cite: 1]
            r_bold.font.size = Pt(10)[cite: 1]
            r_body = cp.add_run(parts[1])[cite: 1]
            r_body.font.name = 'Calibri'[cite: 1]
            r_body.font.size = Pt(10)[cite: 1]
        else:
            r_body = cp.add_run(cap)[cite: 1]
            r_body.font.name = 'Calibri'[cite: 1]
            r_body.font.size = Pt(10)[cite: 1]

    add_heading("HONORS & RECOGNITION", space_before=2, space_after=8, line_border_above=True, is_multiple=False)
    for idx, h in enumerate(MASTER_STATIC['honors']):[cite: 1]
        p = doc.add_paragraph()[cite: 1]
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY[cite: 1]
        p.paragraph_format.left_indent = Inches(0.45)
        p.paragraph_format.first_line_indent = Inches(-0.20)
        is_last = (idx == len(MASTER_STATIC['honors']) - 1)
        apply_xml_spacing(p, before_pt=0, after_pt=6 if is_last else 0, line_twips=240)
        
        r_bullet = p.add_run("•\t")
        r_bullet.font.name = 'Calibri'[cite: 1]
        r_bullet.font.size = Pt(10)[cite: 1]
        
        if "https://" in h:
            parts = h.split(" - ")
            r_prefix = p.add_run(parts[0] + " - ")
            r_prefix.font.name = 'Calibri'
            r_prefix.font.size = Pt(10)
            add_hyperlink(p, parts[1].strip(), parts[1].strip(), color_rgb="004B87", underline=True, font_size_pt=10)
        else:
            r_t = p.add_run(h)[cite: 1]
            r_t.font.name = 'Calibri'[cite: 1]
            r_t.font.size = Pt(10)[cite: 1]

    add_heading("EDUCATION", space_before=0, space_after=8, line_border_above=False, is_multiple=False)
    for idx, edu in enumerate(MASTER_STATIC['education']):[cite: 1]
        p = doc.add_paragraph()[cite: 1]
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY[cite: 1]
        p.paragraph_format.left_indent = Inches(0.45)
        p.paragraph_format.first_line_indent = Inches(-0.20)
        is_last = (idx == len(MASTER_STATIC['education']) - 1)
        apply_xml_spacing(p, before_pt=0, after_pt=6 if is_last else 0, line_twips=240)
        
        r_bullet = p.add_run("•\t")
        r_bullet.font.name = 'Calibri'[cite: 1]
        r_bullet.font.size = Pt(10)[cite: 1]
        
        r_bp = p.add_run(edu['degree'] + " – ")[cite: 1]
        r_bp.bold = True[cite: 1]
        r_bp.font.name = 'Calibri'[cite: 1]
        r_bp.font.size = Pt(10)[cite: 1]
        r_t = p.add_run(edu['details'])[cite: 1]
        r_t.font.name = 'Calibri'[cite: 1]
        r_t.font.size = Pt(10)[cite: 1]

    add_heading("LANGUAGES & INTERESTS :", space_before=0, space_after=8, line_border_above=False, is_multiple=False)
    p_lang1 = doc.add_paragraph()[cite: 1]
    p_lang1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY[cite: 1]
    p_lang1.paragraph_format.left_indent = Inches(0.45)
    p_lang1.paragraph_format.first_line_indent = Inches(-0.20)
    apply_xml_spacing(p_lang1, before_pt=0, after_pt=0, line_twips=240)
    r_bullet_l1 = p_lang1.add_run("•\t")
    r_bullet_l1.font.name = 'Calibri'[cite: 1]
    r_bullet_l1.font.size = Pt(10)[cite: 1]
    r_l1 = p_lang1.add_run(MASTER_STATIC['languages'])[cite: 1]
    r_l1.font.name = 'Calibri'[cite: 1]
    r_l1.font.size = Pt(10)[cite: 1]

    p_lang2 = doc.add_paragraph()[cite: 1]
    p_lang2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY[cite: 1]
    p_lang2.paragraph_format.left_indent = Inches(0.45)
    p_lang2.paragraph_format.first_line_indent = Inches(-0.20)
    apply_xml_spacing(p_lang2, before_pt=0, after_pt=0, line_twips=240)
    r_bullet_l2 = p_lang2.add_run("•\t")
    r_bullet_l2.font.name = 'Calibri'[cite: 1]
    r_bullet_l2.font.size = Pt(10)[cite: 1]
    r_l2 = p_lang2.add_run(MASTER_STATIC['interests'])[cite: 1]
    r_l2.font.name = 'Calibri'[cite: 1]
    r_l2.font.size = Pt(10)[cite: 1]

    # ---------------- PAGE 2 (LOCKED V1) ----------------
    doc.add_page_break()[cite: 1]

    add_heading("PROFESSIONAL EXPERIENCE", space_before=0, space_after=8, line_border_above=False, is_multiple=False)
    
    table = doc.add_table(rows=2, cols=3)[cite: 1]
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False[cite: 1]
    
    tblPr = table._tbl.tblPr
    tblpPr = parse_xml(r'<w:tblpPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:tblpX="-274" w:tblpY="0"/>')
    tblPr.append(tblpPr)
    
    col_widths = [Inches(2.63), Inches(2.63), Inches(2.63)]
    for row in table.rows:[cite: 1]
        for i, cell in enumerate(row.cells):[cite: 1]
            cell.width = col_widths[i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    hdr_titles = ["Traditional FMCG Operator", "Digital FMCG Distribution", "Distribution Transformation"][cite: 1]
    for i, title in enumerate(hdr_titles):[cite: 1]
        cell = table.rows[0].cells[i][cite: 1]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0][cite: 1]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_xml_spacing(p, before_pt=3, after_pt=3, line_twips=240)
        r = p.add_run(title)[cite: 1]
        r.bold = True[cite: 1]
        r.font.name = 'Calibri'[cite: 1]
        r.font.size = Pt(12)
        tcPr = cell._tc.get_or_add_tcPr()[cite: 1]
        tcPr.append(parse_xml(r'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="DCE6F1"/>'))

    def populate_cell_content(cell, item_list):
        cell.text = ""[cite: 1]
        for idx, item in enumerate(item_list):[cite: 1]
            p = cell.add_paragraph()[cite: 1]
            apply_xml_spacing(p, before_pt=item.get("space_before", 0), after_pt=item.get("space_after", 2), line_twips=220)
            
            if item.get("is_bullet", False):[cite: 1]
                p.paragraph_format.left_indent = Inches(0.18)
                p.paragraph_format.first_line_indent = Inches(-0.15)
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r_b = p.add_run("•\t")
                r_b.font.name = 'Calibri'
                r_b.font.size = Pt(10)
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                
            r = p.add_run(item["text"])[cite: 1]
            r.bold = item.get("bold", False)[cite: 1]
            r.italic = item.get("italic", False)[cite: 1]
            r.font.name = 'Calibri'[cite: 1]
            r.font.size = Pt(item.get("size", 10))
            if highlight_changes and item.get("highlight", False):[cite: 1]
                r.font.highlight_color = docx.enum.text.WD_COLOR_INDEX.YELLOW[cite: 1]

    c0_items = [
        {"text": "Britannia Industries Ltd | 2007 – 2011", "bold": True, "size": 10, "space_before": 2},
        {"text": "Regional Sales Head – GCC", "bold": True, "size": 10},
        {"text": "Regional Sales & Capability Head- India", "bold": True, "size": 10, "space_after": 4},
        {"text": "Owned $100M+ P&L across GCC (Saudi Arabia, UAE, Kuwait, Oman, Bahrain, Qatar) & South India.", "is_bullet": True, "size": 10},
        {"text": "Directed 250+ distributor networks & 600+ frontline sales staff across GT, MT, wholesale, and institutional trade.", "is_bullet": True, "size": 10},
        {"text": "Spearheaded Britannia's 1st national SFA rollout (1,000+ users), transforming legacy trade into performance-managed selling.", "is_bullet": True, "size": 10},
        {"text": "Delivered ~30% numeric distribution growth, increased LPC to ~120%, and cut sales admin costs by ~30%.", "is_bullet": True, "size": 10},
        {"text": "Turnaround RSM GCC: achieved record monthly sales for 3 consecutive months (Best Employee Award from Group MD).", "is_bullet": True, "size": 10, "space_after": 3},
        {"text": "Airtel | Reliance | Tyco | 2001 – 2007", "bold": True, "size": 10, "space_before": 5},
        {"text": "Commercial & Training Roles –", "bold": True, "size": 10, "space_after": 4},
        {"text": "Built foundations in frontline trade execution, journey planning, and merchandiser enablement in telecom & enterprise security.", "is_bullet": True, "size": 10},
        {"text": "Deployed capability training (SPIN selling) & integrated Oracle e-CRM & LMS infrastructure at scale.", "is_bullet": True, "size": 10}
    ]

    c0_extra = tailored_data.get("column_1_extra_bullet", "")
    if c0_extra and c0_extra.strip():
        c0_items.insert(7, {"text": c0_extra.strip(), "is_bullet": True, "size": 10, "highlight": True})

    conektr_cat = tailored_data.get("conektr_category_bullet", "Deep FMCG Category Aggregation: Scaled multi-category catalogs across ambient, packaged food, and consumer goods portfolios.")[cite: 1]
    c1_items = [
        {"text": "Digital FMCG Principal / Distributor", "bold": True, "size": 10, "space_before": 2},
        {"text": "Conektr Tech Global Ltd | UAE & India", "bold": True, "size": 10, "space_after": 4},
        {"text": "Chief Executive Officer & Founder", "bold": True, "size": 10},
        {"text": "May 2016 – Aug 2024", "size": 10, "space_after": 4},
        {"text": "Founded UAE’s 1st Digital FMCG Principal-Distributor serving 8,000+ retailers (2,000+ MAU) & 100+ brands.", "is_bullet": True, "size": 10},
        {"text": conektr_cat, "is_bullet": True, "size": 10, "highlight": True},
        {"text": "Owned full P&L, trade terms, warehousing, last-mile delivery, trade credit, and collections.", "is_bullet": True, "size": 10},
        {"text": "Built app/web/WhatsApp self-ordering engine scaling annual GMV from zero to ~AED 50M (~$13.6M) at ~18% gross margin.", "is_bullet": True, "size": 10},
        {"text": "Cut coverage cost by >50% and improved field execution productivity by ~150% vs traditional trade.", "is_bullet": True, "size": 10},
        {"text": "Deployed Dynamics 365 + Power BI and AI route optimization, cutting logistics costs by ~40%.", "is_bullet": True, "size": 10},
        {"text": "Raised ~$15M from C-suite FMCG leaders; executed M&A exit to Al Maya Group ($1B+ retail conglomerate).", "is_bullet": True, "size": 10}
    ]

    c2_items = [
        {"text": "Post Exit –", "size": 10, "space_before": 2, "space_after": 3},
        {"text": "Transformation Advisor (Director)", "bold": True, "size": 10},
        {"text": "TransCPG Inc. &", "bold": True, "size": 10},
        {"text": "FieldAssist | 2025 – Present", "bold": True, "size": 10, "space_after": 4},
        {"text": "Board Member guiding global operations scaling & platform build across FMCG principals & distributors.", "is_bullet": True, "size": 10},
        {"text": "Advising CPG leaders on modernizing RTM & SAP/Oracle SFA/DMS integrations, driving ~150% coverage growth.", "is_bullet": True, "size": 10},
        {"text": "Built Bid2Bill AI/Voice-bot & WhatsApp B2B2C bidding platform, cutting CAC by ~40% with 4x engagement.", "is_bullet": True, "size": 10, "space_after": 3},
        {"text": "Business Head – MEA", "bold": True, "size": 10, "space_before": 5},
        {"text": "Ivy Mobility Pte Ltd | 2011 – 2016", "bold": True, "size": 10, "space_after": 4},
        {"text": "Built MEA setup from scratch into 2nd largest global setup ($10M+ pipeline across 10+ countries).", "is_bullet": True, "size": 10},
        {"text": "Won 22 enterprise logos: Haleon/GSK, P&G, Nestlé, Coca-Cola, Mars, Red Bull, BAT, and AKI Group.", "is_bullet": True, "size": 10},
        {"text": "Personally led on-ground field deployment of mobile SFA for P&G distributor networks in Kenya.", "is_bullet": True, "size": 10},
        {"text": "Deployed Cloud SaaS SFA/DMS to 3,000+ sales users, driving post-implementation adoption and trade ROI.", "is_bullet": True, "size": 10}
    ]

    c1_extra = tailored_data.get("column_2_extra_bullet", "")[cite: 1]
    if c1_extra and c1_extra.strip():[cite: 1]
        c1_items.insert(7, {"text": c1_extra.strip(), "is_bullet": True, "size": 10, "highlight": True})

    c2_extra = tailored_data.get("column_3_extra_bullet", "")[cite: 1]
    if c2_extra and c2_extra.strip():[cite: 1]
        c2_items.insert(4, {"text": c2_extra.strip(), "is_bullet": True, "size": 10, "highlight": True})

    populate_cell_content(table.rows[1].cells[0], c0_items)
    populate_cell_content(table.rows[1].cells[1], c1_items)
    populate_cell_content(table.rows[1].cells[2], c2_items)

    tblBorders = parse_xml(
        r'<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        r'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        r'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'</w:tblBorders>'
    )
    table._tbl.tblPr.append(tblBorders)

    add_heading("TECHNOLOGY STACK & DIGITAL ARCHITECTURE:", space_before=8, space_after=8, line_border_above=False, is_multiple=False)
    for category, stack in MASTER_STATIC['tech_stack'].items():[cite: 1]
        tp = doc.add_paragraph()[cite: 1]
        tp.paragraph_format.left_indent = Inches(0.20)
        tp.paragraph_format.first_line_indent = Inches(-0.25)
        tp.alignment = WD_ALIGN_PARAGRAPH.LEFT[cite: 1]
        apply_xml_spacing(tp, before_pt=0, after_pt=4, line_twips=240)
        
        r_b = tp.add_run("•\t")
        r_b.font.name = 'Calibri'[cite: 1]
        r_b.font.size = Pt(10)[cite: 1]
        
        r_cat = tp.add_run(f"{category}: ")[cite: 1]
        r_cat.bold = True[cite: 1]
        r_cat.font.name = 'Calibri'[cite: 1]
        r_cat.font.size = Pt(10)[cite: 1]
        
        r_st = tp.add_run(stack)[cite: 1]
        r_st.font.name = 'Calibri'[cite: 1]
        r_st.font.size = Pt(10)[cite: 1]

    add_heading("WHY HIRE ME", space_before=8, space_after=4, line_border_above=False, is_multiple=False, is_underline=True)
    
    p_why = doc.add_paragraph()[cite: 1]
    p_why.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY[cite: 1]
    apply_xml_spacing(p_why, before_pt=0, after_pt=6, line_twips=240)
    
    for text_segment, is_plus in MASTER_STATIC['why_hire_me_parts']:[cite: 1]
        r_part = p_why.add_run(text_segment)[cite: 1]
        r_part.font.name = 'Calibri'[cite: 1]
        r_part.font.size = Pt(10)[cite: 1]
        if is_plus:[cite: 1]
            r_part.bold = True[cite: 1]
            r_part.font.color.rgb = RGBColor(0x00, 0xB0, 0xF0)[cite: 1]

def create_master_resume_docx(tailored_data, highlight_changes=False):
    doc = Document()[cite: 1]
    for section in doc.sections:[cite: 1]
        section.top_margin = Inches(0.40)[cite: 1]
        section.bottom_margin = Inches(0.40)[cite: 1]
        section.left_margin = Inches(0.50)[cite: 1]
        section.right_margin = Inches(0.50)[cite: 1]
    populate_resume_document(doc, tailored_data, highlight_changes)[cite: 1]
    doc_io = io.BytesIO()[cite: 1]
    doc.save(doc_io)[cite: 1]
    doc_io.seek(0)[cite: 1]
    return doc_io.getvalue()[cite: 1]

# ==============================================================================
# 4. WORD COVER, MATRIX & COMBINED PACK BUILDER (.DOCX)
# ==============================================================================
def populate_cover_letter_docx_page(doc, cover_data):
    p_title = doc.add_paragraph()[cite: 1]
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER[cite: 1]
    p_title.paragraph_format.space_after = Pt(10)[cite: 1]
    r_t = p_title.add_run("COVER LETTER")[cite: 1]
    r_t.bold = True[cite: 1]
    r_t.font.name = 'Calibri'[cite: 1]
    r_t.font.size = Pt(14)[cite: 1]

    p_sub = doc.add_paragraph()[cite: 1]
    p_sub.paragraph_format.space_before = Pt(6)[cite: 1]
    p_sub.paragraph_format.space_after = Pt(8)[cite: 1]
    r_sb = p_sub.add_run(f"Subject: {cover_data.get('subject_line', '')}")[cite: 1]
    r_sb.bold = True[cite: 1]
    r_sb.font.name = 'Calibri'[cite: 1]
    r_sb.font.size = Pt(11)[cite: 1]

    p_d = doc.add_paragraph("Dear Hiring Team,")[cite: 1]
    p_d.paragraph_format.space_before = Pt(6)[cite: 1]
    p_d.paragraph_format.space_after = Pt(8)[cite: 1]

    p_p1 = doc.add_paragraph(cover_data.get("cover_para_1", ""))[cite: 1]
    p_p1.paragraph_format.space_after = Pt(8)[cite: 1]
    p_p1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY[cite: 1]

    p_p2 = doc.add_paragraph(cover_data.get("cover_para_2", ""))[cite: 1]
    p_p2.paragraph_format.space_after = Pt(8)[cite: 1]
    p_p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY[cite: 1]

    p_kh = doc.add_paragraph()[cite: 1]
    p_kh.paragraph_format.space_after = Pt(6)[cite: 1]
    r_kh = p_kh.add_run("Key highlights of what I bring to this mandate include:")[cite: 1]
    r_kh.bold = True[cite: 1]
    r_kh.font.name = 'Calibri'[cite: 1]
    r_kh.font.size = Pt(11)[cite: 1]

    for b in cover_data.get("cover_bullets", []):[cite: 1]
        bp = doc.add_paragraph(style='List Bullet')[cite: 1]
        bp.paragraph_format.left_indent = Inches(0.5)[cite: 1]
        bp.paragraph_format.right_indent = Inches(0.2)[cite: 1]
        bp.paragraph_format.space_after = Pt(6)[cite: 1]
        bp.paragraph_format.line_spacing = 1.15[cite: 1]
        bp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY[cite: 1]
        
        parts = b.split(":", 1)[cite: 1]
        if len(parts) == 2:[cite: 1]
            r_head = bp.add_run(parts[0] + ": ")[cite: 1]
            r_head.bold = True[cite: 1]
            r_head.font.name = 'Calibri'[cite: 1]
            r_head.font.size = Pt(10.5)[cite: 1]
            r_tail = bp.add_run(parts[1].strip())[cite: 1]
            r_tail.font.name = 'Calibri'[cite: 1]
            r_tail.font.size = Pt(10.5)[cite: 1]
        else:
            r_b = bp.add_run(b)[cite: 1]
            r_b.font.name = 'Calibri'[cite: 1]
            r_b.font.size = Pt(10.5)[cite: 1]

    p_cl = doc.add_paragraph(cover_data.get("cover_para_closing", ""))[cite: 1]
    p_cl.paragraph_format.space_before = Pt(6)[cite: 1]
    p_cl.paragraph_format.space_after = Pt(8)[cite: 1]
    p_cl.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY[cite: 1]
    
    p_sign = doc.add_paragraph()[cite: 1]
    p_sign.paragraph_format.space_before = Pt(8)[cite: 1]
    r_s0 = p_sign.add_run("Sincerely,\n")[cite: 1]
    r_s0.font.name = 'Calibri'[cite: 1]
    r_s1 = p_sign.add_run("Madhusudhanan Janakarajan (Madhu)\n")[cite: 1]
    r_s1.bold = True[cite: 1]
    r_s1.font.name = 'Calibri'[cite: 1]
    r_s2 = p_sign.add_run("+971 50 654 7858 | sjrmadhu20@gmail.com")[cite: 1]
    r_s2.font.name = 'Calibri'[cite: 1]

def populate_match_matrix_docx_page(doc, cover_data):
    p_mtitle = doc.add_paragraph()[cite: 1]
    p_mtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER[cite: 1]
    p_mtitle.paragraph_format.space_after = Pt(8)[cite: 1]
    r_mt = p_mtitle.add_run("MATCH MATRIX")[cite: 1]
    r_mt.bold = True[cite: 1]
    r_mt.font.name = 'Calibri'[cite: 1]
    r_mt.font.size = Pt(14)[cite: 1]

    matrix_items = cover_data.get("matrix_items", [])[cite: 1]
    table = doc.add_table(rows=len(matrix_items) + 1, cols=2)[cite: 1]
    table.alignment = WD_TABLE_ALIGNMENT.CENTER[cite: 1]
    table.autofit = False[cite: 1]
    table.rows[0].cells[0].width = Inches(2.5)[cite: 1]
    table.rows[0].cells[1].width = Inches(5.0)[cite: 1]

    cell_0 = table.rows[0].cells[0][cite: 1]
    cell_1 = table.rows[0].cells[1][cite: 1]
    
    p_h0 = cell_0.paragraphs[0][cite: 1]
    p_h0.alignment = WD_ALIGN_PARAGRAPH.CENTER[cite: 1]
    r_h0 = p_h0.add_run("Target Job Requirement / Focus Domain")[cite: 1]
    r_h0.bold = True[cite: 1]
    r_h0.font.name = 'Calibri'[cite: 1]
    r_h0.font.size = Pt(10.5)[cite: 1]

    p_h1 = cell_1.paragraphs[0][cite: 1]
    p_h1.alignment = WD_ALIGN_PARAGRAPH.CENTER[cite: 1]
    r_h1 = p_h1.add_run("How I Match (Evidence & Track Record)")[cite: 1]
    r_h1.bold = True[cite: 1]
    r_h1.font.name = 'Calibri'[cite: 1]
    r_h1.font.size = Pt(10.5)[cite: 1]

    for idx, item in enumerate(matrix_items):[cite: 1]
        row_cells = table.rows[idx + 1].cells[cite: 1]
        row_cells[0].width = Inches(2.5)[cite: 1]
        row_cells[1].width = Inches(5.0)[cite: 1]
        
        p0 = row_cells[0].paragraphs[0][cite: 1]
        p0.alignment = WD_ALIGN_PARAGRAPH.LEFT[cite: 1]
        p0.paragraph_format.space_before = Pt(3)[cite: 1]
        p0.paragraph_format.space_after = Pt(3)[cite: 1]
        r_rt = p0.add_run(item.get('requirement_title', ''))[cite: 1]
        r_rt.bold = True[cite: 1]
        r_rt.font.name = 'Calibri'[cite: 1]
        r_rt.font.size = Pt(9.5)[cite: 1]
        
        p1 = row_cells[1].paragraphs[0][cite: 1]
        p1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY[cite: 1]
        p1.paragraph_format.space_before = Pt(3)[cite: 1]
        p1.paragraph_format.space_after = Pt(3)[cite: 1]
        r_mt = p1.add_run(item.get('match_desc', ''))[cite: 1]
        r_mt.font.name = 'Calibri'[cite: 1]
        r_mt.font.size = Pt(9.5)[cite: 1]

    tblBorders = parse_xml(
        r'<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        r'<w:top w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        r'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>'
        r'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>'
        r'</w:tblBorders>'
    )[cite: 1]
    table._tbl.tblPr.append(tblBorders)[cite: 1]

def create_combined_application_docx(cover_data, tailored_data):
    doc = Document()[cite: 1]
    for section in doc.sections:[cite: 1]
        section.top_margin = Inches(0.40)[cite: 1]
        section.bottom_margin = Inches(0.40)[cite: 1]
        section.left_margin = Inches(0.5)[cite: 1]
        section.right_margin = Inches(0.5)[cite: 1]
    
    populate_cover_letter_docx_page(doc, cover_data)[cite: 1]
    doc.add_page_break()[cite: 1]
    populate_resume_document(doc, tailored_data, highlight_changes=False)[cite: 1]
    doc.add_page_break()[cite: 1]
    populate_match_matrix_docx_page(doc, cover_data)[cite: 1]
    
    doc_io = io.BytesIO()[cite: 1]
    doc.save(doc_io)[cite: 1]
    doc_io.seek(0)[cite: 1]
    return doc_io.getvalue()[cite: 1]

def create_master_application_zip(comb_docx, review_docx, clean_docx):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr("1_Complete_Application_Set_Cover_Resume_Matrix.docx", comb_docx)
        zip_file.writestr("2_Madhusudhanan_Janakarajan_Resume_Highlighted_Review.docx", review_docx)
        zip_file.writestr("3_Madhusudhanan_Janakarajan_Resume_Clean.docx", clean_docx)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

# ==============================================================================
# 5. STREAMLIT FRONTEND & ENGINE CONTROLLER
# ==============================================================================
st.title("🎯 Executive ATS Resume & Application Engine")[cite: 1]
st.caption("Real-Time AI Tailoring • Word (.docx) Suite • 3-Asset Master Bundle • ATS Scoring")

with st.sidebar:
    st.header("⚡ System Status")[cite: 1]
    if api_key:[cite: 1]
        st.success("🟢 Gemini AI Engine: Active (Optimized ⚡)")[cite: 1]
    else:
        st.error("🔴 AI Engine Key Missing (Set GEMINI_API_KEY in Secrets)")[cite: 1]
    st.markdown("---")[cite: 1]
    if st.button("🔄 Reset / Clear Cached Data"):[cite: 1]
        for key in list(st.session_state.keys()):[cite: 1]
            del st.session_state[key][cite: 1]
        st.rerun()[cite: 1]

col1, col2 = st.columns([1.1, 0.9])[cite: 1]

with col1:
    st.subheader("1. Job Inputs & Specifics")[cite: 1]
    job_desc = st.text_area("Target Job Description (JD):", height=240, placeholder="Paste target Job Description here...")[cite: 1]

    st.markdown("##### Special Instructions & Context (Optional)")[cite: 1]
    
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
    )[cite: 1]

    special_instructions = st.text_area(
        "Voice or Typed Notes:",
        height=90,
        placeholder="Click mic above or type instructions here...",
        label_visibility="collapsed"
    )[cite: 1]
    
    generate_btn = st.button("🚀 Generate Full Tailored Application Pack", type="primary")[cite: 1]

if generate_btn:[cite: 1]
    if not job_desc or not job_desc.strip():[cite: 1]
        st.warning("Please paste a target Job Description first.")[cite: 1]
    elif not api_key:[cite: 1]
        st.error("API Key is missing. Please configure GEMINI_API_KEY in Streamlit Secrets.")[cite: 1]
    else:
        with col2:[cite: 1]
            with st.spinner("⚡ High-Speed Synthesis: Generating complete tailored application pack..."):[cite: 1]
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

                3. EXECUTIVE SUMMARY (STRICTLY 155 TO 170 WORDS / EXACTLY 8 FULL JUSTIFIED LINES):
                   - Write an authoritative, high-impact, passionate executive summary of EXACTLY 155 to 170 words tailored directly to the target mandate and company.
                   - It must completely fill 8 full justified lines in Calibri 10pt (line spacing multiple 1.16).
                   - Deliver a compelling, high-conviction narrative covering:
                     * 23+ years driving FMCG commercial strategy, digital commerce, and enterprise sales technology across MEA, India, and Asia.
                     * Rare 360° vantage combining principal-led FMCG commercial leadership, digital distribution entrepreneurship, and enterprise SaaS transformations with $100M+ P&L/portfolio ownership.
                     * Founding & scaling UAE's premier digital B2B2C distribution ecosystem (Conektr), aggregating extensive catalogs across ambient, packaged goods, and personal care brands to 8,000+ retailers.
                     * Leading enterprise technology advisory & SFA/DMS modernizations (FieldAssist & Ivy Mobility) for 10+ Tier-1 CPG leaders (P&G, Nestlé, GSK, Coca-Cola), delivering ~40% logistics cost reduction, >50% drop in coverage cost, and ~20% productivity uplifts.

                4. CAPABILITY ORDERING (PRIORITIZATION):
                   - Rank the 5 capability keys based on the JD's highest priorities (place the top 2 matching keys first):
                     Available keys: ["commercial", "digital", "transformation", "capability", "entrepreneurship"]
                   - "capability_order": An array containing all 5 keys in ordered priority.

                5. CONEKTR CATEGORY BULLET:
                   - Category aggregation bullet strictly tailored to the products/domain of the target company.

                6. DYNAMIC EXPERIENCE INJECTIONS (STRICT 18 TO 24 WORDS EACH):
                   - "column_1_extra_bullet": 18-24 words under Britannia / Traditional FMCG regarding Route-to-Market, distributor governance, or commercial expansion aligned with JD, else "".
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
                  "column_1_extra_bullet": "string",
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
                """[cite: 1]
                
                tailored_data = None
                cover_data = None
                last_error = ""

                model_candidates = [
                    "gemini-2.5-flash",
                    "gemini-3.5-flash-lite",
                    "gemini-2.0-flash"
                ][cite: 1]

                client = genai.Client(api_key=api_key)[cite: 1]
                
                for model_candidate in model_candidates:[cite: 1]
                    for attempt in range(2):[cite: 1]
                        try:
                            response = client.models.generate_content(
                                model=model_candidate,
                                contents=prompt,
                                config=types.GenerateContentConfig(
                                    response_mime_type="application/json",
                                    temperature=0.2
                                )
                            )[cite: 1]
                            raw_text = response.text.strip()[cite: 1]
                            if raw_text.startswith("```"):[cite: 1]
                                raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text)[cite: 1]
                                raw_text = re.sub(r"\s*```$", "", raw_text)[cite: 1]

                            parsed_json = json.loads(raw_text)[cite: 1]
                            
                            ordered_keys = parsed_json.get("capability_order", ["commercial", "digital", "transformation", "capability", "entrepreneurship"])[cite: 1]
                            full_capabilities = [][cite: 1]
                            for k in ordered_keys:[cite: 1]
                                if k in MASTER_CAPABILITIES:[cite: 1]
                                    full_capabilities.append(MASTER_CAPABILITIES[k])[cite: 1]
                            for k, cap_text in MASTER_CAPABILITIES.items():[cite: 1]
                                if cap_text not in full_capabilities:[cite: 1]
                                    full_capabilities.append(cap_text)[cite: 1]
                            
                            parsed_json["capabilities"] = full_capabilities[cite: 1]
                            tailored_data = parsed_json[cite: 1]
                            cover_data = parsed_json.get("cover_letter_data", {})[cite: 1]
                            break
                        except Exception as e:
                            last_error = str(e)[cite: 1]
                            time.sleep(1.5 * (attempt + 1))[cite: 1]
                    if tailored_data:[cite: 1]
                        break

                if tailored_data:[cite: 1]
                    clean_docx = create_master_resume_docx(tailored_data, highlight_changes=False)[cite: 1]
                    review_docx = create_master_resume_docx(tailored_data, highlight_changes=True)[cite: 1]
                    comb_docx = create_combined_application_docx(cover_data, tailored_data)[cite: 1]
                    master_zip = create_master_application_zip(comb_docx, review_docx, clean_docx)

                    st.session_state["tailored_data"] = tailored_data[cite: 1]
                    st.session_state["cover_data"] = cover_data[cite: 1]
                    st.session_state["comb_docx"] = comb_docx[cite: 1]
                    st.session_state["review_docx"] = review_docx[cite: 1]
                    st.session_state["clean_docx"] = clean_docx
                    st.session_state["master_zip"] = master_zip[cite: 1]
                    st.session_state["has_results"] = True[cite: 1]
                else:
                    st.error(f"Generation Error: {last_error}. Please check your API key and network permissions.")[cite: 1]

# ==============================================================================
# 6. PERSISTENT DISPLAY & MULTI-FILE DOWNLOAD AREA (3-DOCX SUITE)
# ==============================================================================
if st.session_state.get("has_results", False):[cite: 1]
    with col2:[cite: 1]
        tailored_data = st.session_state["tailored_data"][cite: 1]
        cover_data = st.session_state.get("cover_data", {})[cite: 1]
        score = tailored_data.get("ats_match_score", 94)[cite: 1]
        target_co = tailored_data.get("target_company", "Target Organization")[cite: 1]
        target_rl = tailored_data.get("target_role", "Executive Position")[cite: 1]

        st.subheader(f"2. Application Pack: {target_co}")[cite: 1]
        st.caption(f"Role: **{target_rl}**")[cite: 1]

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
        """, unsafe_allow_html=True)[cite: 1]

        # 1-Click Master ZIP (3 Files)
        st.download_button(
            label=f"📦 Download Application Bundle (.ZIP) — 3 Word Files",
            data=st.session_state["master_zip"],
            file_name=f"Madhusudhanan_Janakarajan_{target_co.replace(' ', '_')}_Application_Bundle.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True
        )

        st.markdown("---")[cite: 1]
        st.write("📄 **Individual Application Files (.docx):**")

        st.download_button(
            label="📝 1. Complete Application Set (.docx) — Cover + Resume + Matrix",
            data=st.session_state["comb_docx"],
            file_name=f"1_Complete_Application_Set_Cover_Resume_Matrix.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )[cite: 1]

        col_b1, col_b2 = st.columns(2)
        with col_b1:
            st.download_button(
                label="🟡 2. Highlighted Review Resume (.docx)",
                data=st.session_state["review_docx"],
                file_name=f"2_Madhusudhanan_Janakarajan_Resume_Highlighted_Review.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        with col_b2:
            st.download_button(
                label="📄 3. Clean Master Resume (.docx)",
                data=st.session_state["clean_docx"],
                file_name=f"3_Madhusudhanan_Janakarajan_Resume_Clean.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

        with st.expander("🔍 View AI Tailored Dynamic Variables"):[cite: 1]
            st.write("**Identified Company:**", target_co)[cite: 1]
            st.write("**Identified Role:**", target_rl)[cite: 1]
            st.write("**Header Focus 1:**", tailored_data.get("header_focus_1"))[cite: 1]
            st.write("**Header Focus 2:**", tailored_data.get("header_focus_2"))[cite: 1]
            st.write("**Executive Summary:**", tailored_data.get("executive_summary"))[cite: 1]
            st.write("**Injected Bullet (Britannia/Traditional):**", tailored_data.get("column_1_extra_bullet"))
            st.write("**Injected Bullet (Conektr):**", tailored_data.get("column_2_extra_bullet"))[cite: 1]
            st.write("**Injected Bullet (TransCPG/Ivy):**", tailored_data.get("column_3_extra_bullet"))[cite: 1]
            st.write("**Match Matrix Rows Generated:**", len(cover_data.get("matrix_items", [])))[cite: 1]
