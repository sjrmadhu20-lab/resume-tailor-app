import io
import json
import os
import re
import time
import zipfile
import docx
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.shared import Inches, Pt, RGBColor
from google import genai
from google.genai import types
import streamlit as st

st.set_page_config(
    page_title="Executive ATS Application Engine", page_icon="🎯", layout="wide"
)

# ==============================================================================
# 1. API CONFIGURATION & DYNAMIC MODEL DISCOVERY
# ==============================================================================
api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))

def get_verified_model_list(client):
    canonical_fallbacks = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-2.5-pro", "gemini-1.5-pro"]
    try:
        discovered = []
        for m in client.models.list():
            raw_name = getattr(m, 'name', '') or ''
            clean_name = raw_name.replace('models/', '').strip()
            actions = getattr(m, 'supported_actions', []) or getattr(m, 'supported_generation_methods', []) or []
            if actions and 'generateContent' not in actions:
                continue
            if 'gemini' in clean_name.lower() and not any(x in clean_name for x in ['image', 'live', 'tts', 'embedding', 'computer-use', 'audio']):
                discovered.append(clean_name)
        if discovered:
            preferred = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-2.5-pro", "gemini-1.5-pro"]
            sorted_models = [m for m in preferred if m in discovered]
            sorted_models += [m for m in discovered if m not in sorted_models]
            return sorted_models
    except Exception:
        pass
    return canonical_fallbacks

def generate_with_fallback(client, contents, config, max_retries_per_model=3):
    models_to_try = get_verified_model_list(client)
    last_captured_error = None
    for model_name in models_to_try:
        for attempt in range(max_retries_per_model):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=contents,
                    config=config
                )
                return response
            except Exception as e:
                err_text = str(e)
                last_captured_error = e
                if "404" in err_text or "NOT_FOUND" in err_text:
                    break
                if any(k in err_text for k in ["503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED"]):
                    if attempt < max_retries_per_model - 1:
                        time.sleep((attempt + 1) * 2)
                        continue
                break
    raise last_captured_error

# ==============================================================================
# 2. MASTER KNOWLEDGE ARCHIVE (LOCKED BASELINE + PERPETUAL INCREMENTAL STORE)
# ==============================================================================
KNOWLEDGE_FILE = "custom_knowledge.json"

BASE_CAPABILITIES = {
    "commercial": (
        "Commercial & GTM Leadership ($100M+ P&L): Owned $100M+ annual FMCG"
        " revenue across GCC & India, directing 250+ distributors and 600+"
        " field sales teams across GT, MT, Wholesale, B2B, and Institutional"
        " channels. Spearheaded RTM redesign, distributor governance, trade"
        " margin economics, pricing/promotions, and Order-to-Cash optimization."
    ),
    "digital": (
        "Digital B2B2C Commerce, Quick Commerce & Omnichannel RTM: Founded and scaled Conektr"
        " (UAE's premier digital FMCG network) to 8,000+ B2B retailers and launched direct B2C consumer app"
        " powered by the proprietary BOSS loyalty engine (Buying, Operating, Selling & Saving)."
        " Converted mom-and-pop grocers into micro-fulfillment dark stores, optimizing retail media,"
        " digital shelf visibility, quick-commerce fulfillment, and omnichannel payment rails."
    ),
    "it_platform": (
        "Enterprise Architecture, Platform Engineering & Digital Product: Full SDLC architect"
        " as Founder-CTO of Conektr. Engineered high-load cloud B2B/B2C marketplace applications,"
        " automated ERP-DMS API integrations (Dynamics 365, SAP, Oracle), microservices, payment gateways,"
        " and AI conversational voice/WhatsApp ordering bots with high uptime and institutional security."
    ),
    "logistics_aggregation": (
        "Logistics Aggregation & Cost-to-Serve Optimization: Pioneered shared-logistics distribution"
        " aggregation at Conektr, solving fragmented drop sizes and high freight costs across UAE trade."
        " Aggregated multi-principal smaller orders into consolidated last-mile dispatches, cutting"
        " fleet logistics costs by ~40% and outlet coverage cost by >50%."
    ),
    "enterprise_sales": (
        "High-Ticket Solution Sales & Enterprise Client Acquisition: Led multi-million-dollar"
        " B2B technology and commercial software sales at Ivy Mobility and FieldAssist. Won 22+"
        " global enterprise logos (P&G, Nestlé, Coca-Cola, GSK/Haleon, Mars), serving as trusted"
        " C-level commercial advisor with high buy-in across GCC and multinational leadership."
    ),
    "transformation": (
        "Enterprise Transformation & Commercial Optimization: Directed multi-country RTM modernizations,"
        " DMS/ERP integrations and SFA deployments (Over 5,000+ Users) for global CPG leaders."
        " Deployed AI route/beat optimization, AI-driven demand forecasting, and automated ordering—delivering"
        " a ~40% drop in logistics/admin costs, >30% reduction in outlet coverage costs, ~30% frontline sales"
        " productivity uplift, and ~150% numeric distribution expansion."
    ),
    "capability": (
        "Sales Capability, Enablement & Operations Excellence: Certified Sales Trainer (CST, DOOR India"
        " Topper, SPIN Certified) with 6+ years heading regional sales training and capability operations."
        " Designed & deployed Right Store execution, BMI distributor infrastructure audits, Train-the-Trainer (TTT),"
        " 70:20:10 learning models, and territory/journey planning frameworks for 1,000+ reps across GCC & India."
    ),
    "entrepreneurship": (
        "Entrepreneurial Venture Scaling & Governance: Raised $15M in funding from DIFC VC and FMCG"
        " C-suite veterans (ex-Mondelēz President, BAT CFO), validating commercial credibility. Executed"
        " successful strategic M&A exit to Al Maya Group ($1B+ conglomerate). Awarded UAE Golden Visa and"
        " USA O-1A (Extraordinary Ability); featured in Bloomberg, Gulf News, and Magnitt."
    ),
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
        "visas": "UAE Golden Visa | USA O-1A (Extraordinary Ability)",
    },
    "honors": [
        "UAE Golden Visa – Recognized for national-scale entrepreneurship and digital commerce impact.",
        "USA O-1A Visa – Extraordinary Ability in FMCG and Digital Commerce.",
        "$15M+ VC funding & exit – Raised $15M+ and successfully exited Conektr to Al Maya Group.",
        "Featured in Gulf News, Bloomberg, Khaleej Times, Yahoo Finance, Magnitt, among others - https://linktr.ee/M_S_J",
    ],
    "education": [
        {
            "degree": "Bachelor of Engineering (2001)",
            "details": "Government College of Engineering (GEC), Tier 1 DOTE College, India",
        },
        {
            "degree": "Executive Sales & Commercial Certifications",
            "details": "Certified Sales Trainer (CST) | SPIN Technique Certified | DOOR Training All-India Topper | CREST Customer Relationship",
        },
    ],
    "languages": "English | Hindi | Tamil | Kannada | Telugu | effective engagement with Arabic-speaking stakeholders.",
    "interests": "Chess Player | Table Tennis Enthusiast | Regular 10K Runner",
    "tech_stack": {
        "AI, Automation & Conversational Commerce": (
            "Agentic Voice Bots (Vapi, ElevenLabs) | Conversational Commerce (Wati, Twilio, Infobip) |"
            " Workflow Automation (Make.com) | CRM & Marketing Automation (Klaviyo)"
        ),
        "Omnichannel & eCommerce Platforms": (
            "Amazon Brand Analytics & DSP | Noon Partner Tools | Talabat & Quick-Commerce Portals |"
            " WooCommerce | Magento | Mobile Commerce SDLC (iOS, Android, Flutter) | Retail Media & ROAS Engines"
        ),
        "Enterprise & Commercial Systems": (
            "SAP (Sales & Distribution) | Oracle eCRM | Microsoft Dynamics 365 | Enterprise SFA / DMS |"
            " ERP–CRM API Integrations"
        ),
        "Data, Analytics & Logistics Optimization": (
            "Power BI (Regional Commercial Data Hubs) | Alteryx | Tableau | Demand Forecasting Models |"
            " Dynamic Route & Fleet Beat Optimization | Logistics Cost-to-Serve Modeling"
        ),
        "Fintech & Digital Payments": (
            "Stripe | PayPal | CCAvenue | Triterras | Tabby | Spotii (trade credit, payments, and trade finance rails)"
        ),
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
        (
            " Recipient of Global recognition for FMCG Contribution: O1A from USA & Golden Visa from UAE - as an extraordinary ability leader.",
            False,
        ),
    ],
}

MASTER_DEEP_EXPERIENCE = """
CANDIDATE DEEP REPOSITORY & VERIFIED ACHIEVEMENTS:
- Inside-Out IT Advantage (Business Operator + Platform CTO + Enterprise Solution Advisor):
  * Three-Sided Technology Acumen:
    1. Operational User (Britannia, Airtel): Lived everyday frontline bottlenecks as a regional business operator managing $100M+ P&Ls, 250+ distributors, and 600+ reps. Knows exactly how users interact with technology on the ground.
    2. Platform Architect & Founder CTO (Conektr): Built the entire B2B2C digital ecosystem from architecture design to launch—spanning mobile apps, web portal, Dynamics 365 backend, automated dispatch, WhatsApp ordering, and fintech credit rails. Deployed and ran it for Conektr's own commercial operations.
    3. Enterprise Solutions Advisor (Ivy Mobility, FieldAssist, TransCPG): Trusted advisor to multinational C-suites (P&G, Nestlé, Coca-Cola), selling and implementing complex SaaS SFA/DMS across 5,000+ users.

- Route-to-Market, Logistics & Supply Aggregation:
  * Conektr Logistics Aggregator Thesis: Solved fragmented grocery supply delivery where dozens of individual principals dispatched half-empty trucks. Conektr combined orders across 100+ brands onto consolidated last-mile dispatches, cutting fleet costs by ~40% and coverage cost by >50%.
  * Strategic Relevance to Global Logistics (e.g., DP World): Master of shared multi-principal logistics, bonded hub operations, middle-mile consolidation, and digitized last-mile fulfillment.

- High-Ticket Enterprise Sales & Client Advisory:
  * Ivy Mobility & FieldAssist Track Record: Built MEA operations into 2nd largest global pipeline ($10M+). Won 22+ tier-1 enterprise accounts through value-led consultative selling.
"""

def load_custom_knowledge():
    if os.path.exists(KNOWLEDGE_FILE):
        try:
            with open(KNOWLEDGE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def save_custom_knowledge(new_entries):
    existing = load_custom_knowledge()
    existing.extend(new_entries)
    try:
        with open(KNOWLEDGE_FILE, "w", encoding="utf-8") as f:
            json.dump(existing, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False

def get_full_knowledge_context():
    custom_items = load_custom_knowledge()
    custom_text = "\n".join([f"- USER RECORDED INSIGHT: {item.get('a', '')}" for item in custom_items])
    return MASTER_DEEP_EXPERIENCE + ("\n\nDYNAMICALLY ACCUMULATED KNOWLEDGE BASE:\n" + custom_text if custom_text else "")

def clean_ai_generated_text(text):
    if not isinstance(text, str):
        return text
    replacements = [
        (r"\bthirty-six-degree\b", "360°"),
        (r"\bthirty six degree\b", "360°"),
        (r"\bthree hundred and sixty degree\b", "360°"),
        (r"\b360 degree\b", "360°"),
        (r"\b360-degree\b", "360°"),
        (r"\btwenty-three years\b", "23+ years"),
        (r"\bone hundred million dollars\b", "$100M+"),
        (r"\beight thousand\b", "8,000+"),
        (r"\bfifteen million\b", "$15M+"),
    ]
    for pattern, repl in replacements:
        text = re.sub(pattern, repl, text, flags=re.IGNORECASE)
    return text

def sanitize_json_payload(data):
    if isinstance(data, dict):
        return {k: sanitize_json_payload(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [sanitize_json_payload(item) for item in data]
    elif isinstance(data, str):
        return clean_ai_generated_text(data)
    return data

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

def apply_xml_spacing(p, before_pt=0, after_pt=8, line_twips=278):
    pPr = p._p.get_or_add_pPr()
    spPr = parse_xml(f'<w:spacing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:before="{int(before_pt*20)}" w:after="{int(after_pt*20)}" w:line="{line_twips}" w:lineRule="auto"/>')
    pPr.append(spPr)

# ==============================================================================
# 3. WORD RESUME ENGINE (DYNAMIC TRACK & 3-COLUMN ROUTING)
# ==============================================================================
def populate_resume_document(doc, tailored_data, highlight_changes=False):
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(10)
    style.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    def add_heading(title, space_before=0, space_after=8, line_border_above=False, is_multiple=False, is_underline=False):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        apply_xml_spacing(p, before_pt=space_before, after_pt=space_after, line_twips=278 if is_multiple else 240)
        if line_border_above:
            pPr = p._p.get_or_add_pPr()
            pBdr = parse_xml(r'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                             r'<w:top w:val="single" w:sz="6" w:space="4" w:color="000000"/>'
                             r'</w:pBdr>')
            pPr.append(pBdr)
        r = p.add_run(title.upper() if title != "LANGUAGES & INTERESTS :" else title)
        r.bold = True
        r.underline = is_underline
        r.font.name = 'Calibri'
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    # PAGE 1: HEADER
    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_xml_spacing(p_name, before_pt=0, after_pt=0, line_twips=278)
    r_name = p_name.add_run(MASTER_STATIC['name'])
    r_name.bold = True
    r_name.font.name = 'Calibri'
    r_name.font.size = Pt(12)

    f1 = tailored_data.get("header_focus_1", "Commercial Strategy & Market Access Lead")
    f2 = tailored_data.get("header_focus_2", "RTM & Enterprise Growth")
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_xml_spacing(p_sub, before_pt=0, after_pt=8, line_twips=278)
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
    apply_xml_spacing(p_contact, before_pt=0, after_pt=8, line_twips=278)
    p_contact.add_run(f"{c['location']} | {c['phone']} | ")
    add_hyperlink(p_contact, c['email_url'], c['email'], color_rgb="004B87", underline=True, font_size_pt=10)
    p_contact.add_run("\n")
    add_hyperlink(p_contact, c['linkedin'], c['linkedin'], color_rgb="004B87", underline=True, font_size_pt=10)
    p_contact.add_run(" | Portfolio: ")
    add_hyperlink(p_contact, c['portfolio'], c['portfolio'], color_rgb="004B87", underline=True, font_size_pt=10)
    p_contact.add_run("\n")
    r_v = p_contact.add_run("Visa Status: ")
    r_v.bold = True
    p_contact.add_run(c['visas'])

    # PAGE 1: EXECUTIVE SUMMARY (STRICT 8 TO 9 LINES CALIBRATED)
    add_heading("EXECUTIVE SUMMARY", space_before=0, space_after=6, line_border_above=False, is_multiple=True)
    sp = doc.add_paragraph()
    sp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    # 275 line twips (~1.15 multiple) with 165-190 words strictly locks to 8-9 lines
    apply_xml_spacing(sp, before_pt=0, after_pt=6, line_twips=275)
    r_sum = sp.add_run(tailored_data.get("executive_summary", ""))
    r_sum.font.name = 'Calibri'
    r_sum.font.size = Pt(10)
    if highlight_changes:
        r_sum.font.highlight_color = docx.enum.text.WD_COLOR_INDEX.YELLOW

    # PAGE 1: CAPABILITIES
    add_heading("EXECUTIVE CAPABILITIES & IMPACT HIGHLIGHTS", space_before=2, space_after=8, line_border_above=True, is_multiple=False)
    for cap in tailored_data.get("capabilities", []):
        cp = doc.add_paragraph()
        cp.paragraph_format.left_indent = Inches(0.20)
        cp.paragraph_format.first_line_indent = Inches(-0.25)
        cp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        apply_xml_spacing(cp, before_pt=0, after_pt=4.5, line_twips=240)
        cp.add_run("•\t")
        parts = cap.split(":", 1)
        if len(parts) == 2:
            r_bold = cp.add_run(parts[0] + ":")
            r_bold.bold = True
            cp.add_run(parts[1])
        else:
            cp.add_run(cap)

    # PAGE 1: HONORS
    add_heading("HONORS & RECOGNITION", space_before=2, space_after=8, line_border_above=True, is_multiple=False)
    for idx, h in enumerate(MASTER_STATIC['honors']):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Inches(0.45)
        p.paragraph_format.first_line_indent = Inches(-0.20)
        apply_xml_spacing(p, before_pt=0, after_pt=6 if idx == len(MASTER_STATIC['honors']) - 1 else 0, line_twips=240)
        p.add_run("•\t")
        if "https://" in h:
            parts = h.split(" - ")
            p.add_run(parts[0] + " - ")
            add_hyperlink(p, parts[1].strip(), parts[1].strip(), color_rgb="004B87", underline=True, font_size_pt=10)
        else:
            p.add_run(h)

    # PAGE 1: EDUCATION
    add_heading("EDUCATION & CERTIFICATIONS", space_before=0, space_after=8, line_border_above=False, is_multiple=False)
    for idx, edu in enumerate(MASTER_STATIC['education']):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Inches(0.45)
        p.paragraph_format.first_line_indent = Inches(-0.20)
        apply_xml_spacing(p, before_pt=0, after_pt=6 if idx == len(MASTER_STATIC['education']) - 1 else 0, line_twips=240)
        p.add_run("•\t")
        r_bp = p.add_run(edu['degree'] + " – ")
        r_bp.bold = True
        p.add_run(edu['details'])

    # PAGE 1: LANGUAGES & INTERESTS
    add_heading("LANGUAGES & INTERESTS :", space_before=0, space_after=8, line_border_above=False, is_multiple=False)
    for text_val in [MASTER_STATIC['languages'], MASTER_STATIC['interests']]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Inches(0.45)
        p.paragraph_format.first_line_indent = Inches(-0.20)
        apply_xml_spacing(p, before_pt=0, after_pt=0, line_twips=240)
        p.add_run("•\t")
        p.add_run(text_val)

    # ---------------- PAGE 2 (DYNAMIC TRACK CONFIGURATION) ----------------
    doc.add_page_break()
    add_heading("PROFESSIONAL EXPERIENCE", space_before=0, space_after=8, line_border_above=False, is_multiple=False)
    
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

    page2_mode = tailored_data.get("page2_mode", "commercial")
    
    if page2_mode == "it_cto_digital":
        h0 = tailored_data.get("exp_col_header_1", "Business Operator & Enterprise User")
        h1 = tailored_data.get("exp_col_header_2", "Platform Architect & Founder CTO")
        h2 = tailored_data.get("exp_col_header_3", "Enterprise Tech Advisory & CxO Sales")
    elif page2_mode == "logistics_market_access":
        h0 = tailored_data.get("exp_col_header_1", "FMCG Operator & Demand Dynamics")
        h1 = tailored_data.get("exp_col_header_2", "Digital Aggregation & Logistics Optimization")
        h2 = tailored_data.get("exp_col_header_3", "Client Acquisition & Enterprise Advisory")
    elif page2_mode == "capability":
        h0 = tailored_data.get("exp_col_header_1", "Sales Operations")
        h1 = tailored_data.get("exp_col_header_2", "Sales Capability - Traditional")
        h2 = tailored_data.get("exp_col_header_3", "Sales Capability - Digital")
    elif page2_mode == "ecomm_b2c":
        h0 = tailored_data.get("exp_col_header_1", "Omnichannel & Commercial Operations")
        h1 = tailored_data.get("exp_col_header_2", "Digital B2B2C & Quick-Commerce")
        h2 = tailored_data.get("exp_col_header_3", "Digital Shelf & Transformation")
    else:
        h0 = tailored_data.get("exp_col_header_1", "Traditional FMCG Operator")
        h1 = tailored_data.get("exp_col_header_2", "Digital FMCG Distribution")
        h2 = tailored_data.get("exp_col_header_3", "Distribution Transformation")

    for i, title in enumerate([h0, h1, h2]):
        cell = table.rows[0].cells[i]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_xml_spacing(p, before_pt=3, after_pt=3, line_twips=240)
        r = p.add_run(title)
        r.bold = True
        r.font.name = 'Calibri'
        r.font.size = Pt(11)
        cell._tc.get_or_add_tcPr().append(parse_xml(r'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="DCE6F1"/>'))

    def populate_cell_content(cell, item_list):
        cell.text = ""
        for item in item_list:
            p = cell.add_paragraph()
            apply_xml_spacing(p, before_pt=item.get("space_before", 0), after_pt=item.get("space_after", 2), line_twips=220)
            if item.get("is_bullet", False):
                p.paragraph_format.left_indent = Inches(0.18)
                p.paragraph_format.first_line_indent = Inches(-0.15)
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r_b = p.add_run("•\t")
                r_b.font.name = 'Calibri'
                r_b.font.size = Pt(10)
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(item["text"])
            r.bold = item.get("bold", False)
            r.italic = item.get("italic", False)
            r.font.name = 'Calibri'
            r.font.size = Pt(10)
            if highlight_changes and item.get("highlight", False):
                r.font.highlight_color = docx.enum.text.WD_COLOR_INDEX.YELLOW

    if page2_mode == "it_cto_digital":
        c0_items = [
            {"text": "Britannia Industries Ltd | 2007 – 2011", "bold": True, "size": 10, "space_before": 2},
            {"text": "Regional Sales Head – GCC", "bold": True, "size": 10},
            {"text": "Regional Sales & Capability Head- India", "bold": True, "size": 10, "space_after": 4},
            {"text": "Frontline Business Operator: Owned $100M+ P&L across 6 GCC countries, experiencing end-user operational pain points firsthand.", "is_bullet": True, "size": 10},
            {"text": "Directed 250+ distributor networks and 600+ reps, setting daily transaction, ERP reconciliation, and Order-to-Cash workflows.", "is_bullet": True, "size": 10},
            {"text": "Spearheaded Britannia's 1st national SFA rollout (1,000+ users), ensuring real-world user adoption and data fidelity.", "is_bullet": True, "size": 10, "space_after": 3},
            {"text": "Airtel | Reliance | Tyco | 2001 – 2007", "bold": True, "size": 10, "space_before": 5},
            {"text": "Commercial & Capability Roles", "bold": True, "size": 10, "space_after": 4},
            {"text": "Integrated Oracle CRM and enterprise billing systems across high-velocity telecom retail showrooms and distribution.", "is_bullet": True, "size": 10}
        ]
        c1_items = [
            {"text": "Digital FMCG Platform", "bold": True, "size": 10, "space_before": 2},
            {"text": "Conektr Tech Global Ltd | UAE & India", "bold": True, "size": 10, "space_after": 4},
            {"text": "Founder & Platform CTO / CEO", "bold": True, "size": 10},
            {"text": "May 2016 – Aug 2024", "size": 10, "space_after": 4},
            {"text": "Architected end-to-end cloud platform serving 8,000+ B2B retailers (2,000+ MAU), managing full SDLC from Figma to production.", "is_bullet": True, "size": 10},
            {"text": "Inside-Out IT Mastery: Built the technology and operated it daily as the primary commercial user for our own business.", "is_bullet": True, "size": 10},
            {"text": "Engineered multi-tenant architecture, automated Dynamics 365 + Power BI integrations, and AI route optimization.", "is_bullet": True, "size": 10},
            {"text": "Built conversational WhatsApp ordering bots and integrated payment rails (Stripe, CCAvenue, Tabby), scaling to ~AED 50M GMV.", "is_bullet": True, "size": 10},
            {"text": "Raised $15M from VC and FMCG C-suite veterans; executed successful strategic M&A exit to Al Maya Group.", "is_bullet": True, "size": 10}
        ]
        c2_items = [
            {"text": "TransCPG & FieldAssist | 2025 – Present", "bold": True, "size": 10, "space_before": 2},
            {"text": "Board Advisor – Commercial Tech", "bold": True, "size": 10, "space_after": 3},
            {"text": "Advising C-suites on enterprise digital architectures, regional Power BI data hubs, and microservices integration.", "is_bullet": True, "size": 10},
            {"text": "Built Bid2Bill AI conversational voice-bot platform, lowering integration friction and cutting onboarding time by ~40%.", "is_bullet": True, "size": 10, "space_after": 3},
            {"text": "Ivy Mobility Pte Ltd | 2011 – 2016", "bold": True, "size": 10, "space_before": 5},
            {"text": "Business Head – MEA", "bold": True, "size": 10, "space_after": 4},
            {"text": "Led enterprise solution sales to CxOs of tier-1 global brands (P&G, Nestlé, Coca-Cola, GSK/Haleon, Mars).", "is_bullet": True, "size": 10},
            {"text": "Directed large-scale cloud SFA/DMS implementations across 5,000+ field users with >95% adoption rates.", "is_bullet": True, "size": 10}
        ]
    elif page2_mode == "logistics_market_access":
        c0_items = [
            {"text": "Britannia Industries Ltd | 2007 – 2011", "bold": True, "size": 10, "space_before": 2},
            {"text": "Regional Sales Head – GCC", "bold": True, "size": 10},
            {"text": "Regional Sales & Capability Head- India", "bold": True, "size": 10, "space_after": 4},
            {"text": "Directed $100M+ P&L across 6 GCC markets, mastering regional supply chain, trade corridors, and distributor economics.", "is_bullet": True, "size": 10},
            {"text": "Governed 250+ distributor networks, setting drop-size thresholds, credit governance, and working capital cycles.", "is_bullet": True, "size": 10},
            {"text": "Engineered route-to-market overhauls, eliminating logistical bottlenecks and cutting distribution admin costs by ~30%.", "is_bullet": True, "size": 10},
            {"text": "Turnaround RSM GCC: Honored with Best Employee Award by Group MD for record volume deliveries.", "is_bullet": True, "size": 10, "space_after": 3},
            {"text": "Airtel | Reliance | Tyco | 2001 – 2007", "bold": True, "size": 10, "space_before": 5},
            {"text": "Commercial & Institutional Tender Roles", "bold": True, "size": 10, "space_after": 4},
            {"text": "Managed large-scale institutional supply contracts, tender governance, and enterprise security distributions.", "is_bullet": True, "size": 10}
        ]
        c1_items = [
            {"text": "Digital FMCG Principal / Distributor", "bold": True, "size": 10, "space_before": 2},
            {"text": "Conektr Tech Global Ltd | UAE & India", "bold": True, "size": 10, "space_after": 4},
            {"text": "Chief Executive Officer & Founder", "bold": True, "size": 10},
            {"text": "May 2016 – Aug 2024", "size": 10, "space_after": 4},
            {"text": "Solved fragmented delivery economics by operating as a multi-principal digital logistics aggregator for 8,000+ grocers.", "is_bullet": True, "size": 10},
            {"text": "Eliminated supplier fleet duplication by aggregating orders from Red Bull, Coca-Cola, and Unilever onto shared routes.", "is_bullet": True, "size": 10},
            {"text": "Cut coverage costs by >50% and logistics transit expenses by ~40% via AI dynamic routing and warehouse optimization.", "is_bullet": True, "size": 10},
            {"text": "Scaled annual GMV to ~AED 50M (~$13.6M) at ~18% gross margin; successfully executed M&A exit to Al Maya Group.", "is_bullet": True, "size": 10}
        ]
        c2_items = [
            {"text": "TransCPG & FieldAssist | 2025 – Present", "bold": True, "size": 10, "space_before": 2},
            {"text": "Board Advisor – Commercial Solutions", "bold": True, "size": 10, "space_after": 3},
            {"text": "Advising enterprise boards on supply chain integration, regional data hubs, and modern route logistics.", "is_bullet": True, "size": 10},
            {"text": "High-conviction FMCG C-level credibility, unlocking multi-stakeholder buy-in across global brands and regional distributors.", "is_bullet": True, "size": 10, "space_after": 3},
            {"text": "Ivy Mobility Pte Ltd | 2011 – 2016", "bold": True, "size": 10, "space_before": 5},
            {"text": "Business Head – MEA", "bold": True, "size": 10, "space_after": 4},
            {"text": "Spearheaded enterprise client acquisition, securing 22+ high-ticket logos (P&G, Nestlé, Coca-Cola, GSK/Haleon, Mars).", "is_bullet": True, "size": 10},
            {"text": "Built $10M+ enterprise pipeline across 10+ countries, acting as senior trusted advisor for supply chain digitization.", "is_bullet": True, "size": 10}
        ]
    elif page2_mode == "ecomm_b2c":
        c0_items = [
            {"text": "Britannia Industries Ltd | 2007 – 2011", "bold": True, "size": 10, "space_before": 2},
            {"text": "Regional Sales Head – GCC", "bold": True, "size": 10},
            {"text": "Regional Sales & Capability Head- India", "bold": True, "size": 10, "space_after": 4},
            {"text": "Owned $100M+ P&L across 6 GCC countries; directed key account joint business planning (JBP) with top modern trade.", "is_bullet": True, "size": 10},
            {"text": "Governed 250+ distributor networks, trade spend ROI, customer agreements, and margin parity frameworks.", "is_bullet": True, "size": 10},
            {"text": "Pioneered Britannia's first automated SFA rollout (1,000+ reps), boosting field strike rates and LPC to ~120%.", "is_bullet": True, "size": 10, "space_after": 3},
            {"text": "Airtel | Reliance | Tyco | 2001 – 2007", "bold": True, "size": 10, "space_before": 5},
            {"text": "Commercial Trade & Channel Roles", "bold": True, "size": 10, "space_after": 4},
            {"text": "Built foundations in multi-channel distribution, CRM integrations, and category promotional activations.", "is_bullet": True, "size": 10}
        ]
        c1_items = [
            {"text": "Digital FMCG Principal / Distributor", "bold": True, "size": 10, "space_before": 2},
            {"text": "Conektr Tech Global Ltd | UAE & India", "bold": True, "size": 10, "space_after": 4},
            {"text": "Chief Executive Officer & Founder", "bold": True, "size": 10},
            {"text": "May 2016 – Aug 2024", "size": 10, "space_after": 4},
            {"text": "Founded UAE's first digital FMCG distributor; scaled platform to 8,000+ B2B retailers and 2,000+ MAU.", "is_bullet": True, "size": 10},
            {"text": "Expanded into B2C quick commerce: launched consumer app and BOSS engine, turning retail network into dark stores.", "is_bullet": True, "size": 10},
            {"text": "Orchestrated brand exclusivity contracts and digital shelf presence for Red Bull, Unilever, and Coca-Cola.", "is_bullet": True, "size": 10},
            {"text": "Supervised retail media budgets, driving high-impact digital merchandising, search placement, and superior ROAS.", "is_bullet": True, "size": 10},
            {"text": "Scaled annual GMV to ~AED 50M (~$13.6M) at ~18% gross margin; executed strategic M&A exit to Al Maya Group.", "is_bullet": True, "size": 10}
        ]
        c2_items = [
            {"text": "TransCPG & FieldAssist | 2025 – Present", "bold": True, "size": 10, "space_before": 2},
            {"text": "Transformation Advisor (Director)", "bold": True, "size": 10, "space_after": 3},
            {"text": "Designed enterprise Power BI commercial dashboards and regional data hubs connecting primary, secondary, and tertiary retail data.", "is_bullet": True, "size": 10},
            {"text": "Integrated Bid2Bill AI/voice-bot and WhatsApp commerce ordering engine, cutting customer acquisition cost by ~40%.", "is_bullet": True, "size": 10, "space_after": 3},
            {"text": "Ivy Mobility Pte Ltd | 2011 – 2016", "bold": True, "size": 10, "space_before": 5},
            {"text": "Business Head – MEA", "bold": True, "size": 10, "space_after": 4},
            {"text": "Won 22 enterprise logos (P&G, Nestlé, GSK/Haleon, Coca-Cola), modernizing omnichannel route-to-market systems.", "is_bullet": True, "size": 10},
            {"text": "Led on-ground mobile retail execution rollouts for P&G distributor networks, ensuring complete digital adoption.", "is_bullet": True, "size": 10}
        ]
    elif page2_mode == "capability":
        c0_items = [
            {"text": "Britannia Industries Ltd | 2008 – 2011", "bold": True, "size": 10, "space_before": 2},
            {"text": "Regional Sales Manager (RSM) – GCC", "bold": True, "size": 10, "space_after": 3},
            {"text": "Directed sales operations and RTM across 6 GCC markets generating $100M+ NSV across GT, MT, and wholesale.", "is_bullet": True, "size": 10},
            {"text": "Right Store Optimization: Implemented outlet profiling, call frequency compliance, and ~30% numeric distribution gain.", "is_bullet": True, "size": 10},
            {"text": "Turnaround RSM: Shattered consecutive quarterly records; honored with Best Employee Award by Group MD.", "is_bullet": True, "size": 10, "space_after": 3},
            {"text": "Conektr Tech Global | 2016 – 2024", "bold": True, "size": 10, "space_before": 4},
            {"text": "Chief Executive Officer & Founder", "bold": True, "size": 10, "space_after": 3},
            {"text": "Led complete distribution ops, P&L, supply chain, and trade margin architecture for 8,000+ B2B grocery outlets.", "is_bullet": True, "size": 10},
            {"text": "Optimized RTM cost-to-serve by >30% using hybrid distribution, micro-dark stores, and dynamic routing.", "is_bullet": True, "size": 10}
        ]
        c1_items = [
            {"text": "Britannia Industries Ltd | 2007 – 2008", "bold": True, "size": 10, "space_before": 2},
            {"text": "Regional Sales Capability Head – India", "bold": True, "size": 10, "space_after": 3},
            {"text": "Led capability architecture across South 1 & South 2 regions covering 200+ distributors & 600+ reps.", "is_bullet": True, "size": 10},
            {"text": "Designed BMI (Business Measurement Index) DMS system for distributor infrastructure & sales standard audits.", "is_bullet": True, "size": 10},
            {"text": "Established TTT, Star Rating, Champion Scorecards, and trade 'Profit Clubs' lifting LPC to ~120%.", "is_bullet": True, "size": 10, "space_after": 3},
            {"text": "Bharti Airtel Ltd | 2005 – 2007", "bold": True, "size": 10, "space_before": 4},
            {"text": "Circle Sales Training Manager", "bold": True, "size": 10, "space_after": 3},
            {"text": "Established Karnataka Circle Training Wing; deployed induction, buddy programs, and SPIN selling techniques.", "is_bullet": True, "size": 10}
        ]
        c2_items = [
            {"text": "TransCPG & FieldAssist | 2025 – Present", "bold": True, "size": 10, "space_before": 2},
            {"text": "Board Advisor – Commercial Tech", "bold": True, "size": 10, "space_after": 3},
            {"text": "Advising CPG boards on RTM modernization, AI beat planning, and Power BI commercial analytics hubs.", "is_bullet": True, "size": 10},
            {"text": "Integrated AI conversational coaching bots (Bid2Bill), reducing sales onboarding cycles by ~40%.", "is_bullet": True, "size": 10, "space_after": 3},
            {"text": "Ivy Mobility Pte Ltd | 2011 – 2016", "bold": True, "size": 10, "space_before": 4},
            {"text": "Business Head – MEA", "bold": True, "size": 10, "space_after": 3},
            {"text": "Deployed enterprise SaaS SFA/DMS across 22 top logos (P&G, Nestlé, Red Bull, GSK/Haleon, Coca-Cola).", "is_bullet": True, "size": 10},
            {"text": "Led on-ground mobile sales tool enablement for P&G Kenya distributor force, ensuring 100% field adoption.", "is_bullet": True, "size": 10}
        ]
    else:
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
        conektr_cat = tailored_data.get("conektr_category_bullet", "Deep FMCG Category Aggregation: Scaled multi-category catalogs across ambient, packaged food, and consumer goods portfolios.")
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

    c0_extra = tailored_data.get("column_1_extra_bullet", "")
    if c0_extra and c0_extra.strip():
        c0_items.insert(6, {"text": c0_extra.strip(), "is_bullet": True, "size": 10, "highlight": True})

    c1_extra = tailored_data.get("column_2_extra_bullet", "")
    if c1_extra and c1_extra.strip():
        c1_items.insert(6, {"text": c1_extra.strip(), "is_bullet": True, "size": 10, "highlight": True})

    c2_extra = tailored_data.get("column_3_extra_bullet", "")
    if c2_extra and c2_extra.strip():
        c2_items.insert(5, {"text": c2_extra.strip(), "is_bullet": True, "size": 10, "highlight": True})

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

    # PAGE 2: TECH STACK
    add_heading("TECHNOLOGY STACK & DIGITAL ARCHITECTURE:", space_before=8, space_after=8, line_border_above=False, is_multiple=False)
    for category, stack in MASTER_STATIC['tech_stack'].items():
        tp = doc.add_paragraph()
        tp.paragraph_format.left_indent = Inches(0.20)
        tp.paragraph_format.first_line_indent = Inches(-0.25)
        tp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        apply_xml_spacing(tp, before_pt=0, after_pt=4, line_twips=240)
        tp.add_run("•\t")
        r_cat = tp.add_run(f"{category}: ")
        r_cat.bold = True
        tp.add_run(stack)

    # PAGE 2: WHY HIRE ME
    add_heading("WHY HIRE ME", space_before=8, space_after=4, line_border_above=False, is_multiple=False, is_underline=True)
    p_why = doc.add_paragraph()
    p_why.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    apply_xml_spacing(p_why, before_pt=0, after_pt=6, line_twips=240)
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
# 4. WORD COVER LETTER & 10PT MATCH MATRIX
# ==============================================================================
def populate_cover_letter_docx_page(doc, cover_data):
    # Direct Opening: Single 'Dear Hiring Team,' paragraph
    p_d = doc.add_paragraph("Dear Hiring Team,")
    apply_xml_spacing(p_d, before_pt=8, after_pt=10, line_twips=260)
    p_d.runs[0].bold = True
    p_d.runs[0].font.size = Pt(11)

    # Sanitize cover_para_1 to ensure NO duplicate salutation generated by AI
    p1_text = cover_data.get("cover_para_1", "").strip()
    p1_cleaned = re.sub(r"^(?:dear\s+hiring\s+team\s*,\s*)+", "", p1_text, flags=re.IGNORECASE).strip()

    p_p1 = doc.add_paragraph(p1_cleaned)
    apply_xml_spacing(p_p1, before_pt=0, after_pt=8, line_twips=270)
    p_p1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    p_p2 = doc.add_paragraph(cover_data.get("cover_para_2", ""))
    apply_xml_spacing(p_p2, before_pt=0, after_pt=8, line_twips=270)
    p_p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    p_kh = doc.add_paragraph()
    apply_xml_spacing(p_kh, before_pt=0, after_pt=6, line_twips=260)
    r_kh = p_kh.add_run("Key highlights of what I bring to this mandate include:")
    r_kh.bold = True
    r_kh.font.name = 'Calibri'
    r_kh.font.size = Pt(10.5)

    for b in cover_data.get("cover_bullets", []):
        bp = doc.add_paragraph()
        bp.paragraph_format.left_indent = Inches(0.25)
        bp.paragraph_format.first_line_indent = Inches(-0.18)
        apply_xml_spacing(bp, before_pt=0, after_pt=5, line_twips=250)
        bp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        bp.add_run("•\t")
        parts = b.split(":", 1)
        if len(parts) == 2:
            r_head = bp.add_run(parts[0] + ": ")
            r_head.bold = True
            bp.add_run(parts[1].strip())
        else:
            bp.add_run(b)

    p_cl = doc.add_paragraph(cover_data.get("cover_para_closing", ""))
    apply_xml_spacing(p_cl, before_pt=4, after_pt=8, line_twips=270)
    p_cl.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    p_sign = doc.add_paragraph()
    apply_xml_spacing(p_sign, before_pt=6, after_pt=0, line_twips=240)
    p_sign.add_run("Sincerely,\n")
    r_s1 = p_sign.add_run("Madhusudhanan Janakarajan (Madhu)\n")
    r_s1.bold = True
    p_sign.add_run("+971 50 654 7858 | sjrmadhu20@gmail.com")

def populate_match_matrix_docx_page(doc, cover_data):
    """
    Renders an impactful MATCH MATRIX formatted strictly at 10pt font,
    with column height >= 0.6 inches per row to prevent cramped rendering.
    """
    p_mtitle = doc.add_paragraph()
    p_mtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_xml_spacing(p_mtitle, before_pt=0, after_pt=2, line_twips=220)
    r_mt = p_mtitle.add_run("MATCH MATRIX")
    r_mt.bold = True
    r_mt.font.name = 'Calibri'
    r_mt.font.size = Pt(12)

    dynamic_sub = cover_data.get(
        "matrix_subtitle",
        "Granular cross-enterprise alignment of commercial leadership and demonstrated track record against mandate priorities."
    )
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_xml_spacing(p_sub, before_pt=0, after_pt=6, line_twips=200)
    r_sub = p_sub.add_run(dynamic_sub)
    r_sub.italic = True
    r_sub.font.name = 'Calibri'
    r_sub.font.size = Pt(9.5)

    matrix_items = cover_data.get("matrix_items", [])
    table = doc.add_table(rows=len(matrix_items) + 1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    col_w0 = Inches(2.35)
    col_w1 = Inches(5.15)

    cell_0 = table.rows[0].cells[0]
    cell_1 = table.rows[0].cells[1]
    cell_0.width = col_w0
    cell_1.width = col_w1
    cell_0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    cell_1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    shd_xml = r'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="DCE6F1"/>'
    cell_0._tc.get_or_add_tcPr().append(parse_xml(shd_xml))
    cell_1._tc.get_or_add_tcPr().append(parse_xml(shd_xml))

    p_h0 = cell_0.paragraphs[0]
    p_h0.alignment = WD_ALIGN_PARAGRAPH.LEFT
    apply_xml_spacing(p_h0, before_pt=3, after_pt=3, line_twips=220)
    r_h0 = p_h0.add_run("Target Mandate Requirement")
    r_h0.bold = True
    r_h0.font.name = 'Calibri'
    r_h0.font.size = Pt(10)

    p_h1 = cell_1.paragraphs[0]
    p_h1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    apply_xml_spacing(p_h1, before_pt=3, after_pt=3, line_twips=220)
    r_h1 = p_h1.add_run("Demonstrated Track Record & Proof Points")
    r_h1.bold = True
    r_h1.font.name = 'Calibri'
    r_h1.font.size = Pt(10)

    for idx, item in enumerate(matrix_items):
        row = table.rows[idx + 1]
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(r'<w:cantSplit xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
        trPr.append(parse_xml(r'<w:trHeight xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="864" w:hRule="atLeast"/>'))

        r_cells = row.cells
        r_cells[0].width = col_w0
        r_cells[1].width = col_w1
        r_cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        r_cells[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        p0 = r_cells[0].paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
        apply_xml_spacing(p0, before_pt=3, after_pt=3, line_twips=220)
        r_rt = p0.add_run(item.get('requirement_title', ''))
        r_rt.bold = True
        r_rt.font.name = 'Calibri'
        r_rt.font.size = Pt(10)

        p1 = r_cells[1].paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        apply_xml_spacing(p1, before_pt=3, after_pt=3, line_twips=220)
        r_mt = p1.add_run(item.get('match_desc', ''))
        r_mt.font.name = 'Calibri'
        r_mt.font.size = Pt(10)

    tblBorders = parse_xml(
        r'<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        r'<w:top w:val="single" w:sz="6" w:space="0" w:color="004B87"/>'
        r'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="004B87"/>'
        r'<w:insideH w:val="single" w:sz="3" w:space="0" w:color="E0E0E0"/>'
        r'<w:insideV w:val="single" w:sz="3" w:space="0" w:color="D3D3D3"/>'
        r'</w:tblBorders>'
    )
    table._tbl.tblPr.append(tblBorders)

def create_combined_application_docx(cover_data, tailored_data):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.35)
        section.bottom_margin = Inches(0.35)
        section.left_margin = Inches(0.50)
        section.right_margin = Inches(0.50)
    populate_cover_letter_docx_page(doc, cover_data)
    doc.add_page_break()
    populate_resume_document(doc, tailored_data, highlight_changes=False)
    doc.add_page_break()
    populate_match_matrix_docx_page(doc, cover_data)
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io.getvalue()

def create_master_application_zip(comb_docx, review_docx, clean_docx):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr('1_Complete_Application_Set_Cover_Resume_Matrix.docx', comb_docx)
        zip_file.writestr('2_Madhusudhanan_Janakarajan_Resume_Highlighted_Review.docx', review_docx)
        zip_file.writestr('3_Madhusudhanan_Janakarajan_Resume_Clean.docx', clean_docx)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

def rebuild_all_documents():
    tailored_data = st.session_state["tailored_data"]
    cover_data = st.session_state["cover_data"]
    clean_docx = create_master_resume_docx(tailored_data, highlight_changes=False)
    review_docx = create_master_resume_docx(tailored_data, highlight_changes=True)
    comb_docx = create_combined_application_docx(cover_data, tailored_data)
    master_zip = create_master_application_zip(comb_docx, review_docx, clean_docx)
    st.session_state["comb_docx"] = comb_docx
    st.session_state["review_docx"] = review_docx
    st.session_state["clean_docx"] = clean_docx
    st.session_state["master_zip"] = master_zip

# ==============================================================================
# 5. STREAMLIT FRONTEND & DYNAMIC CONTROLLER
# ==============================================================================
st.title("🎯 Executive ATS Resume & Application Engine")
st.caption("Commercial Leadership • Digital Transformation • Adaptive Multi-Track Experience Architecture")

with st.sidebar:
    st.header("⚡ System Status")
    if api_key:
        st.success("🟢 Gemini AI Engine: Active")
    else:
        st.error("🔴 AI Engine Key Missing (Set GEMINI_API_KEY in Secrets)")
    st.markdown("---")
    stored_facts = load_custom_knowledge()
    st.write(f"📚 **Dynamic Custom Facts Stored:** {len(stored_facts)}")
    if st.button("🔄 Reset / Clear Cached Session"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

col1, col2 = st.columns([1.1, 0.9])

with col1:
    st.subheader("1. Job Inputs & Context")
    job_desc = st.text_area(
        "Target Job Description (JD):",
        height=220,
        placeholder="Paste target Job Description here (e.g., IT & Digital Development Manager, DP World Lead, Clorox eCommerce)...",
    )
    st.markdown("##### Special Instructions & Context (Auto-Fed to Permanent Knowledge Base)")
    st.caption("Any context entered here is permanently integrated into your custom knowledge base for all future resumes.")
    special_instructions = st.text_area(
        "Voice or Typed Notes:",
        height=90,
        placeholder="e.g., Highlight my 3-sided IT advantage: Operator at Britannia/Airtel, Founder-CTO who built & used Conektr, and CxO Sales at FieldAssist/Ivy...",
    )

    col_btn1, col_btn2 = st.columns([1, 1])
    with col_btn1:
        generate_btn = st.button("🚀 Generate Tailored Application Pack", type="primary", use_container_width=True)
    with col_btn2:
        gap_analyze_btn = st.button("🧠 Analyze JD for Knowledge Gaps", type="secondary", use_container_width=True)

    if gap_analyze_btn:
        if not job_desc.strip():
            st.warning("Please paste a target Job Description first to analyze gaps.")
        else:
            with st.spinner("Analyzing JD against candidate master knowledge archive..."):
                full_archive_text = get_full_knowledge_context()
                gap_prompt = f"""
                Compare the target Job Description against Madhusudhanan Janakarajan's master knowledge archive:
                {full_archive_text}
                TARGET JOB DESCRIPTION:
                {job_desc}
                Identify any gaps or formulate 3 to 5 targeted questions to strengthen the match.
                """
                client = genai.Client(api_key=api_key)
                try:
                    gap_resp = generate_with_fallback(
                        client=client,
                        contents=gap_prompt,
                        config=types.GenerateContentConfig(temperature=0.2)
                    )
                    st.session_state["gap_questions"] = gap_resp.text.strip()
                except Exception as e:
                    st.error(f"Gap Analysis Error: {str(e)}")

    if st.session_state.get("gap_questions"):
        st.markdown("---")
        st.info(st.session_state["gap_questions"])

# ==============================================================================
# MAIN GENERATION CONTROLLER
# ==============================================================================
if generate_btn:
    if not job_desc or not job_desc.strip():
        st.warning("Please paste a target Job Description first.")
    elif not api_key:
        st.error("API Key is missing. Please configure GEMINI_API_KEY in Streamlit Secrets.")
    else:
        # Step 1: Auto-persist special instructions to knowledge base if provided
        if special_instructions.strip():
            save_custom_knowledge([{"q": "Dynamic Prompt Context", "a": special_instructions.strip()}])

        with col2:
            with st.spinner("⚡ Tailoring Application Suite: Aligning Domain Track, Experience Narrative, and Match Matrix..."):
                archive_context = get_full_knowledge_context()
                
                prompt = f"""
                You are an executive resume architect and career strategist for Madhusudhanan Janakarajan.
                Analyze the target Job Description (JD) and special instructions. Cross-reference with the candidate archive.

                CANDIDATE MASTER KNOWLEDGE ARCHIVE:
                {archive_context}

                TARGET TRACK ROUTING & PAGE 2 COLUMN DYNAMICS:
                - CRITICAL TRACK EVALUATION:
                  1. IF IT, Software Architecture, Digital Product, CTO, Digital Development Manager, Tech Transformation (e.g., Arla IT Lead, CTO):
                     - Set "page2_mode": "it_cto_digital"
                     - Set "exp_col_header_1": "Business Operator & Enterprise User"
                     - Set "exp_col_header_2": "Platform Architect & Founder CTO"
                     - Set "exp_col_header_3": "Enterprise Tech Advisory & CxO Sales"
                     - Set "capability_order": ["it_platform", "transformation", "commercial", "enterprise_sales", "entrepreneurship"]

                  2. IF Logistics, Market Access, Port Operations, 3PL, Supply Chain, or Aggregator mandate (e.g., DP World):
                     - Set "page2_mode": "logistics_market_access"
                     - Set "exp_col_header_1": "FMCG Operator & Demand Dynamics"
                     - Set "exp_col_header_2": "Digital Aggregation & Logistics Optimization"
                     - Set "exp_col_header_3": "Client Acquisition & Enterprise Advisory"
                     - Set "capability_order": ["commercial", "logistics_aggregation", "enterprise_sales", "transformation", "entrepreneurship"]

                  3. IF E-Commerce, Digital Shelf, Retail Media, Direct-to-Consumer, Quick Commerce (e.g., Clorox):
                     - Set "page2_mode": "ecomm_b2c"
                     - Set "exp_col_header_1": "Omnichannel & Commercial Operations"
                     - Set "exp_col_header_2": "Digital B2B2C & Quick-Commerce"
                     - Set "exp_col_header_3": "Digital Shelf & Transformation"
                     - Set "capability_order": ["digital", "commercial", "transformation", "capability", "entrepreneurship"]

                  4. IF Sales Capability, Training, Sales Ops, or RTM Excellence (e.g., Mondelēz):
                     - Set "page2_mode": "capability"
                     - Set "exp_col_header_1": "Sales Operations"
                     - Set "exp_col_header_2": "Sales Capability - Traditional"
                     - Set "exp_col_header_3": "Sales Capability - Digital"
                     - Set "capability_order": ["capability", "commercial", "transformation", "digital", "entrepreneurship"]

                  5. OTHERWISE (General Management / FMCG Leadership):
                     - Set "page2_mode": "commercial"
                     - Set "exp_col_header_1": "Traditional FMCG Operator"
                     - Set "exp_col_header_2": "Digital FMCG Distribution"
                     - Set "exp_col_header_3": "Distribution Transformation"

                COVER LETTER SPECIFICS:
                - In "cover_para_1", DO NOT include any salutation like 'Dear Hiring Team,'. START DIRECTLY with the first sentence of your opening pitch (e.g., "As an executive technologist and commercial operator...").
                - DO NOT include headers like 'COVER LETTER' or 'Subject:'.
                - Paragraph 1: Catchy, authoritative opening tailored to the company, demonstrating how candidate's unique cross-domain mastery solves their exact challenge.

                EXECUTIVE SUMMARY (STRICT REQUIREMENT):
                - MUST be strictly between 165 and 190 words.
                - This word count guarantees it occupies AT LEAST 8 TO 9 FULL LINES in 10pt justified Calibri. Do NOT make it short.

                MATCH MATRIX SPECIFICS:
                - Heading MUST simply be "MATCH MATRIX".
                - Dynamic subtitle: Generate a targeted 1-sentence subtitle describing alignment with the specific JD priorities.
                - Column Headers: "Target Mandate Requirement" and "Demonstrated Track Record & Proof Points". NEVER use words like "Candidate Evidence".
                - Exactly 9 to 10 rows. Each match description must be 28 to 36 words, dense with verified metrics ($100M+ NSV, 250+ distributors, 8,000+ stores, ~30% ND, etc.) and correlate at least two of the candidate's experiences.

                STRICT EXECUTIVE WRITING RULES:
                - Never spell out numbers into words (Always use: '360°', '$100M+', '23+ years', '8,000+', '~40%').

                INPUT JOB DESCRIPTION:
                {job_desc}

                INPUT SPECIAL INSTRUCTIONS:
                {special_instructions}

                Return ONLY a valid JSON object matching this schema:
                {{
                  "target_company": "string",
                  "target_role": "string",
                  "page2_mode": "it_cto_digital",
                  "header_focus_1": "string",
                  "header_focus_2": "string",
                  "executive_summary": "string",
                  "capability_order": ["it_platform", "transformation", "commercial", "enterprise_sales", "entrepreneurship"],
                  "exp_col_header_1": "string",
                  "exp_col_header_2": "string",
                  "exp_col_header_3": "string",
                  "conektr_category_bullet": "string",
                  "column_1_extra_bullet": "string",
                  "column_2_extra_bullet": "string",
                  "column_3_extra_bullet": "string",
                  "ats_match_score": 96,
                  "cover_letter_data": {{
                    "target_company": "string",
                    "cover_para_1": "string",
                    "cover_para_2": "string",
                    "cover_bullets": ["string", "string", "string", "string"],
                    "cover_para_closing": "string",
                    "matrix_subtitle": "string",
                    "matrix_items": [
                      {{
                        "requirement_title": "string",
                        "match_desc": "string"
                      }}
                    ]
                  }}
                }}
                """

                client = genai.Client(api_key=api_key)
                tailored_data = None
                cover_data = None
                last_error = ""

                try:
                    response = generate_with_fallback(
                        client=client,
                        contents=prompt,
                        config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.2),
                    )
                    raw_text = response.text.strip()
                    if raw_text.startswith("```"):
                        raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text)
                        raw_text = re.sub(r"\s*```$", "", raw_text)

                    parsed_json = json.loads(raw_text)
                    parsed_json = sanitize_json_payload(parsed_json)

                    ordered_keys = parsed_json.get(
                        "capability_order",
                        ["it_platform", "transformation", "commercial", "enterprise_sales", "entrepreneurship"],
                    )
                    full_capabilities = [BASE_CAPABILITIES[k] for k in ordered_keys if k in BASE_CAPABILITIES]
                    for k, cap_text in BASE_CAPABILITIES.items():
                        if cap_text not in full_capabilities and len(full_capabilities) < 5:
                            full_capabilities.append(cap_text)

                    parsed_json["capabilities"] = full_capabilities
                    tailored_data = parsed_json
                    cover_data = parsed_json.get("cover_letter_data", {})
                except Exception as e:
                    last_error = str(e)

                if tailored_data:
                    st.session_state["tailored_data"] = tailored_data
                    st.session_state["cover_data"] = cover_data
                    st.session_state["job_desc"] = job_desc
                    rebuild_all_documents()
                    st.session_state["has_results"] = True
                else:
                    st.error(f"Generation Error: {last_error}")

# ==============================================================================
# 6. RESULTS & REVISION CONTROLLER
# ==============================================================================
if st.session_state.get("has_results", False):
    with col2:
        tailored_data = st.session_state["tailored_data"]
        cover_data = st.session_state.get("cover_data", {})
        score = tailored_data.get("ats_match_score", 96)
        target_co = tailored_data.get("target_company", "Target Organization")
        target_rl = tailored_data.get("target_role", "Executive Position")

        st.subheader(f"2. Application Pack: {target_co}")
        st.caption(f"Role: **{target_rl}** | Track Routing: **{tailored_data.get('page2_mode', 'commercial').upper()}**")

        st.download_button(
            label="📦 Download Application Bundle (.ZIP) — 3 Word Files",
            data=st.session_state["master_zip"],
            file_name=f"Madhusudhanan_Janakarajan_{target_co.replace(' ', '_')}_Application_Bundle.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True,
        )

        st.download_button(
            label="📝 1. Complete Application Set (.docx) — Cover + Resume + Matrix",
            data=st.session_state["comb_docx"],
            file_name="1_Complete_Application_Set_Cover_Resume_Matrix.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

        col_b1, col_b2 = st.columns(2)
        with col_b1:
            st.download_button(
                label="🟡 2. Highlighted Review Resume (.docx)",
                data=st.session_state["review_docx"],
                file_name="2_Madhusudhanan_Janakarajan_Resume_Highlighted_Review.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with col_b2:
            st.download_button(
                label="📄 3. Clean Master Resume (.docx)",
                data=st.session_state["clean_docx"],
                file_name="3_Madhusudhanan_Janakarajan_Resume_Clean.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

        st.markdown("---")
        st.subheader("✍️ Instant Revisions & Perpetual Feedback")
        st.caption("Tweaks entered below update the current document and are saved permanently into your custom knowledge store.")

        correction_text = st.text_area("Enter corrections or adjustments:", height=85)

        if st.button("🔄 Apply Revisions & Learn", type="secondary"):
            if not correction_text.strip():
                st.warning("Please type your feedback first.")
            else:
                save_custom_knowledge([{"q": f"Feedback for {target_co}", "a": correction_text.strip()}])
                
                with st.spinner("⚡ Applying targeted revisions and saving to master store..."):
                    revise_prompt = f"""
                    Refine the existing application JSON for Madhusudhanan Janakarajan.
                    CURRENT APPLICATION DATA:
                    {json.dumps(st.session_state["tailored_data"])}
                    USER REVISION REQUEST:
                    {correction_text}

                    STRICT RULES:
                    - Keep executive summary strictly between 165 and 190 words (MUST be 8 to 9 lines in Calibri 10pt).
                    - In "cover_para_1", DO NOT include "Dear Hiring Team,".
                    - In MATCH MATRIX, use 'Demonstrated Track Record & Proof Points'. Keep each row 28-36 words with quantitative metrics.
                    - Maintain all numeric conventions ('$100M+', '23+ years', '8,000+', '~40%').
                    Return ONLY valid JSON.
                    """
                    client = genai.Client(api_key=api_key)
                    try:
                        rev_resp = generate_with_fallback(
                            client=client,
                            contents=revise_prompt,
                            config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.2),
                        )
                        rev_raw = rev_resp.text.strip()
                        if rev_raw.startswith("```"):
                            rev_raw = re.sub(r"^```(?:json)?\s*", "", rev_raw)
                            rev_raw = re.sub(r"\s*```$", "", rev_raw)
                        rev_json = json.loads(rev_raw)
                        rev_json = sanitize_json_payload(rev_json)
                        st.session_state["tailored_data"] = rev_json
                        st.session_state["cover_data"] = rev_json.get("cover_letter_data", {})
                        rebuild_all_documents()
                        st.success("✅ Revisions applied and permanently saved to knowledge base!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Failed to revise: {str(e)}")
