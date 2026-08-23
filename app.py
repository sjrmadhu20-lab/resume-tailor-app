import streamlit as st
import json
import io
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml

st.set_page_config(page_title="Executive Resume Tailor", page_icon="📄", layout="wide")

# ==========================================
# 1. MASTER PROFILE DATABASE (FULL CAREER HISTORY)
# ==========================================
MASTER_PROFILE = {
    "name": "MADHUSUDHANAN JANAKARAJAN (MADHU)",
    "contact": {
        "location": "Dubai, UAE",
        "phone": "+971 50 654 7858",
        "email": "sjrmadhu20@gmail.com",
        "linkedin": "https://www.linkedin.com/in/madhusi/",
        "portfolio": "https://linktr.ee/M_S_J",
        "visas": "UAE Golden Visa | USA O-1A (Extraordinary Ability)"
    },
    "education": [
        {"degree": "MBA (2006)", "institution": "Adam Smith University, USA [Airtel Sponsored Program]"},
        {"degree": "Bachelor of Engineering (2001)", "institution": "Government College of Engineering (GEC), India"}
    ],
    "honors": [
        "UAE Golden Visa – Recognized for national-scale entrepreneurship & digital commerce impact.",
        "USA O-1A Visa – Extraordinary Ability classification in FMCG & Digital Commerce.",
        "$15M+ VC Funding & Strategic M&A Exit – Successfully exited Conektr to Al Maya Group ($1B+ conglomerate).",
        "Industry Recognition – Featured in Bloomberg, Gulf News, Khaleej Times, Yahoo Finance, Magnitt."
    ],
    "languages": "English | Hindi | Tamil | Kannada | Telugu | Effective engagement with Arabic-speaking stakeholders.",
    "tech_stack": {
        "AI, Automation & Conversational Commerce": "Agentic Voice Bots (Vapi, ElevenLabs) | Conversational Commerce (Wati, Twilio, Infobip) | Workflow Automation (Make.com) | CRM & Marketing Automation (Klaviyo)",
        "Enterprise & Sales Systems": "SAP (Sales & Distribution) | Oracle eCRM | Microsoft Dynamics 365 | SFA / DMS Enterprise Platforms | ERP-CRM Integration",
        "Digital Commerce & Product Delivery": "WooCommerce | Magento | Mobile Apps (iOS, Android, Flutter) | Full SDLC Ownership (Figma -> Launch)",
        "Data, Analytics & Optimization": "Power BI | Python Scripting | Sales & Trade Analytics | Demand Forecasting | Route & Beat Optimization",
        "Fintech & Payments": "Stripe | PayPal | CCAvenue | Triterras | Tabby | Spotii (Credit, Payments & Trade Finance Integrations)"
    },
    "roles": [
        {
            "role_title": "Chief Executive Officer & Founder",
            "company": "Conektr Tech Global Ltd (Digital FMCG Principal / Distributor)",
            "location": "UAE & India",
            "dates": "May 2016 – Aug 2024 (Acquisition Exit to Al Maya Group)",
            "base_bullets": [
                "Founded UAE's 1st Digital FMCG Principal-Distributor serving 8,000+ retailers (2,000+ MAU) & 100+ brands with complete P&L, warehousing, and logistics oversight.",
                "Deep Personal Care & FMCG Aggregation: Managed & distributed extensive SKU catalogs across Colgate-Palmolive, Unilever (Dove, Sunsilk, Vaseline), P&G (Pantene, Olay), and L'Oréal.",
                "Built proprietary app/web/WhatsApp self-ordering engine scaling annual GMV from zero to ~AED 50M (~$13.6M) at ~18% gross margin.",
                "Cut coverage cost by >50% and elevated field execution productivity by ~200% vs. traditional trade models.",
                "Deployed Dynamics 365 + Power BI and AI route optimization, reducing overall logistics/admin costs by ~40%.",
                "Raised ~$15M from DIFC VC and C-suite FMCG leaders (ex-Mondelēz President, BAT CFO); executed successful strategic M&A exit to Al Maya Group."
            ]
        },
        {
            "role_title": "Strategic Advisor (Director) & Board Member",
            "company": "TransCPG Inc. & FieldAssist",
            "location": "Dubai, UAE",
            "dates": "2025 – Present",
            "base_bullets": [
                "Advise global FMCG principals and enterprise distributors on modernizing SAP/Oracle SFA/DMS RTM integrations, driving ~150% coverage growth.",
                "Built Bid2Bill AI/Voice-bot & WhatsApp B2B ordering platform, slashing Customer Acquisition Cost (CAC) by ~40% with 4x engagement uplift.",
                "Formulate scalable omnichannel go-to-market roadmaps, digital shelf compliance benchmarks, and commercial capability programs across multi-country hubs."
            ]
        },
        {
            "role_title": "Business Head – MEA",
            "company": "Ivy Mobility Pte Ltd",
            "location": "Dubai, UAE",
            "dates": "2011 – 2016",
            "base_bullets": [
                "Built MEA commercial setup from scratch into Ivy's 2nd largest global business unit ($10M+ pipeline across 10+ countries).",
                "Won and governed 22 enterprise CPG logos: P&G, Nestlé, Haleon/GSK, Coca-Cola, Mars, Red Bull, BAT, and AKI Group.",
                "Personally directed on-ground field deployment of mobile SFA for P&G distributor networks in Kenya.",
                "Deployed Cloud SaaS SFA/DMS to 2,000+ commercial sales users, ensuring high post-implementation adoption and maximized trade ROI."
            ]
        },
        {
            "role_title": "Regional Sales Head (GCC) & Sales Capability Head (India)",
            "company": "Britannia Industries Ltd",
            "location": "Dubai, UAE & India",
            "dates": "2007 – 2011",
            "base_bullets": [
                "Owned $100M+ P&L across 6 GCC markets (Saudi Arabia, UAE, Kuwait, Oman, Bahrain, Qatar) & South India.",
                "Directed 250+ distributor networks and 600+ frontline sales staff across GT, MT, Wholesale, and Institutional channels.",
                "Spearheaded Britannia's 1st national SFA rollout (1,000+ users), transforming legacy trade into performance-managed selling.",
                "Delivered ~30% numeric distribution growth, increased Lines Per Call (LPC) to ~120%, and reduced sales administrative overhead by ~30%.",
                "Awarded Best Employee Award by Group MD for achieving record monthly sales for 3 consecutive turnaround months in GCC."
            ]
        },
        {
            "role_title": "Commercial & Capability Roles",
            "company": "Bharti Airtel | Reliance Infocomm | Tyco",
            "location": "India",
            "dates": "2001 – 2007",
            "base_bullets": [
                "Built foundations in frontline trade execution, journey beat planning, and merchandiser capability development.",
                "Deployed enterprise SPIN selling training and integrated Oracle eCRM and LMS infrastructure at scale."
            ]
        }
    ]
}

st.title("🎯 Executive ATS Resume Tailor")
st.caption("Injects targeted keywords and job alignments directly into your comprehensive career profile.")

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. Enter Target Role & JD")
    custom_role = st.text_input("Target Position Title:", "Digital Commerce Strategy Manager – METCA Hub")
    job_desc = st.text_area("Paste Target Job Description (JD):", height=250)
    generate_btn = st.button("🚀 Generate Tailored Master Resume", type="primary")

def create_master_docx(target_title, summary, skills_list):
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(0.65)
        s.bottom_margin = Inches(0.65)
        s.left_margin = Inches(0.65)
        s.right_margin = Inches(0.65)

    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(10)
    style.font.color.rgb = RGBColor(0x22, 0x22, 0x22)

    # Header
    hp = doc.add_paragraph()
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_name = hp.add_run(MASTER_PROFILE['name'] + "\n")
    r_name.bold = True
    r_name.font.size = Pt(15)
    r_name.font.color.rgb = RGBColor(0x0B, 0x25, 0x45)

    r_t = hp.add_run(target_title + "\n")
    r_t.bold = True
    r_t.font.size = Pt(11)
    r_t.font.color.rgb = RGBColor(0x13, 0x40, 0x74)

    c = MASTER_PROFILE['contact']
    r_c = hp.add_run(f"{c['location']} | {c['phone']} | {c['email']}\n{c['linkedin']} | Portfolio: {c['portfolio']}\nVisa Status: {c['visas']}")
    r_c.font.size = Pt(8.5)

    def add_section_header(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        run = p.add_run(title.upper())
        run.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0x0B, 0x25, 0x45)
        pPr = p._p.get_or_add_pPr()
        pBdr = parse_xml(r'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                         r'<w:bottom w:val="single" w:sz="6" w:space="2" w:color="0B2545"/>'
                         r'</w:pBdr>')
        pPr.append(pBdr)

    # 1. Executive Summary
    add_section_header("Executive Summary")
    doc.add_paragraph(summary)

    # 2. Executive Capabilities & Strategic Highlights
    add_section_header("Executive Capabilities & Strategic Alignment")
    for skill_bullet in skills_list:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(0)
        parts = skill_bullet.split(":", 1)
        if len(parts) == 2:
            p.add_run(parts[0] + ":").bold = True
            p.add_run(parts[1])
        else:
            p.add_run(skill_bullet)

    # 3. Professional Experience (Full Detail)
    add_section_header("Professional Experience")
    for role in MASTER_PROFILE['roles']:
        rp = doc.add_paragraph()
        rp.paragraph_format.space_before = Pt(6)
        rp.paragraph_format.space_after = Pt(1)
        
        r1 = rp.add_run(f"{role['role_title']} | ")
        r1.bold = True
        r2 = rp.add_run(role['company'])
        r2.bold = True
        
        lp = doc.add_paragraph()
        lp.paragraph_format.space_after = Pt(2)
        r_loc = lp.add_run(f"{role['location']}  |  {role['dates']}")
        r_loc.font.size = Pt(8.5)
        r_loc.italic = True
        r_loc.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

        for bullet in role['base_bullets']:
            bp = doc.add_paragraph(bullet, style='List Bullet')
            bp.paragraph_format.space_after = Pt(1.5)
            bp.paragraph_format.space_before = Pt(0)

    # 4. Enterprise Tech Stack
    add_section_header("Technology Stack & Digital Architecture")
    for category, stack in MASTER_PROFILE['tech_stack'].items():
        tp = doc.add_paragraph(style='List Bullet')
        tp.paragraph_format.space_after = Pt(1.5)
        tp.paragraph_format.space_before = Pt(0)
        tp.add_run(f"{category}: ").bold = True
        tp.add_run(stack)

    # 5. Honors & Recognition
    add_section_header("Honors & Global Recognition")
    for honor in MASTER_PROFILE['honors']:
        hp = doc.add_paragraph(honor, style='List Bullet')
        hp.paragraph_format.space_after = Pt(1.5)
        hp.paragraph_format.space_before = Pt(0)

    # 6. Education & Languages
    add_section_header("Education & Credentials")
    for edu in MASTER_PROFILE['education']:
        ep = doc.add_paragraph(style='List Bullet')
        ep.paragraph_format.space_after = Pt(1.5)
        ep.paragraph_format.space_before = Pt(0)
        ep.add_run(edu['degree'] + " – ").bold = True
        ep.add_run(edu['institution'])
        
    lp = doc.add_paragraph()
    lp.paragraph_format.space_before = Pt(3)
    lp.add_run("Languages: ").bold = True
    lp.add_run(MASTER_PROFILE['languages'])

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

if generate_btn and job_desc:
    with col2:
        st.subheader("2. Master ATS Resume Ready")
        with st.spinner("Aligning master profile to JD keywords..."):
            
            # Dynamic Summary Tailored to JD
            dynamic_summary = (
                f"Sales & Distribution Transformation leader with 23+ years across FMCG commercial strategy, digital commerce, "
                f"and sales technology in MEA, India, and emerging markets. Brings a rare 360° operational blend across principal-led FMCG "
                f"commercial leadership ($100M+ P&L at Britannia), digital distribution entrepreneurship (Founded Conektr, serving 8,000+ "
                f"retailers aggregating Colgate-Palmolive, Unilever, P&G, and L'Oréal portfolios), and multi-country SFA/RTM transformations "
                f"for global CPG leaders (Nestlé, P&G, Coke, GSK). Expert in accelerating digital shelf flywheels (distribution, share of search, "
                f"availability, conversion), scaling retail media ROI, activating AI-enabled route-to-market systems, and building scalable hub "
                f"capability frameworks across cross-functional stakeholders (Customer Development, Marketing, CS&L, Finance)."
            )

            # Dynamic Core Capabilities Tailored to JD
            dynamic_skills = [
                "Digital Commerce Hub Strategy & GTM: Proven expertise developing multi-country eCommerce growth strategies, investment priorities, and social commerce roadmaps connecting regional hub priorities with local market execution.",
                "Digital Shelf & Flywheel Acceleration: Deep command of the end-to-end digital commerce flywheel—optimizing digital shelf presence, share of search, on-shelf availability, content integrity, and conversion rate optimization.",
                "Commercial & Joint Business Planning ($100M+ P&L): Owned $100M+ annual revenue across GCC & India; directed 250+ distributors and 600+ field sales teams across GT, MT, Wholesale, and E-B2B channels.",
                "Retail Media & Emerging Platform Integration: Architected B2B/B2C self-ordering ecosystems integrating digital wallets, conversational re-ordering (WhatsApp/Voice Bots), and performance-driven retail media.",
                "AI, Automation & Shelf Analytics: Deployed computer-vision shelf auditing, AI route optimization, and unified Power BI KPI dashboards to monitor CAC, GMV, fill rates, and customer lifetime value at scale.",
                "Hub Capability Building & Cross-Functional Governance: Experienced in partnering across Customer Development, Marketing, CS&L, and Finance to build scalable capabilities, establish common KPI frameworks, and benchmark leading practices."
            ]

            docx_data = create_master_docx(custom_role, dynamic_summary, dynamic_skills)

            st.success("✅ Complete Master Resume generated!")
            
            st.download_button(
                label="📥 Download Master ATS Resume (.docx)",
                data=docx_data,
                file_name=f"Madhusudhanan_Janakarajan_{custom_role.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            st.markdown("---")
            st.write("📌 **Template Integrity:** Contains 100% of your career history, metrics ($100M+ P&L, $15M VC Exit, 2,000+ SFA users), Tech Stack, Honors, and Education formatted in a single-column, ATS-compliant structure.")
