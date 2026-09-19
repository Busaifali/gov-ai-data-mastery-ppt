import os
import smtplib
import json
import re
import time
import random
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from google import genai
from google.genai import errors

# جلب النماذج المتاحة لحسابك تلقائياً لتفادي أخطاء 404
def get_supported_models(client):
    try:
        discovered = []
        for m in client.models.list():
            m_name = getattr(m, "name", "")
            clean_name = m_name.replace("models/", "")
            if any(k in clean_name for k in ["flash", "pro"]) and not any(x in clean_name for x in ["embed", "imagen", "vision", "realtime"]):
                discovered.append(clean_name)
        if discovered:
            print(f"[*] النماذج المكتشفة والمدعومة لحسابك: {discovered[:5]}")
            return discovered
    except Exception as e:
        print(f"[!] تعذر جلب قائمة النماذج تلقائياً ({e})، سيتم استخدام القائمة الافتراضية.")
    return ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]

# استدعاء Gemini مع التراجع الأسي وسلسلة النماذج المكتشفة
def generate_with_fallback(client, prompt, candidate_models):
    for model_name in candidate_models:
        print(f"[*] محاولة التوليد عبر النموذج: {model_name}...")
        for attempt in range(1, 4):
            try:
                res = client.models.generate_content(
                    model=model_name,
                    contents=prompt
                )
                if res and res.text:
                    print(f"[✔] تم استخراج المحتوى بنجاح عبر {model_name}!")
                    return res.text
            except (errors.ServerError, errors.APIError, Exception) as e:
                err_msg = str(e)
                print(f"[!] تنبيه في {model_name} (محاولة {attempt}): {err_msg[:80]}...")
                if "404" in err_msg or "NOT_FOUND" in err_msg:
                    break # الانتقال فوراً للموديل التالي إذا كان غير موجود
                if attempt < 3:
                    sleep_time = (2 ** attempt) + random.uniform(1.0, 2.5)
                    time.sleep(sleep_time)
    return None

# استخراج كائن JSON من رد النموذج بدقة
def parse_slides_json(raw_text):
    if not raw_text:
        return None
    try:
        # البحث عن أول مصفوفة JSON [ ... ]
        match = re.search(r'\[\s*\{.*\}\s*\]', raw_text, re.DOTALL)
        if match:
            return json.loads(match.group(0))
        # محاولة تنظيف الكود البرمجي المباشر
        clean = raw_text.strip()
        if clean.startswith("```json"):
            clean = clean[7:]
        elif clean.startswith("```"):
            clean = clean[3:]
        if clean.endswith("```"):
            clean = clean[:-3]
        return json.loads(clean.strip())
    except Exception as e:
        print(f"[!] فشل تحليل الـ JSON: {e}")
        return None

# بناء ملف الباوربوينت بتصميم عريض 16:9 وبطاقات مزدوجة
def create_deck(filename, title_theme, subtitle_theme, slides_data, primary_rgb, accent_rgb):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # 1. شريحة الغلاف الرئيسية
    slide_title = prs.slides.add_slide(blank_layout)
    banner = slide_title.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.4))
    banner.fill.solid()
    banner.fill.fore_color.rgb = primary_rgb
    banner.line.color.rgb = primary_rgb

    txBox = slide_title.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(3.2))
    tf = txBox.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.text = title_theme
    p.font.size = Pt(38)
    p.font.bold = True
    p.font.name = "Arial"
    p.font.color.rgb = primary_rgb

    p2 = tf.add_paragraph()
    p2.text = subtitle_theme
    p2.font.size = Pt(20)
    p2.font.name = "Arial"
    p2.font.color.rgb = accent_rgb

    p3 = tf.add_paragraph()
    today_str = datetime.now().strftime("%B %d, %Y")
    p3.text = f"Executive Governance Series | Published: {today_str}"
    p3.font.size = Pt(13)
    p3.font.name = "Arial"
    p3.font.color.rgb = RGBColor(120, 144, 156)

    # 2. شرائح المحتوى العشر
    for idx, item in enumerate(slides_data, start=1):
        slide = prs.slides.add_slide(blank_layout)

        # الشريط العلوي
        ribbon = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
        ribbon.fill.solid()
        ribbon.fill.fore_color.rgb = primary_rgb
        ribbon.line.color.rgb = primary_rgb

        # رقم الشريحة
        badge = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.2), Inches(0.7), Inches(0.7))
        badge.fill.solid()
        badge.fill.fore_color.rgb = accent_rgb
        badge.line.color.rgb = accent_rgb
        bp = badge.text_frame.paragraphs[0]
        bp.text = f"{idx:02d}"
        bp.font.size = Pt(18)
        bp.font.bold = True
        bp.font.color.rgb = RGBColor(255, 255, 255)
        bp.alignment = PP_ALIGN.CENTER

        # عنوان الشريحة
        title_box = slide.shapes.add_textbox(Inches(1.7), Inches(0.15), Inches(11.0), Inches(0.8))
        tp = title_box.text_frame.paragraphs[0]
        tp.text = item.get("title", f"Dimension {idx}")
        tp.font.size = Pt(22)
        tp.font.bold = True
        tp.font.name = "Arial"
        tp.font.color.rgb = RGBColor(255, 255, 255)

        # البطاقة اليسرى: المفهوم والإطار
        card1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(5.6), Inches(5.5))
        card1.fill.solid()
        card1.fill.fore_color.rgb = RGBColor(245, 247, 250)
        card1.line.color.rgb = RGBColor(220, 226, 235)

        tb_c1 = slide.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(5.2), Inches(5.1))
        tf_c1 = tb_c1.text_frame
        tf_c1.word_wrap = True

        h1 = tf_c1.paragraphs[0]
        h1.text = "CORE PRINCIPLE & ARCHITECTURE"
        h1.font.size = Pt(13)
        h1.font.bold = True
        h1.font.name = "Arial"
        h1.font.color.rgb = primary_rgb

        desc_p = tf_c1.add_paragraph()
        desc_p.text = item.get("concept", "")
        desc_p.font.size = Pt(11.5)
        desc_p.font.name = "Arial"
        desc_p.font.color.rgb = RGBColor(55, 71, 79)

        align_h = tf_c1.add_paragraph()
        align_h.text = "\nFRAMEWORK & STANDARD ALIGNMENT:"
        align_h.font.size = Pt(12)
        align_h.font.bold = True
        align_h.font.name = "Arial"
        align_h.font.color.rgb = accent_rgb

        align_p = tf_c1.add_paragraph()
        align_p.text = item.get("framework_alignment", "")
        align_p.font.size = Pt(11)
        align_p.font.name = "Arial"
        align_p.font.color.rgb = RGBColor(69, 90, 100)

        # البطاقة اليمنى: القياس والتطبيق
        card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.4), Inches(5.7), Inches(5.5))
        card2.fill.solid()
        card2.fill.fore_color.rgb = RGBColor(255, 255, 255)
        card2.line.color.rgb = RGBColor(200, 210, 225)

        tb_c2 = slide.shapes.add_textbox(Inches(7.0), Inches(1.6), Inches(5.3), Inches(5.1))
        tf_c2 = tb_c2.text_frame
        tf_c2.word_wrap = True

        h2 = tf_c2.paragraphs[0]
        h2.text = "DEEP-DIVE MEASUREMENT & AUDIT"
        h2.font.size = Pt(13)
        h2.font.bold = True
        h2.font.name = "Arial"
        h2.font.color.rgb = primary_rgb

        bullets = item.get("deep_dive_metrics", [])
        for b in bullets:
            bp_item = tf_c2.add_paragraph()
            bp_item.text = f"• {b}"
            bp_item.font.size = Pt(11)
            bp_item.font.name = "Arial"
            bp_item.font.color.rgb = RGBColor(38, 50, 56)

        bp_h = tf_c2.add_paragraph()
        bp_h.text = "\nSTRATEGIC ACTION / KPI:"
        bp_h.font.size = Pt(12)
        bp_h.font.bold = True
        bp_h.font.name = "Arial"
        bp_h.font.color.rgb = accent_rgb

        bp_p = tf_c2.add_paragraph()
        bp_p.text = item.get("implementation_kpi", "")
        bp_p.font.size = Pt(11)
        bp_p.font.name = "Arial"
        bp_p.font.color.rgb = RGBColor(55, 71, 79)

    prs.save(filename)
    print(f"[✔] تم بنجاح إنشاء ملف العرض: {filename}")

# بيانات احتياطية ممتازة لضمان توليد الملفات حتى لو تعطل الاتصال السحابي
DEFAULT_AI_SLIDES = [
    {
        "title": "AIMS Organizational Context & Scope",
        "concept": "Establishing an Artificial Intelligence Management System (AIMS) requires defining organizational objectives, risk boundaries, and assessing the ethical implications of autonomous decision-making systems.",
        "framework_alignment": "ISO/IEC 42001 Clause 4 (Context of the Organization) & NIST AI RMF 'GOVERN' Function 1.1.",
        "deep_dive_metrics": ["Define AI application inventory and classification tiers.", "Identify internal and external stakeholders affected by AI deployment.", "Document regulatory and ethical boundaries across all AI models."],
        "implementation_kpi": "100% of production AI systems registered with defined operational boundaries."
    },
    {
        "title": "AI Risk Assessment & Impact Profiling",
        "concept": "Systematic evaluation of risks throughout the AI lifecycle, evaluating safety, potential bias, security threats, and performance degradation in production environments.",
        "framework_alignment": "ISO/IEC 42001 Clause 6.1 (Actions to address risks) & NIST AI RMF 'MAP' Function.",
        "deep_dive_metrics": ["Establish AI Risk Thresholds across socio-technical domains.", "Conduct adversarial robustness testing against prompt injections.", "Map algorithmic impact assessments for high-stakes decisions."],
        "implementation_kpi": "Risk classification matrix completed for all internal and citizen-facing LLMs."
    },
    {
        "title": "Algorithmic Transparency & Explainability",
        "concept": "Ensuring decisions made or assisted by AI models are interpretable by human auditors, stakeholders, and end-users with traceable reasoning paths.",
        "framework_alignment": "ISO/IEC 42001 Annex A.7 (AI System Transparency) & NIST AI RMF 'MEASURE' 2.8.",
        "deep_dive_metrics": ["Implement post-hoc explainability techniques (SHAP, LIME).", "Maintain continuous audit trails of model inputs, prompts, and inference outputs.", "Establish transparent user disclosure whenever AI agents interact with citizens."],
        "implementation_kpi": "Explainability documentation signed off for all tier-1 automated decisions."
    },
    {
        "title": "Data Governance for AI Training & RAG",
        "concept": "Rigorous verification of training datasets, retrieval-augmented generation (RAG) vector embeddings, ensuring provenance, license compliance, and bias mitigation.",
        "framework_alignment": "ISO/IEC 42001 Annex A.6 (Data for AI systems) & DAMA-DMBOK Data Quality.",
        "deep_dive_metrics": ["Audit dataset representation to prevent demographic and systemic bias.", "Validate data lineage from ingestion to vector database embeddings.", "Enforce data privacy anonymization before model fine-tuning."],
        "implementation_kpi": "Zero data provenance defects detected in enterprise vector knowledge bases."
    },
    {
        "title": "Continuous Model Monitoring & Drift Mitigation",
        "concept": "Post-deployment monitoring mechanisms to continuously evaluate concept drift, data drift, hallucination rates, and performance degradation over time.",
        "framework_alignment": "ISO/IEC 42001 Clause 9 (Performance evaluation) & NIST AI RMF 'MANAGE' 3.1.",
        "deep_dive_metrics": ["Real-time tracking of latency, token consumption, and refusal rates.", "Automated anomaly alerts when semantic hallucination index exceeds 2.5%.", "Periodic re-evaluation against static baseline benchmark datasets."],
        "implementation_kpi": "Model drift evaluation conducted on a weekly automated cadence."
    },
    {
        "title": "AI Cybersecurity & Adversarial Robustness",
        "concept": "Defending AI pipelines against emerging attack vectors including model inversion, data poisoning, prompt extraction, and indirect prompt injection attacks.",
        "framework_alignment": "ISO/IEC 42001 Annex A.8 (AI Safety) & OWASP Top 10 for LLM Applications.",
        "deep_dive_metrics": ["Deploy automated input guardrails and jailbreak detection filters.", "Enforce strict isolation between model execution and underlying system shell.", "Conduct scheduled red-teaming exercises against Agentic workflows."],
        "implementation_kpi": "Zero critical vulnerabilities identified during quarterly AI red-teaming."
    },
    {
        "title": "Human Oversight & Intervention Mechanisms",
        "concept": "Ensuring meaningful human-in-the-loop (HITL) and human-on-the-loop governance to override, alter, or halt automated AI actions when thresholds fail.",
        "framework_alignment": "ISO/IEC 42001 Annex A.9 (Human Oversight) & UAE National AI Guidelines.",
        "deep_dive_metrics": ["Implement emergency kill-switch protocols for autonomous agents.", "Mandate human sign-off for actions involving legal or financial commitments.", "Log all human override decisions to continuously refine model calibration."],
        "implementation_kpi": "Emergency halt response time verified within under 60 seconds."
    },
    {
        "title": "AI Ethics & Fairness Auditing",
        "concept": "Structured ethical auditing ensuring systems do not amplify disparities, respect citizen dignity, and comply with national ethical AI charters.",
        "framework_alignment": "ISO/IEC 42001 Annex A.5 (Ethical Policies) & UNESCO AI Ethics Framework.",
        "deep_dive_metrics": ["Compute disparate impact ratios across user cohorts.", "Establish an independent institutional AI Ethics Review Board.", "Publish annual transparency reports detailing AI governance performance."],
        "implementation_kpi": "100% compliance with UAE National Charter for AI Ethics."
    },
    {
        "title": "Supplier & Third-Party AI Model Governance",
        "concept": "Evaluating external commercial LLMs, foundation models, and vendor AI integrations against security, privacy, and sovereignty benchmarks.",
        "framework_alignment": "ISO/IEC 42001 Clause 8.2 (Third-Party Relations) & Enterprise Procurement.",
        "deep_dive_metrics": ["Audit vendor data retention and training-on-customer-data clauses.", "Verify model weights hosting locations and data residency guarantees.", "Evaluate vendor SLA guarantees regarding uptime and model deprecation."],
        "implementation_kpi": "All external AI vendors certified under organizational Data Privacy Addendum."
    },
    {
        "title": "AI Maturity Progression (CMMI Levels 1 to 5)",
        "concept": "Transforming ad-hoc, siloed AI experiments into an enterprise-wide, optimized capability driven by quantitative metrics and automated governance.",
        "framework_alignment": "Enterprise Architecture & CMMI Maturity Scale (Initial to Optimizing).",
        "deep_dive_metrics": ["Level 1 (Ad-Hoc): Uncoordinated experimentation.", "Level 3 (Defined): Centralized ISO 42001 policies and standard tooling.", "Level 5 (Optimizing): Self-healing pipelines and continuous autonomous governance."],
        "implementation_kpi": "Target progression to CMMI Level 4 (Quantitatively Managed) within 12 months."
    }
]

DEFAULT_DATA_SLIDES = [
    {
        "title": "DAMA Wheel: Enterprise Data Governance",
        "concept": "Data Governance serves as the foundational hub of the DAMA-DMBOK2 framework, orchestrating authority, stewardship, policy formulation, and strategy across all data assets.",
        "framework_alignment": "DAMA-DMBOK2 Core Knowledge Area: Data Governance & CMMI DMM Level 3.",
        "deep_dive_metrics": ["Establish Enterprise Data Governance Council (DGC).", "Define and assign Business Data Stewards across all operational domains.", "Formulate data privacy, retention, and acceptable usage enterprise policies."],
        "implementation_kpi": "100% of critical core business entities assigned to designated data stewards."
    },
    {
        "title": "Data Architecture & Data Lineage",
        "concept": "Defining the enterprise data blueprints, canonical data models, and mapping end-to-end data lineage from source of origin through transformation to analytical consumption.",
        "framework_alignment": "DAMA-DMBOK2 Data Architecture & Metadata Management.",
        "deep_dive_metrics": ["Implement automated column-level data lineage tracking.", "Standardize conceptual, logical, and physical data models.", "Maintain enterprise repository of API data integration points."],
        "implementation_kpi": "Automated data lineage mapped for 100% of regulatory reporting pipelines."
    },
    {
        "title": "Data Quality Dimensions & Automated Profiling",
        "concept": "Measuring and advancing data quality across six critical dimensions: Accuracy, Completeness, Consistency, Timeliness, Validity, and Uniqueness.",
        "framework_alignment": "DAMA-DMBOK2 Data Quality Management & ISO 8000.",
        "deep_dive_metrics": ["Deploy automated daily data profiling rules on ingestion.", "Establish automated alerts when invalid field rates exceed 0.5%.", "Define Data Quality SLAs with upstream application owners."],
        "implementation_kpi": "Core master customer records achieve 98%+ composite Data Quality Score."
    },
    {
        "title": "Master Data Management (MDM) & Golden Record",
        "concept": "Creating a trusted, deduplicated single source of truth ('Golden Record') for core enterprise entities such as Citizens, Commercial Licenses, and Real Estate Assets.",
        "framework_alignment": "DAMA-DMBOK2 Master & Reference Data Management.",
        "deep_dive_metrics": ["Implement probabilistic matching and deterministic merge algorithms.", "Manage Reference Data Domain Tables via centralized catalog.", "Enforce bidirectional synchronization between MDM and core transaction systems."],
        "implementation_kpi": "Entity duplication rate reduced below 0.2% across active commercial registries."
    },
    {
        "title": "Metadata Management & Data Catalogs",
        "concept": "Cataloging business, technical, and operational metadata to make data assets easily discoverable, understandable, and securely accessible across the organization.",
        "framework_alignment": "DAMA-DMBOK2 Metadata Management.",
        "deep_dive_metrics": ["Implement searchable enterprise Data Catalog with business glossaries.", "Tag all schemas with sensitivity and regulatory classifications.", "Track data freshness, volume changes, and operational execution statistics."],
        "implementation_kpi": "Over 90% of production analytical tables documented with verified business terms."
    },
    {
        "title": "Data Security, Classification & Privacy",
        "concept": "Protecting data assets throughout their lifecycle in compliance with UAE Data Protection laws, implementing classification schemas (Public, Internal, Confidential, Restricted).",
        "framework_alignment": "DAMA-DMBOK2 Data Security & UAE National Cyber Security Standards.",
        "deep_dive_metrics": ["Enforce role-based access control (RBAC) and attribute-based access (ABAC).", "Implement automated data masking and column-level encryption.", "Conduct periodic user access entitlement reviews for restricted datasets."],
        "implementation_kpi": "Zero unencrypted sensitive PII fields stored across analytical environments."
    },
    {
        "title": "Data Warehousing & Business Intelligence",
        "concept": "Architecting modern analytical platforms (Data Lakehouse, Star Schemas) that deliver timely, integrated, and reliable data for strategic executive decision-making.",
        "framework_alignment": "DAMA-DMBOK2 Data Warehousing and Business Intelligence.",
        "deep_dive_metrics": ["Implement medallion architecture (Bronze, Silver, Gold data layers).", "Enforce slow-changing dimension (SCD) standards for historical audits.", "Maintain decoupled compute and storage for elastic analytical workloads."],
        "implementation_kpi": "Daily executive dashboard refresh completed before 06:00 AM without fail."
    },
    {
        "title": "Data Storage, Operations & Lifecycle Management",
        "concept": "Managing physical database instances, backup retention policies, archiving strategies, and high-availability failover architectures to guarantee business continuity.",
        "framework_alignment": "DAMA-DMBOK2 Data Storage and Operations.",
        "deep_dive_metrics": ["Automate immutable daily database snapshots across multi-zone clouds.", "Define automated data archival rules for records older than statutory requirements.", "Validate Disaster Recovery Recovery Time Objective (RTO < 15 mins)."],
        "implementation_kpi": "Quarterly disaster recovery failover drill completed with 100% data integrity."
    },
    {
        "title": "Data Integration & Interoperability",
        "concept": "Facilitating seamless data exchange between government entities using standardized APIs, event-driven architectures (Kafka), and national federated exchange protocols.",
        "framework_alignment": "DAMA-DMBOK2 Data Integration & Interoperability.",
        "deep_dive_metrics": ["Standardize REST and GraphQL JSON schema exchange contracts.", "Implement message queue dead-letter queues to prevent payload loss.", "Enforce mutual TLS (mTLS) authentication on all inter-entity integrations."],
        "implementation_kpi": "API gateway uptime maintained at 99.95% with zero dropped transaction records."
    },
    {
        "title": "Data Maturity Assessment (DMM & CMMI Tiers)",
        "concept": "Systematically benchmarking organizational data practices against CMMI Data Management Maturity (DMM) models to transition from reactive to strategic optimization.",
        "framework_alignment": "CMMI Data Management Maturity (DMM) & DAMA Maturity Model.",
        "deep_dive_metrics": ["Level 1: Ad-hoc, isolated data silos.", "Level 2: Managed at individual project levels.", "Level 4 & 5: Measured quantitatively with predictive optimization and value realization."],
        "implementation_kpi": "Advance overall enterprise DMM maturity score from Level 2.8 to Level 3.8."
    }
]

# ===============================
# بدء تنفيذ سير العمل الرئيسي
# ===============================
now = datetime.now()
today_str = now.strftime("%Y-%m-%d")

client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])
supported_models = get_supported_models(client)

# 1. توليد العرض الأول: حوكمة الذكاء الاصطناعي
prompt_ai = """
Generate exactly 10 comprehensive, highly detailed executive slides in English on "AI Governance & Maturity Assessment".
Target standards: ISO/IEC 42001 (AIMS), NIST AI RMF 1.0, and UAE National AI Ethics guidelines.
Output ONLY a valid raw JSON array containing exactly 10 objects with keys:
"title", "concept", "framework_alignment", "deep_dive_metrics" (array of 3 strings), "implementation_kpi".
"""

print("جاري إعداد محتوى العرض الأول (AI Governance)...")
ai_raw = generate_with_fallback(client, prompt_ai, supported_models)
ai_slides = parse_slides_json(ai_raw) or DEFAULT_AI_SLIDES

file_ai = f"AI_Governance_Maturity_Mastery_{today_str}.pptx"
create_deck(
    filename=file_ai,
    title_theme="AI Governance & Maturity Assessment",
    subtitle_theme="ISO/IEC 42001 & NIST AI RMF Strategic Implementation Guide",
    slides_data=ai_slides,
    primary_rgb=RGBColor(26, 54, 93),      # Navy Blue
    accent_rgb=RGBColor(197, 48, 48)       # Crimson Red
)

# 2. توليد العرض الثاني: حوكمة البيانات وداما
prompt_data = """
Generate exactly 10 comprehensive, highly detailed executive slides in English on "Data Governance & DAMA-DMBOK2 Maturity Assessment".
Target standard: DAMA-DMBOK2 and CMMI Data Management Maturity (DMM) model.
Output ONLY a valid raw JSON array containing exactly 10 objects covering DAMA Wheel areas with keys:
"title", "concept", "framework_alignment", "deep_dive_metrics" (array of 3 strings), "implementation_kpi".
"""

print("جاري إعداد محتوى العرض الثاني (Data Governance & DAMA)...")
data_raw = generate_with_fallback(client, prompt_data, supported_models)
data_slides = parse_slides_json(data_raw) or DEFAULT_DATA_SLIDES

file_data = f"Data_Governance_DAMA_Maturity_{today_str}.pptx"
create_deck(
    filename=file_data,
    title_theme="Data Governance & DAMA-DMBOK Maturity",
    subtitle_theme="Enterprise Data Management, Stewardship & DMM Measurement",
    slides_data=data_slides,
    primary_rgb=RGBColor(15, 118, 110),     # Deep Teal
    accent_rgb=RGBColor(180, 83, 9)        # Amber Bronze
)

# 3. إرسال ملفات الباوربوينت بالبريد الإلكتروني
sender_email = os.environ["GMAIL_ADDRESS"]
app_password = os.environ["GMAIL_APP_PASSWORD"]

msg = MIMEMultipart()
msg["Subject"] = f"📊 Daily Executive Masterclass: 2 PowerPoint Decks (AI Governance & DAMA Data Maturity) - {today_str}"
msg["From"] = sender_email
msg["To"] = sender_email

body = (
    "Dear Leader,\n\n"
    f"Please find attached your daily executive PowerPoint briefing decks for {today_str}.\n\n"
    "Deck 1: AI Governance & Maturity Assessment (ISO/IEC 42001 & NIST AI RMF)\n"
    "• 10 High-density slides in English covering algorithmic accountability, bias mitigation, maturity tiers, and steering governance.\n\n"
    "Deck 2: Data Governance & DAMA-DMBOK2 Maturity\n"
    "• 10 High-density slides in English covering the DAMA Wheel, Data Quality dimensions, stewardship, Metadata, and DMM levels.\n\n"
    "Both decks are formatted in 16:9 widescreen with executive dual-card layouts ready for presentation.\n\n"
    "Best regards,\n"
    "Executive Governance Assistant"
)
msg.attach(MIMEText(body, "plain", "utf-8"))

for f_name in [file_ai, file_data]:
    with open(f_name, "rb") as f:
        part = MIMEBase("application", "vnd.openxmlformats-officedocument.presentationml.presentation")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{f_name}"')
        msg.attach(part)

with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
    server.login(sender_email, app_password)
    server.sendmail(sender_email, sender_email, msg.as_string())

print("[✔] تم إرسال ملفي العرض التقديمي بنجاح تام إلى البريد الإلكتروني!")
