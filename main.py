import os
import smtplib
import json
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

# دالة الاستدعاء مع التراجع الأسي وسلسلة النماذج الاحتياطية لتفادي ضغط الخوادم
def generate_with_fallback(client, prompt):
    models = ["gemini-2.5-flash", "gemini-2.5-pro", "gemini-3.6-flash"]
    for model_name in models:
        print(f"[*] Calling Gemini via model: {model_name}...")
        for attempt in range(1, 4):
            try:
                res = client.models.generate_content(
                    model=model_name,
                    contents=prompt
                )
                if res and res.text:
                    print(f"[✔] Success with {model_name} on attempt {attempt}!")
                    return res.text
            except (errors.ServerError, errors.APIError, Exception) as e:
                print(f"[!] Warning on {model_name}: {str(e)[:70]}...")
                if attempt < 3:
                    time.sleep((2 ** attempt) + random.uniform(1.0, 2.5))
    return None

def clean_json_response(raw_text):
    text = raw_text.strip()
    if text.startswith("```json"):
        text = text[7:]
    elif text.startswith("```"):
        text = text[3:]
    if text.endswith("```"):
        text = text[:-3]
    return text.strip()

def create_deck(filename, title_theme, subtitle_theme, slides_data, primary_rgb, accent_rgb):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    blank_slide_layout = prs.slide_layouts[6]

    # 1. Title Slide
    slide_title = prs.slides.add_slide(blank_slide_layout)
    
    # Top bar banner
    banner = slide_title.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.4))
    banner.fill.solid()
    banner.fill.fore_color.rgb = primary_rgb
    banner.line.color.rgb = primary_rgb

    # Title box
    txBox = slide_title.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(3.0))
    tf = txBox.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.text = title_theme
    p.font.size = Pt(40)
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
    p3.text = f"Executive Knowledge Series | Published: {today_str}"
    p3.font.size = Pt(13)
    p3.font.name = "Arial"
    p3.font.color.rgb = RGBColor(120, 144, 156)

    # 2. Content Slides (10 Slides)
    for idx, slide_item in enumerate(slides_data, start=1):
        slide = prs.slides.add_slide(blank_slide_layout)

        # Header Ribbon
        ribbon = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
        ribbon.fill.solid()
        ribbon.fill.fore_color.rgb = primary_rgb
        ribbon.line.color.rgb = primary_rgb

        # Slide Number Badge
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

        # Slide Header Title
        title_box = slide.shapes.add_textbox(Inches(1.7), Inches(0.15), Inches(11.0), Inches(0.8))
        tp = title_box.text_frame.paragraphs[0]
        tp.text = slide_item.get("title", f"Core Principle {idx}")
        tp.font.size = Pt(22)
        tp.font.bold = True
        tp.font.name = "Arial"
        tp.font.color.rgb = RGBColor(255, 255, 255)

        # Left Card: Core Concept & Framework
        card1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(5.6), Inches(5.5))
        card1.fill.solid()
        card1.fill.fore_color.rgb = RGBColor(245, 247, 250)
        card1.line.color.rgb = RGBColor(220, 226, 235)

        tb_c1 = slide.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(5.2), Inches(5.1))
        tf_c1 = tb_c1.text_frame
        tf_c1.word_wrap = True

        h1 = tf_c1.paragraphs[0]
        h1.text = "FOUNDATIONAL CONCEPT & FRAMEWORK"
        h1.font.size = Pt(13)
        h1.font.bold = True
        h1.font.name = "Arial"
        h1.font.color.rgb = primary_rgb

        desc_p = tf_c1.add_paragraph()
        desc_p.text = slide_item.get("concept", "")
        desc_p.font.size = Pt(12)
        desc_p.font.name = "Arial"
        desc_p.font.color.rgb = RGBColor(55, 71, 79)

        align_h = tf_c1.add_paragraph()
        align_h.text = "\nSTANDARDS & FRAMEWORK ALIGNMENT:"
        align_h.font.size = Pt(12)
        align_h.font.bold = True
        align_h.font.name = "Arial"
        align_h.font.color.rgb = accent_rgb

        align_p = tf_c1.add_paragraph()
        align_p.text = slide_item.get("framework_alignment", "")
        align_p.font.size = Pt(11.5)
        align_p.font.name = "Arial"
        align_p.font.color.rgb = RGBColor(69, 90, 100)

        # Right Card: Deep Dive & Best Practices
        card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.4), Inches(5.7), Inches(5.5))
        card2.fill.solid()
        card2.fill.fore_color.rgb = RGBColor(255, 255, 255)
        card2.line.color.rgb = RGBColor(200, 210, 225)

        tb_c2 = slide.shapes.add_textbox(Inches(7.0), Inches(1.6), Inches(5.3), Inches(5.1))
        tf_c2 = tb_c2.text_frame
        tf_c2.word_wrap = True

        h2 = tf_c2.paragraphs[0]
        h2.text = "DEEP-DIVE MEASUREMENT & METRICS"
        h2.font.size = Pt(13)
        h2.font.bold = True
        h2.font.name = "Arial"
        h2.font.color.rgb = primary_rgb

        bullets = slide_item.get("deep_dive_metrics", [])
        for b in bullets:
            bp_item = tf_c2.add_paragraph()
            bp_item.text = f"• {b}"
            bp_item.font.size = Pt(11)
            bp_item.font.name = "Arial"
            bp_item.font.color.rgb = RGBColor(38, 50, 56)

        bp_h = tf_c2.add_paragraph()
        bp_h.text = "\nSTRATEGIC IMPLEMENTATION & KPI:"
        bp_h.font.size = Pt(12)
        bp_h.font.bold = True
        bp_h.font.name = "Arial"
        bp_h.font.color.rgb = accent_rgb

        bp_p = tf_c2.add_paragraph()
        bp_p.text = slide_item.get("implementation_kpi", "")
        bp_p.font.size = Pt(11)
        bp_p.font.name = "Arial"
        bp_p.font.color.rgb = RGBColor(55, 71, 79)

    prs.save(filename)
    print(f"[✔] Successfully created deck: {filename}")

# =======================
# Main Workflow Execution
# =======================
now = datetime.now()
today_str = now.strftime("%Y-%m-%d")

client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])

# 1. Generate Deck 1: AI Governance & Maturity
prompt_ai = """
Generate exactly 10 comprehensive, highly detailed executive slides in English on "AI Governance & Maturity Assessment".
Target standards: ISO/IEC 42001 (Artificial Intelligence Management System - AIMS), NIST AI Risk Management Framework (NIST AI RMF 1.0), and UAE National AI Ethics guidelines.

Provide a valid JSON array of 10 objects, where each object has:
- "title": (Concise, professional title of the principle/dimension)
- "concept": (In-depth paragraph explaining the core governance mechanism, maturity measurement, and organizational scope)
- "framework_alignment": (Specific ISO 42001 clause, NIST AI RMF function - GOVERN, MAP, MEASURE, MANAGE, or maturity tier)
- "deep_dive_metrics": (Array of 3 detailed bullet points covering audits, fairness/bias thresholds, prompt safety, algorithmic explainability, or risk classification)
- "implementation_kpi": (Tangible KPI or executive action item to measure maturity)

Ensure content is rigorous, highly informative, and enterprise-grade. Output ONLY the raw JSON array.
"""

print("Generating content for Deck 1 (AI Governance)...")
ai_content_raw = generate_with_fallback(client, prompt_ai)
ai_data = json.loads(clean_json_response(ai_content_raw))

file_ai = f"AI_Governance_Maturity_Mastery_{today_str}.pptx"
create_deck(
    filename=file_ai,
    title_theme="AI Governance & Maturity Assessment",
    subtitle_theme="ISO/IEC 42001 & NIST AI RMF Implementation Guide",
    slides_data=ai_data,
    primary_rgb=RGBColor(26, 54, 93),      # Navy Blue
    accent_rgb=RGBColor(197, 48, 48)       # Crimson Red
)

# 2. Generate Deck 2: Data Governance & DAMA-DMBOK Maturity
prompt_data = """
Generate exactly 10 comprehensive, highly detailed executive slides in English on "Data Governance & DAMA-DMBOK2 Maturity Assessment".
Target standard: DAMA-DMBOK2 (Data Management Body of Knowledge) and CMMI Data Management Maturity (DMM) model.

Provide a valid JSON array of 10 objects covering distinct DAMA Wheel knowledge areas (e.g. Data Governance, Data Quality, Metadata, Master Data / MDM, Data Architecture, Data Security, Data Lineage, Maturity Scoring Levels 1 to 5):
- "title": (Professional dimension title)
- "concept": (In-depth paragraph explaining the management principle, stewardship model, and governance lifecycle)
- "framework_alignment": (Specific DAMA-DMBOK2 knowledge area and DMM Maturity Level: Initial, Managed, Defined, Measured, Optimized)
- "deep_dive_metrics": (Array of 3 detailed bullet points covering data quality dimensions, cataloging standards, policy enforcement, or classification)
- "implementation_kpi": (Tangible KPI or executive best practice to advance maturity)

Ensure content is rich in information and enterprise-grade. Output ONLY the raw JSON array.
"""

print("Generating content for Deck 2 (Data Governance & DAMA)...")
data_content_raw = generate_with_fallback(client, prompt_data)
data_data = json.loads(clean_json_response(data_content_raw))

file_data = f"Data_Governance_DAMA_Maturity_{today_str}.pptx"
create_deck(
    filename=file_data,
    title_theme="Data Governance & DAMA-DMBOK Maturity",
    subtitle_theme="Enterprise Data Management, Stewardship & DMM Measurement",
    slides_data=data_data,
    primary_rgb=RGBColor(15, 118, 110),     # Deep Teal
    accent_rgb=RGBColor(180, 83, 9)        # Amber Bronze
)

# 3. Send Email with Both Presentations
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

print("[✔] Both PowerPoint decks sent successfully!")
