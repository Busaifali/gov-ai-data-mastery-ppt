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

# جلب النماذج المتاحة لحسابك تلقائياً
def get_supported_models(client):
    try:
        discovered = []
        for m in client.models.list():
            m_name = getattr(m, "name", "")
            clean_name = m_name.replace("models/", "")
            if any(k in clean_name for k in ["flash", "pro"]) and not any(x in clean_name for x in ["embed", "imagen", "vision", "realtime"]):
                discovered.append(clean_name)
        if discovered:
            print(f"[*] النماذج المكتشفة لحسابك: {discovered[:5]}")
            return discovered
    except Exception as e:
        print(f"[!] تعذر جلب قائمة النماذج تلقائياً ({e}).")
    return ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]

# استدعاء Gemini مع التراجع الأسي وسلسلة النماذج
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
                    break
                if attempt < 3:
                    time.sleep((2 ** attempt) + random.uniform(1.0, 2.5))
    return None

# استخراج كائن JSON من رد النموذج
def parse_slides_json(raw_text):
    if not raw_text:
        return None
    try:
        match = re.search(r'\[\s*\{.*\}\s*\]', raw_text, re.DOTALL)
        if match:
            return json.loads(match.group(0))
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

# دالة إنشاء العرض التقديمي (تدعم الإنجليزية والعربية RTL/LTR)
def create_deck(filename, title_theme, subtitle_theme, slides_data, primary_rgb, accent_rgb):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # 1. شريحة الغلاف
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
    p.font.size = Pt(36)
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
    p3.text = f"Executive Governance Series (20 Slides: 10 EN + 10 AR) | Published: {today_str}"
    p3.font.size = Pt(13)
    p3.font.name = "Arial"
    p3.font.color.rgb = RGBColor(120, 144, 156)

    # 2. توليد الشرائح الـ 20
    for idx, item in enumerate(slides_data, start=1):
        slide = prs.slides.add_slide(blank_layout)
        is_arabic = (item.get("lang", "en") == "ar") or any('\u0600' <= c <= '\u06FF' for c in item.get("title", ""))
        align_choice = PP_ALIGN.RIGHT if is_arabic else PP_ALIGN.LEFT

        # الشريط العلوي
        ribbon = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
        ribbon.fill.solid()
        ribbon.fill.fore_color.rgb = primary_rgb
        ribbon.line.color.rgb = primary_rgb

        # رقم الشريحة
        badge_left = Inches(11.8) if is_arabic else Inches(0.8)
        badge = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, badge_left, Inches(0.2), Inches(0.7), Inches(0.7))
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
        title_left = Inches(0.8) if is_arabic else Inches(1.7)
        title_box = slide.shapes.add_textbox(title_left, Inches(0.15), Inches(10.8), Inches(0.8))
        tp = title_box.text_frame.paragraphs[0]
        tp.text = item.get("title", f"Slide {idx}")
        tp.font.size = Pt(21)
        tp.font.bold = True
        tp.font.name = "Arial"
        tp.font.color.rgb = RGBColor(255, 255, 255)
        tp.alignment = align_choice

        # البطاقة الأولى: المفهوم والإطار
        c1_left = Inches(6.9) if is_arabic else Inches(0.8)
        card1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, c1_left, Inches(1.4), Inches(5.6), Inches(5.5))
        card1.fill.solid()
        card1.fill.fore_color.rgb = RGBColor(245, 247, 250)
        card1.line.color.rgb = RGBColor(220, 226, 235)

        tb_c1 = slide.shapes.add_textbox(c1_left + Inches(0.2), Inches(1.6), Inches(5.2), Inches(5.1))
        tf_c1 = tb_c1.text_frame
        tf_c1.word_wrap = True

        h1 = tf_c1.paragraphs[0]
        h1.text = "المبدأ الأساسي والإطار التنظيمي" if is_arabic else "CORE PRINCIPLE & FRAMEWORK"
        h1.font.size = Pt(13)
        h1.font.bold = True
        h1.font.name = "Arial"
        h1.font.color.rgb = primary_rgb
        h1.alignment = align_choice

        desc_p = tf_c1.add_paragraph()
        desc_p.text = item.get("concept", "")
        desc_p.font.size = Pt(11.5)
        desc_p.font.name = "Arial"
        desc_p.font.color.rgb = RGBColor(55, 71, 79)
        desc_p.alignment = align_choice

        align_h = tf_c1.add_paragraph()
        align_h.text = ("\nالمعيار ونموذج النضج:" if is_arabic else "\nSTANDARDS & FRAMEWORK ALIGNMENT:")
        align_h.font.size = Pt(12)
        align_h.font.bold = True
        align_h.font.name = "Arial"
        align_h.font.color.rgb = accent_rgb
        align_h.alignment = align_choice

        align_p = tf_c1.add_paragraph()
        align_p.text = item.get("framework_alignment", "")
        align_p.font.size = Pt(11)
        align_p.font.name = "Arial"
        align_p.font.color.rgb = RGBColor(69, 90, 100)
        align_p.alignment = align_choice

        # البطاقة الثانية: القياس والتطبيق
        c2_left = Inches(0.8) if is_arabic else Inches(6.8)
        card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, c2_left, Inches(1.4), Inches(5.7), Inches(5.5))
        card2.fill.solid()
        card2.fill.fore_color.rgb = RGBColor(255, 255, 255)
        card2.line.color.rgb = RGBColor(200, 210, 225)

        tb_c2 = slide.shapes.add_textbox(c2_left + Inches(0.2), Inches(1.6), Inches(5.3), Inches(5.1))
        tf_c2 = tb_c2.text_frame
        tf_c2.word_wrap = True

        h2 = tf_c2.paragraphs[0]
        h2.text = "مؤشرات القياس والتدقيق المتعمق" if is_arabic else "DEEP-DIVE MEASUREMENT & METRICS"
        h2.font.size = Pt(13)
        h2.font.bold = True
        h2.font.name = "Arial"
        h2.font.color.rgb = primary_rgb
        h2.alignment = align_choice

        bullets = item.get("deep_dive_metrics", [])
        for b in bullets:
            bp_item = tf_c2.add_paragraph()
            bp_item.text = f"• {b}"
            bp_item.font.size = Pt(11)
            bp_item.font.name = "Arial"
            bp_item.font.color.rgb = RGBColor(38, 50, 56)
            bp_item.alignment = align_choice

        bp_h = tf_c2.add_paragraph()
        bp_h.text = ("\nالمؤشر وممارسة التطبيق التنفيذية:" if is_arabic else "\nSTRATEGIC ACTION / KPI:")
        bp_h.font.size = Pt(12)
        bp_h.font.bold = True
        bp_h.font.name = "Arial"
        bp_h.font.color.rgb = accent_rgb
        bp_h.alignment = align_choice

        bp_p = tf_c2.add_paragraph()
        bp_p.text = item.get("implementation_kpi", "")
        bp_p.font.size = Pt(11)
        bp_p.font.name = "Arial"
        bp_p.font.color.rgb = RGBColor(55, 71, 79)
        bp_p.alignment = align_choice

    prs.save(filename)
    print(f"[✔] تم بنجاح إنشاء العرض التقديمي: {filename} ({len(slides_data)} شريحة)")

# ===============================
# التنفيذ الرئيسي
# ===============================
now = datetime.now()
today_str = now.strftime("%Y-%m-%d")

client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])
supported_models = get_supported_models(client)

# 1. توليد 20 شريحة لملف حوكمة ونضج الذكاء الاصطناعي (10 إنجليزي + 10 عربي)
prompt_ai = """
Generate exactly 20 comprehensive, highly detailed executive slides on "AI Governance & Maturity Assessment" (ISO/IEC 42001, NIST AI RMF 1.0, UAE AI Ethics).
STRUCTURE REQUIREMENT:
- Slides 1 to 10: In ENGLISH ("lang": "en") covering 10 distinct advanced dimensions.
- Slides 11 to 20: In ARABIC ("lang": "ar") with rich, professional enterprise terminology covering 10 complementary strategic dimensions.

Output ONLY a valid raw JSON array containing exactly 20 objects with keys:
- "lang": ("en" for the first 10, "ar" for the remaining 10)
- "title": (Professional slide title)
- "concept": (In-depth explanatory paragraph)
- "framework_alignment": (Specific clause/tier alignment)
- "deep_dive_metrics": (Array of 3 detailed bullet points)
- "implementation_kpi": (Tangible KPI or executive action item)
"""

print("جاري إعداد محتوى العرض الأول: حوكمة الذكاء الاصطناعي (20 شريحة)...")
ai_raw = generate_with_fallback(client, prompt_ai, supported_models)
ai_slides = parse_slides_json(ai_raw)

if not ai_slides or len(ai_slides) < 10:
    print("[!] حدث تعثر في جلب الـ 20 شريحة كاملة، سيتم إعادة المحاولة بنموذج مستقر...")
    ai_raw = generate_with_fallback(client, prompt_ai, ["gemini-2.5-flash", "gemini-2.5-pro"])
    ai_slides = parse_slides_json(ai_raw)

file_ai = f"AI_Governance_Maturity_20Slides_{today_str}.pptx"
create_deck(
    filename=file_ai,
    title_theme="AI Governance & Maturity Assessment",
    subtitle_theme="ISO/IEC 42001 & NIST AI RMF (10 English + 10 Arabic Slides)",
    slides_data=ai_slides,
    primary_rgb=RGBColor(26, 54, 93),      # Navy Blue
    accent_rgb=RGBColor(197, 48, 48)       # Crimson Red
)

# 2. توليد 20 شريحة لملف حوكمة ونضج البيانات ومعيار داما (10 إنجليزي + 10 عربي)
prompt_data = """
Generate exactly 20 comprehensive, highly detailed executive slides on "Data Governance & DAMA-DMBOK2 Maturity Assessment" (DAMA Wheel, CMMI DMM).
STRUCTURE REQUIREMENT:
- Slides 1 to 10: In ENGLISH ("lang": "en") covering DAMA core knowledge areas.
- Slides 11 to 20: In ARABIC ("lang": "ar") with rich, professional enterprise terminology covering advanced data maturity dimensions.

Output ONLY a valid raw JSON array containing exactly 20 objects with keys:
- "lang": ("en" for the first 10, "ar" for the remaining 10)
- "title": (Professional slide title)
- "concept": (In-depth explanatory paragraph)
- "framework_alignment": (Specific DAMA / DMM alignment)
- "deep_dive_metrics": (Array of 3 detailed bullet points)
- "implementation_kpi": (Tangible KPI or executive action item)
"""

print("جاري إعداد محتوى العرض الثاني: حوكمة البيانات وداما (20 شريحة)...")
data_raw = generate_with_fallback(client, prompt_data, supported_models)
data_slides = parse_slides_json(data_raw)

if not data_slides or len(data_slides) < 10:
    data_raw = generate_with_fallback(client, prompt_data, ["gemini-2.5-flash", "gemini-2.5-pro"])
    data_slides = parse_slides_json(data_raw)

file_data = f"Data_Governance_DAMA_20Slides_{today_str}.pptx"
create_deck(
    filename=file_data,
    title_theme="Data Governance & DAMA-DMBOK Maturity",
    subtitle_theme="Enterprise Data Management & DMM Measurement (10 English + 10 Arabic Slides)",
    slides_data=data_slides,
    primary_rgb=RGBColor(15, 118, 110),     # Deep Teal
    accent_rgb=RGBColor(180, 83, 9)        # Amber Bronze
)

# 3. إرسال ملفات الباوربوينت بالبريد الإلكتروني
sender_email = os.environ["GMAIL_ADDRESS"]
app_password = os.environ["GMAIL_APP_PASSWORD"]

msg = MIMEMultipart()
msg["Subject"] = f"📊 العروض التقديمية اليومية الموسعة: 2 ملف باوربوينت (20 شريحة لكل ملف: عربي وإنجليزي) - {today_str}"
msg["From"] = sender_email
msg["To"] = sender_email

body = (
    "السلام عليكم ورحمة الله وبركاته،\n\n"
    f"مرفق لسيادتكم العروض التقديمية التنفيذية الموسعة ليوم {today_str}.\n\n"
    "الملف الأول: حوكمة ونضج الذكاء الاصطناعي (AI Governance & ISO 42001)\n"
    "• يحتوي على 20 شريحة غنية بالمعلومات (10 شرائح بالإنجليزية + 10 شرائح بالعربية الفصحى).\n\n"
    "الملف الثاني: نضج حوكمة البيانات ومعيار داما (Data Governance & DAMA-DMBOK2)\n"
    "• يحتوي على 20 شريحة غنية بالمعلومات (10 شرائح بالإنجليزية + 10 شرائح بالعربية الفصحى).\n\n"
    "كلا الملفين منسقان بنظام العرض العريض 16:9 مع بطاقات مزدوجة واحترافية جاهزة للعرض والتدريب المباشر.\n\n"
    "تحياتنا،\nالمساعد التنفيذي للحوكمة"
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

print("[✔] تم إرسال ملفي العرض الموسعين (20 شريحة لكل ملف) بنجاح تام إلى بريدك الإلكتروني!")
