#!/usr/bin/env python3
"""Generate a professional Arabic RTL Word CV for Mahmoud Hamed."""

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsmap
from docx.shared import Cm, Pt, RGBColor


ACCENT = RGBColor(0x0B, 0x4F, 0x6C)
INK = RGBColor(0x1A, 0x23, 0x32)
MUTED = RGBColor(0x5C, 0x6B, 0x7A)


def set_run_font(run, size=11, bold=False, color=INK, name="Arial"):
    run.bold = bold
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.name = name
    r = run._element
    rpr = r.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:ascii"), name)
    rfonts.set(qn("w:hAnsi"), name)
    rfonts.set(qn("w:cs"), name)
    rfonts.set(qn("w:eastAsia"), name)


def set_paragraph_rtl(paragraph, space_after=6, space_before=0, align=WD_ALIGN_PARAGRAPH.RIGHT):
    paragraph.alignment = align
    pf = paragraph.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE

    ppr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1")
    ppr.append(bidi)


def add_horizontal_line(paragraph):
    ppr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "18")
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), "0B4F6C")
    pBdr.append(bottom)
    ppr.append(pBdr)


def add_heading_ar(doc, text):
    p = doc.add_paragraph()
    set_paragraph_rtl(p, space_after=8, space_before=14)
    run = p.add_run(text)
    set_run_font(run, size=14, bold=True, color=ACCENT)
    add_horizontal_line(p)
    return p


def add_text(doc, text, size=11, bold=False, color=INK, space_after=6, space_before=0, align=WD_ALIGN_PARAGRAPH.RIGHT):
    p = doc.add_paragraph()
    set_paragraph_rtl(p, space_after=space_after, space_before=space_before, align=align)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    return p


def add_bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    set_paragraph_rtl(p, space_after=4)
    # Clear default run text if any and set Arabic content
    if p.runs:
        p.runs[0].text = text
        set_run_font(p.runs[0], size=11, color=INK)
    else:
        run = p.add_run(text)
        set_run_font(run, size=11, color=INK)
    return p


def set_document_rtl(doc):
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(11)
    normal._element.rPr.rFonts.set(qn("w:cs"), "Arial")

    # Set document default RTL via sectPr / settings bidi
    settings = doc.settings.element
    # Ensure themeFontLang includes Arabic if present; add bidi-friendly view
    # Also set section RTL for page
    section = doc.sections[0]
    sectPr = section._sectPr
    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1")
    sectPr.append(bidi)


def build_cv(output_path: str):
    doc = Document()
    set_document_rtl(doc)

    section = doc.sections[0]
    section.top_margin = Cm(1.6)
    section.bottom_margin = Cm(1.6)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)

    # Header
    name = add_text(
        doc,
        "محمود محمد عبدالفتاح حامد",
        size=22,
        bold=True,
        color=ACCENT,
        space_after=4,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    add_text(
        doc,
        "منسّق عمليات التوصيل (Dispatcher) | طالب علوم حاسب — متوقع التخرج 2027",
        size=12,
        bold=True,
        color=INK,
        space_after=6,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    contact = add_text(
        doc,
        "الهاتف: 01102600149   |   البريد: hhuumahmoud@gmail.com   |   القاهرة، مصر",
        size=10,
        color=MUTED,
        space_after=4,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    # underline under contact
    add_horizontal_line(contact)

    # Professional summary
    add_heading_ar(doc, "الملف المهني")
    add_text(
        doc,
        "منسّق عمليات توصيل بخبرة عملية في بيئات التجارة السريعة والتوصيل عند الطلب، مع قدرة واضحة على إدارة الطلبات في الوقت الفعلي، والتنسيق بين فرق التوصيل والدعم والعملاء، والمتابعة الدقيقة لمؤشرات الأداء. أجمع بين مهارات التنظيم والعمل تحت الضغط وبين خلفية أكاديمية في علوم الحاسب، وأسعى للانضمام إلى مؤسسة كبرى تتيح فرص تطور مهني في مجالات العمليات، خدمة العملاء، أو دعم الأنظمة التشغيلية.",
        size=11,
        space_after=8,
    )

    # Experience
    add_heading_ar(doc, "الخبرة الوظيفية")

    add_text(doc, "منسّق عمليات التوصيل (Dispatcher) — Breadfast", size=12, bold=True, space_after=2)
    add_text(doc, "يناير 2026 — يونيو 2026", size=10, color=MUTED, space_after=4)
    add_text(
        doc,
        "العمل ضمن فريق العمليات لضمان تدفق الطلبات بسلاسة عبر تطبيق التوصيل السريع، مع التركيز على سرعة الاستجابة وجودة التنفيذ ودقة المتابعة.",
        size=11,
        space_after=6,
    )
    for item in [
        "استقبال الطلبات الواردة وتوزيعها على مندوبي التوصيل وفق المناطق والأولوية وكفاءة التغطية المتاحة.",
        "متابعة حالة الطلبات لحظة بلحظة والتدخل السريع عند التأخير أو إعادة التوجيه أو حالات الذروة.",
        "التنسيق المستمر مع فرق الدعم والتشغيل لمعالجة المشكلات التشغيلية وتقليل زمن الاستجابة.",
        "متابعة مؤشرات الأداء الأساسية مثل زمن التوصيل، نسبة الإنجاز، وكفاءة توزيع الطلبات.",
        "التعامل الاحترافي مع الملاحظات الميدانية ورفع التقارير اليومية لدعم اتخاذ القرار التشغيلي.",
    ]:
        add_bullet(doc, item)

    add_text(doc, "منسّق عمليات التوصيل (Dispatcher) — InstaShop", size=12, bold=True, space_before=10, space_after=2)
    add_text(doc, "2024 — 2025", size=10, color=MUTED, space_after=4)
    add_text(
        doc,
        "خبرة لمدة عام في بيئة توصيل سريعة الإيقاع، مع مسؤولية مباشرة عن تنظيم حركة المندوبين وضمان التزام الطلبات بالجداول الزمنية المتفق عليها.",
        size=11,
        space_after=6,
    )
    for item in [
        "إدارة وتنظيم جداول التوصيل اليومية وتوزيع المهام بما يحقق أفضل استفادة من الموارد المتاحة.",
        "متابعة أداء المناديب ميدانيًا ودعمهم في حل المشكلات اللوجستية أثناء تنفيذ الطلبات.",
        "التواصل الفعّال مع العملاء عند الحاجة لتوضيح حالة الطلب أو ترتيب حلول بديلة عند الطوارئ.",
        "المساهمة في تحسين انسيابية العمليات اليومية من خلال الالتزام بالإجراءات ومعايير الجودة.",
        "العمل ضمن فريق متعدد المهام مع الحفاظ على الدقة والهدوء خلال فترات الضغط العالي.",
    ]:
        add_bullet(doc, item)

    # Education
    add_heading_ar(doc, "التعليم")
    add_text(doc, "بكالوريوس علوم الحاسب", size=12, bold=True, space_after=2)
    add_text(doc, "المعهد العالي بالتجمع الخامس", size=11, bold=True, color=ACCENT, space_after=2)
    add_text(doc, "السنة الدراسية الحالية: السنة الثالثة  |  متوقع التخرج: 2027", size=10, color=MUTED, space_after=6)
    for item in [
        "دراسة تخصصات أساسية في علوم الحاسب تشمل البرمجة، قواعد البيانات، الشبكات، ونظم المعلومات.",
        "تنمية مهارات التفكير التحليلي وحل المشكلات من خلال المقررات العملية والمشروعات الدراسية.",
        "الجمع بين الدراسة الأكاديمية والخبرة العملية في مجال العمليات والتوصيل لتعزيز الجاهزية لسوق العمل.",
    ]:
        add_bullet(doc, item)

    # Skills
    add_heading_ar(doc, "المهارات المهنية")
    for item in [
        "إدارة عمليات التوصيل",
        "توزيع الطلبات والمتابعة اللحظية",
        "خدمة العملاء والتنسيق بين الفرق",
        "العمل تحت الضغط وحل المشكلات التشغيلية",
        "إعداد التقارير اليومية ومتابعة مؤشرات الأداء",
    ]:
        add_bullet(doc, item)

    add_heading_ar(doc, "المهارات التقنية")
    for item in [
        "أنظمة إدارة الطلبات وتطبيقات التوصيل",
        "Microsoft Office و Excel",
        "أساسيات علوم الحاسب والتعامل مع البيانات",
    ]:
        add_bullet(doc, item)

    add_heading_ar(doc, "الكفاءات الشخصية")
    for item in [
        "انضباط عالٍ والتزام بالمواعيد",
        "تواصل واضح وفعّال",
        "قدرة على اتخاذ قرارات سريعة",
        "روح تعاون ومرونة في العمل",
        "دقة في المتابعة والتوثيق",
        "قابلية تعلم سريعة",
    ]:
        add_bullet(doc, item)

    add_heading_ar(doc, "اللغات")
    add_bullet(doc, "العربية — لغة أم")
    add_bullet(doc, "الإنجليزية — مستوى متوسط")

    add_heading_ar(doc, "بيانات شخصية")
    add_bullet(doc, "الجنسية: مصري")
    add_bullet(doc, "الموقع: القاهرة، مصر")
    add_bullet(doc, "التفرّغ: متاح للعمل")

    doc.save(output_path)
    print(f"Saved: {output_path}")


if __name__ == "__main__":
    build_cv("/workspace/CV-محمود-محمد-عبدالفتاح-حامد.docx")
