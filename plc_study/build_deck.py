"""Builds the PLC study presentation (Oracle Communications Launch Cloud Service).

Run:  python3 build_deck.py   ->  PLC_Oracle_Launch_Study.pptx
"""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

PURPLE = RGBColor(0x5B, 0x2C, 0x83)
PURPLE_LIGHT = RGBColor(0xEE, 0xE7, 0xF5)
DARK = RGBColor(0x1F, 0x1B, 0x2E)
GREY = RGBColor(0x5F, 0x5A, 0x6B)
LIGHT = RGBColor(0xF4, 0xF3, 0xF7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RED = RGBColor(0xC7, 0x46, 0x34)
TEAL = RGBColor(0x00, 0x8C, 0x95)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
AMBER = RGBColor(0xB2, 0x6A, 0x00)

FONT = "Calibri"
SYSTEM = "Oracle Communications Launch Cloud Service"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


def _style(run, size, color=DARK, bold=False, italic=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic


def text(slide, x, y, w, h, value, size=16, color=DARK, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.05)
    lines = value if isinstance(value, list) else [value]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        _style(run, size, color, bold, italic)
    return box


def bullets(slide, x, y, w, h, items, size=16, color=DARK, space=6):
    """items: str | (str, level) | (str, level, bold_prefix)"""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        if isinstance(item, str):
            item = (item, 0)
        label, level = item[0], item[1]
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(space)
        marker = "\u25A0  " if level == 0 else "\u2013  "
        indent = "      " * level
        r = p.add_run()
        r.text = indent + marker
        _style(r, size - 4 if level == 0 else size - 2, PURPLE if level == 0 else GREY)
        if ": " in label and level == 0:
            head, tail = label.split(": ", 1)
            r1 = p.add_run()
            r1.text = head + ": "
            _style(r1, size, color, bold=True)
            r2 = p.add_run()
            r2.text = tail
            _style(r2, size, color)
        else:
            r1 = p.add_run()
            r1.text = label
            _style(r1, size if level == 0 else size - 2, color if level == 0 else GREY)
    return box


def box(slide, x, y, w, h, value, fill=PURPLE, color=WHITE, size=14, bold=True,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=None, align=PP_ALIGN.CENTER):
    shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(1.25)
    shp.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        shp.adjustments[0] = 0.12
    tf = shp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = Inches(0.08)
    tf.margin_top = tf.margin_bottom = Inches(0.04)
    lines = value if isinstance(value, list) else [value]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = line
        _style(r, size if i == 0 else size - 2, color, bold if i == 0 else False)
    return shp


def arrow(slide, x, y, w, h=0.32, fill=GREY):
    shp = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def down_arrow(slide, x, y, w=0.3, h=0.35, fill=GREY):
    shp = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def table(slide, x, y, w, rows, col_widths, size=12, row_h=0.36, header_fill=PURPLE):
    shape = slide.shapes.add_table(len(rows), len(rows[0]), Inches(x), Inches(y),
                                   Inches(w), Inches(row_h * len(rows)))
    tbl = shape.table
    for c, cw in enumerate(col_widths):
        tbl.columns[c].width = Inches(cw)
    for r, row in enumerate(rows):
        tbl.rows[r].height = Inches(row_h)
        for c, value in enumerate(row):
            cell = tbl.cell(r, c)
            cell.margin_left = cell.margin_right = Inches(0.06)
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if r == 0:
                cell.fill.fore_color.rgb = header_fill
            else:
                cell.fill.fore_color.rgb = LIGHT if r % 2 else WHITE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = value
            _style(run, size, WHITE if r == 0 else DARK, bold=(r == 0 or c == 0))
    return tbl


def header(slide, section, title, subtitle=None):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.18), prs.slide_height)
    bar.fill.solid()
    bar.fill.fore_color.rgb = PURPLE
    bar.line.fill.background()
    if section:
        text(slide, 0.55, 0.3, 9, 0.35, section.upper(), size=12, color=PURPLE, bold=True)
    text(slide, 0.55, 0.58, 12.2, 0.7, title, size=30, color=DARK, bold=True)
    if subtitle:
        text(slide, 0.55, 1.22, 12.2, 0.45, subtitle, size=15, color=GREY, italic=True)


def footer(slide, n):
    text(slide, 0.55, 7.05, 9, 0.3, f"PLC Study  |  {SYSTEM}", size=10, color=GREY)
    text(slide, 12.0, 7.05, 0.8, 0.3, str(n), size=10, color=GREY, align=PP_ALIGN.RIGHT)


def notes(slide, value):
    slide.notes_slide.notes_text_frame.text = value


def new_slide(section, title, subtitle=None):
    slide = prs.slides.add_slide(BLANK)
    header(slide, section, title, subtitle)
    footer(slide, len(prs.slides))
    return slide


# ---------------------------------------------------------------- 1. Title
s = prs.slides.add_slide(BLANK)
bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
bg.fill.solid()
bg.fill.fore_color.rgb = PURPLE
bg.line.fill.background()
text(s, 0.9, 1.0, 11, 0.5, "PRODUCT LIFECYCLE (PLC) SYSTEMS STUDY", size=16, color=PURPLE_LIGHT, bold=True)
text(s, 0.9, 1.7, 11.5, 1.6, ["Oracle Communications", "Launch Cloud Service"], size=48, color=WHITE, bold=True)
text(s, 0.9, 3.75, 11.5, 0.9,
     "How a telecom-grade system takes a product from idea to launch, change, retirement and "
     "customer migration, and what we should build in-house",
     size=20, color=PURPLE_LIGHT)
line = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.9), Inches(5.0), Inches(1.6), Inches(0.06))
line.fill.solid()
line.fill.fore_color.rgb = WHITE
line.line.fill.background()
text(s, 0.9, 5.25, 11, 0.9, ["Telecom Egypt  |  PLC Study Team", "Presenter: [Your name]   |   ~15 minutes"],
     size=16, color=WHITE)
notes(s, """(0:00 – 0:30)
السلام عليكم. النهارده هعرض عليكم نظام PLC حقيقي موجود في السوق ومعمول مخصوص لشركات الاتصالات، وهو Oracle Communications Launch Cloud Service.
الهدف مش إني أشرح يعني إيه PLC نظريًا، الهدف إننا نشوف نظام حقيقي بيدير رحلة المنتج إزاي، ونطلع في الآخر بإجابة على سؤال: لو هنبني PLC عندنا في المصرية للاتصالات، إيه الـ Modules والـ Features اللي لازم تكون فيه.""")

# ---------------------------------------------------------------- 2. Agenda
s = new_slide("Agenda", "What we will cover in 15 minutes")
items = [
    ("01", "System Overview", "Vendor, positioning, why this system"),
    ("02", "Architecture & Modules", "Hub-and-spoke design, work areas, data model"),
    ("03", "Product Lifecycle", "Idea  \u2192  Launch  \u2192  Revise  \u2192  Retire"),
    ("04", "Key Features", "Workflow, approvals, pricing, versioning, permissions"),
    ("05", "CRM Integration", "How offers are published to CRM, billing, provisioning"),
    ("06", "Telecom Use Case", "Launching a home fiber bundle and retiring an old package"),
    ("07", "Conclusion", "What we should build in our own PLC"),
]
for i, (num, title, sub) in enumerate(items):
    col, row = divmod(i, 4)
    x, y = 0.6 + col * 6.3, 1.65 + row * 1.3
    box(s, x, y, 0.95, 1.0, num, fill=PURPLE, size=24)
    text(s, x + 1.1, y + 0.08, 4.9, 0.45, title, size=20, bold=True)
    text(s, x + 1.1, y + 0.52, 4.9, 0.45, sub, size=14, color=GREY)
box(s, 6.9, 5.55, 5.9, 1.0,
    ["Why this system?", "Telecom-specific, TM Forum-aligned, and fully documented in public, "
     "so every claim here can be checked"],
    fill=PURPLE_LIGHT, color=DARK, size=15)
notes(s, """(0:30 – 1:00)
العرض متقسم على السبع نقط اللي اتفقنا عليهم في الاجتماع: نظرة عامة، الـ Architecture والـ Modules، رحلة المنتج، الـ Features الأساسية، الربط مع الـ CRM، مثال عملي من شغل الاتصالات، وفي الآخر الخلاصة: نبني إيه عندنا.
اخترت Oracle Launch لتلات أسباب: معمول مخصوص لشركات الاتصالات، ماشي على معايير TM Forum اللي هي لغة الصناعة، والـ Documentation بتاعته منشورة بالكامل، فكل معلومة في العرض ده ليها مصدر نقدر نرجعله.""")

# ---------------------------------------------------------------- 3. Overview
s = new_slide("01  System Overview", SYSTEM,
              "A cloud (SaaS) catalog and offer-design hub built for communications service providers")
facts = [
    ["Item", "Details"],
    ["Vendor", "Oracle, as part of the Oracle Communications portfolio (runs on Oracle Fusion Cloud)"],
    ["Category", "Product catalog & offer lifecycle management (design-time catalog hub)"],
    ["Main users", "Marketing, product managers, product specialists, catalog admins"],
    ["Core job", "Design offers once, govern them through a lifecycle, then publish them to CRM, "
                 "billing and provisioning"],
    ["Standards", "REST APIs modelled on TM Forum TMF620 (Product Catalog Management)"],
    ["Release model", "Quarterly SaaS updates (e.g. 26A, 26B, 26C), including AI offer validation"],
]
table(s, 0.6, 1.85, 7.6, facts, [1.7, 5.9], size=13, row_h=0.62)
box(s, 8.6, 1.85, 4.2, 0.5, "Problems it solves", fill=PURPLE, size=16)
bullets(s, 8.6, 2.45, 4.25, 4.2, [
    "Slow time-to-market for new offers",
    "The same product defined differently in CRM, billing and network",
    "Approvals by email with no audit trail",
    "Old offers never retired, so the catalog keeps growing",
    "No controlled way to move customers from old plans to new ones",
], size=15)
notes(s, """(1:00 – 2:30)
Oracle Launch هو نظام SaaS على Oracle Fusion Cloud، وظيفته الأساسية إنه يبقى "المكان الوحيد" اللي بيتصمم فيه المنتج أو العرض (Offer).
الفكرة إن المنتج بيتعرف مرة واحدة بس هنا، بيعدي على دورة حياة فيها موافقات واختبار، وبعدين النظام نفسه بيبعته للأنظمة التانية: الـ CRM علشان البيع، والـ Billing علشان التسعير والتحصيل، والـ Provisioning علشان التفعيل على الشبكة.
المستخدمين الأساسيين هم الماركتنج والـ Product Managers.
الـ APIs بتاعته مبنية على معيار TMF620 من TM Forum، وده معيار كتالوج المنتجات في صناعة الاتصالات كلها.
المشاكل اللي بيحلها هي غالبًا نفس المشاكل اللي بنقابلها: المنتج متعرف بشكل مختلف في كل نظام، الموافقات بالإيميل، والعروض القديمة مش بتتقفل.""")

# ---------------------------------------------------------------- 4. Architecture
s = new_slide("02  Architecture", "Hub-and-spoke: design once, publish everywhere",
              "Launch is the design-time master; downstream systems consume what it publishes")
box(s, 0.6, 1.9, 2.6, 0.45, "Who / what feeds it", fill=GREY, size=14)
for i, label in enumerate(["Marketing & product teams (web UI, guided flows)",
                           "REST APIs (TMF620-style)",
                           "One-time import of the existing catalog (e.g. from CRM)"]):
    box(s, 0.6, 2.5 + i * 1.15, 2.6, 0.95, label, fill=LIGHT, color=DARK, size=13, bold=False, line=GREY)
arrow(s, 3.35, 3.6, 0.6)
hub = box(s, 4.1, 1.9, 4.3, 4.6, "", fill=PURPLE_LIGHT, line=PURPLE)
text(s, 4.2, 1.98, 4.1, 0.4, "LAUNCH CLOUD SERVICE", size=15, color=PURPLE, bold=True, align=PP_ALIGN.CENTER)
mods = ["Initiatives (projects & approvals)", "Offer & Promotion Management", "Workbench (product/service specs)",
        "Pricing & Price Lists", "Rules (eligibility, compatibility, migration)", "Lifecycle & Publishing Admin",
        "Reports / Dashboards"]
for i, m in enumerate(mods):
    box(s, 4.3, 2.45 + i * 0.57, 3.9, 0.48, m, fill=WHITE, color=DARK, size=12, bold=False, line=PURPLE)
arrow(s, 8.55, 3.6, 0.6)
box(s, 9.25, 1.9, 0.95, 4.6, ["Integration", "fabric"], fill=TEAL, size=10)
spokes = [("10", "CRM", "e.g. Siebel CRM (prebuilt)"), ("20", "Commerce / digital channels", "web, app, shops"),
          ("30", "Billing & charging", "e.g. Oracle PDC / BRM"), ("40", "Service catalog & provisioning", "network fulfillment")]
for i, (seq, name, ex) in enumerate(spokes):
    y = 1.9 + i * 1.18
    box(s, 10.35, y, 2.45, 1.0, [name, ex], fill=WHITE, color=DARK, size=13, line=TEAL)
    box(s, 12.4, y - 0.12, 0.5, 0.36, seq, fill=RED, size=10, shape=MSO_SHAPE.OVAL)
text(s, 0.6, 6.6, 12.3, 0.4,
     "Red numbers = configurable publish sequence (CRM \u2192 Commerce \u2192 Billing \u2192 Provisioning), "
     "the example used in Oracle's documentation.", size=12, color=GREY, italic=True)
notes(s, """(2:30 – 3:30)
ده أهم شكل في العرض. الـ Architecture اسمها Hub-and-Spoke.
في النص عندنا Launch، وده الـ Master: أي منتج بيتصمم هنا بس.
على الشمال: مين بيدخّل البيانات؟ فرق الماركتنج والمنتجات من خلال الواجهة، أو أنظمة تانية من خلال الـ APIs، وفي البداية بنعمل Import مرة واحدة للكتالوج الحالي من الـ CRM.
على اليمين: الأنظمة اللي بتستقبل (Spokes): الـ CRM، قنوات البيع الديجيتال، الـ Billing، والـ Provisioning. النشر بيعدي على طبقة Integration، وليه ترتيب بنحدده: مثلًا CRM الأول، بعدين Commerce، بعدين Billing، بعدين Provisioning.
النقطة المهمة: الأنظمة دي بتستهلك المنتج بس، مش بتعدّل فيه.""")

# ---------------------------------------------------------------- 5. Modules
s = new_slide("02  Modules & Components", "Work areas and what each one does")
mod_rows = [
    ["Module / work area", "What it does", "Main user"],
    ["Initiatives", "Container for a market release: team, attachments, approvers, notifications, lifecycle status, publish history", "Product Manager"],
    ["Offer Management", "Guided flow to create, revise, clone and bulk-update simple and bundle offers", "Product / Marketing Mgr"],
    ["Promotion Management", "Promotions with events and benefits (e.g. discount for 3 months)", "Marketing Manager"],
    ["Workbench", "Product specifications, service specifications, attributes (the technical side)", "Product Specialist"],
    ["Pricing & Price Lists", "One-time, recurring and usage fees, discounts, price lists by segment or location", "Product Manager"],
    ["Rules", "Eligibility, compatibility, recommendation and migration (upgrade/downgrade) rules", "Product Manager"],
    ["Administration", "Catalogs and categories, lifecycle states, publish destinations, roles", "Catalog Admin"],
    ["Reports + AI agent", "Dashboards for initiatives and offers; AI \"Offer Design Integrity\" checks before publishing", "All / Admin"],
]
table(s, 0.6, 1.6, 12.2, mod_rows, [2.5, 7.5, 2.2], size=13, row_h=0.56)
notes(s, """(3:30 – 4:30)
دي الـ Modules الأساسية في النظام:
- الـ Initiative: ده أهم مفهوم، وهو زي "المشروع" أو "الإطلاق". أي منتج لازم يكون جوه Initiative، وهي اللي بتشيل الفريق والمرفقات والموافقات والحالة.
- الـ Offer Management: بنعمل فيه العروض بخطوات مرتبة (Guided Flow).
- الـ Promotion Management: للعروض الترويجية.
- الـ Workbench: الجانب الفني، يعني مواصفات المنتج والخدمة.
- الـ Pricing والـ Rules: التسعير وقواعد مين يشتري إيه وإيه ينفع مع إيه.
- الـ Administration: الإعدادات، والحالات، والأنظمة اللي هننشر عليها.
- وفي الإصدارات الأخيرة ضافوا AI Agent بيراجع العرض قبل النشر ويطلع الأخطاء.""")

# ---------------------------------------------------------------- 6. Data model
s = new_slide("02  Core Data Model", "Separate the commercial offer from the technical product",
              "This separation is what lets marketing change offers without touching the network")
box(s, 0.6, 1.9, 3.7, 0.5, "COMMERCIAL (what we sell)", fill=PURPLE, size=14)
box(s, 0.6, 2.55, 3.7, 1.15, ["Package", "e.g. Home bundle with a 12-month commitment"], fill=PURPLE_LIGHT, color=DARK, size=15)
down_arrow(s, 2.3, 3.78)
box(s, 0.6, 4.2, 3.7, 1.15, ["Commercial / Service Bundle", "groups simple offers, option groups"], fill=PURPLE_LIGHT, color=DARK, size=15)
down_arrow(s, 2.3, 5.43)
box(s, 0.6, 5.85, 3.7, 0.95, ["Simple Offer", "internet plan, device, add-on"], fill=PURPLE_LIGHT, color=DARK, size=15)
arrow(s, 4.45, 6.15, 0.7)
box(s, 5.3, 1.9, 3.3, 0.5, "TECHNICAL (what we deliver)", fill=TEAL, size=14)
box(s, 5.3, 5.85, 3.3, 0.95, ["Product Specification", "attributes, e.g. speed"], fill=WHITE, color=DARK, size=15, line=TEAL)
down_arrow(s, 6.8, 5.45, fill=TEAL)
s.shapes[-1].rotation = 180
box(s, 5.3, 4.2, 3.3, 1.15, ["Service Specification", "what the network must activate"], fill=WHITE, color=DARK, size=15, line=TEAL)
box(s, 5.3, 2.55, 3.3, 1.15, ["Resource", "equipment / network resource"], fill=WHITE, color=DARK, size=15, line=TEAL)
down_arrow(s, 6.8, 3.78, fill=TEAL)
s.shapes[-1].rotation = 180
box(s, 9.0, 1.9, 3.8, 0.5, "ATTACHED TO EVERY OFFER", fill=GREY, size=14)
bullets(s, 9.0, 2.55, 3.85, 4.3, [
    "Prices: fees, discounts, price lists",
    "Terms: commitment period, penalty",
    "Rules: eligibility, compatibility, migration",
    "Availability: channel, location, market segment",
    "Lifecycle: status, version, start and end dates",
    "Catalog: catalog, category, product line",
], size=15)
notes(s, """(4:30 – 5:15)
النقطة دي مهمة جدًا لو هنبني نظام: Oracle بيفصل بين حاجتين.
الجانب التجاري (الشمال): ده اللي بنبيعه. عندنا Simple Offer زي باقة إنترنت أو راوتر، وBundle بيجمع كذا Offer، وPackage فيها التزام زي 12 شهر.
الجانب الفني (النص): Product Specification، وتحتها Service Specification، وده اللي الشبكة محتاجة تفعّله.
ليه الفصل ده مهم؟ لأن الماركتنج يقدر يعمل 10 عروض مختلفة على نفس الخدمة الفنية من غير ما يلمس الشبكة.
وعلى اليمين، كل عرض بيتربط بيه: السعر، والالتزام، والقواعد، وكمان القناة والمكان والشريحة. يعني مثلًا عرض متاح أونلاين بس، أو في محافظات معينة بس.""")

# ---------------------------------------------------------------- 7. Lifecycle
s = new_slide("03  Product Lifecycle", "End-to-end journey of an offer, from idea to obsolete",
              "Statuses are configurable; these are Oracle's default (seeded) statuses")
stages = [
    ("Idea /\nInitiative", "Marketing opens an initiative", GREY),
    ("In Design", "Define offer, price, rules", PURPLE),
    ("Approvals\nPending", "Optional approval step", AMBER),
    ("Ready to\nPublish", "Locked; sent to test systems", PURPLE),
    ("In Test", "QA in SIT / UAT", PURPLE),
    ("Active", "Promoted to production", PURPLE),
    ("Launched", "Sellable in channels", GREEN),
]
w, gap = 1.55, 0.22
for i, (name, desc, color) in enumerate(stages):
    x = 0.6 + i * (w + gap)
    box(s, x, 1.95, w, 1.0, name.split("\n"), fill=color, size=15)
    text(s, x - 0.05, 3.0, w + 0.1, 0.7, desc, size=12, color=GREY, align=PP_ALIGN.CENTER)
    if i < len(stages) - 1:
        arrow(s, x + w + 0.01, 2.32, gap - 0.02, 0.26)
text(s, 0.6, 3.8, 6, 0.4, "After launch: change or end of life", size=17, color=PURPLE, bold=True)
after = [("Revise", "new version starts In Design", TEAL), ("Awaiting\nRetirement", "linked to an initiative", AMBER),
         ("Retired", "no new sales", RED), ("Obsolete", "archived", GREY)]
for i, (name, desc, color) in enumerate(after):
    x = 0.6 + i * 3.1
    box(s, x, 4.3, 2.6, 0.95, name.split("\n"), fill=color, size=15)
    text(s, x, 5.3, 2.6, 0.4, desc, size=12, color=GREY, align=PP_ALIGN.CENTER)
    if i < len(after) - 1:
        arrow(s, x + 2.65, 4.62, 0.4, 0.28)
bullets(s, 0.6, 5.85, 12.3, 1.1, [
    "Rule: content is designed in the test instance only, then migrated to production; each status controls editing, approval and publishing",
    "Rule: only one launched version of an offer at a time; you can retire an offer only when no version is still in design or test",
], size=13, space=3)
notes(s, """(5:15 – 6:45)
دي رحلة المنتج كاملة، وهي قلب الـ PLC.
1. الفكرة بتبدأ من الماركتنج، وبتتسجل كـ Initiative.
2. In Design: هنا بنعرّف المنتج والسعر والقواعد، وده الوقت الوحيد اللي ينفع نعدّل فيه.
3. Approvals Pending: لو الموافقات مفعّلة، المنتج بيتقفل لحد ما المعتمدين يوافقوا.
4. Ready to Publish: المنتج اتقفل وبيتبعت لأنظمة الاختبار.
5. In Test: بنختبره في بيئة SIT أو UAT. لو فيه مشكلة يرجع In Design.
6. Active: بيتنقل للـ Production.
7. Launched: متاح للبيع في القنوات.
بعد الإطلاق: لو عايزين نعدّل بنعمل Revise وبيطلع Version جديد. ولو عايزين نوقف المنتج: Awaiting Retirement، بعدين Retired يعني مفيش بيع جديد، وفي الآخر Obsolete.
والحالات دي مش ثابتة، الأدمن يقدر يضيف حالة جديدة، زي مرحلة UAT إضافية.""")

# ---------------------------------------------------------------- 8. Ideation & definition
s = new_slide("03  Ideation & Definition", "How an idea from marketing becomes a defined product")
box(s, 0.6, 1.6, 5.9, 0.5, "Product Ideation: the Initiative", fill=PURPLE, size=16)
bullets(s, 0.6, 2.25, 5.9, 3.0, [
    "Marketing / product manager creates an initiative for a market release",
    "Attach the business case, requirements and design documents",
    "Add the team and the approvers for all offers in the initiative",
    "Notifications on assignment, approval requests and publish status",
    "Dashboard shows every initiative and its status",
], size=15)
box(s, 6.9, 1.6, 5.9, 0.5, "Product Definition: the guided offer flow", fill=PURPLE, size=16)
steps = ["General info: name, type, dates, catalog", "Attributes from the product specification",
         "Components and option groups (bundles)", "Pricing and commitment terms",
         "Rules: eligibility, compatibility, migration", "Channels, locations, market segments",
         "Approvals, then Confirm"]
for i, st in enumerate(steps):
    y = 2.25 + i * 0.47
    box(s, 6.9, y, 0.42, 0.4, str(i + 1), fill=TEAL, size=13, shape=MSO_SHAPE.OVAL)
    text(s, 7.45, y + 0.02, 5.4, 0.4, st, size=15)
box(s, 0.6, 5.75, 12.2, 1.05,
    ["Gap we noticed (an opportunity for our own system)",
     "Oracle starts once an initiative exists. It has no idea backlog, idea scoring or business-case approval "
     "before design. If we build in-house, an Idea & Business Case module is a real value-add."],
    fill=PURPLE_LIGHT, color=DARK, size=15, align=PP_ALIGN.LEFT)
notes(s, """(6:45 – 7:45)
Product Ideation: في Oracle، الفكرة بتتحول لـ Initiative. الماركتنج بيفتح Initiative، يرفق فيها الـ Business Case والمتطلبات، ويحدد الفريق والمعتمدين، والنظام بيبعت إشعارات لكل واحد.
Product Definition: تعريف المنتج بيتم من خلال Guided Flow، يعني شاشات بالترتيب: البيانات العامة، الخصائص، المكونات لو Bundle، السعر، القواعد، القنوات والأماكن، وفي الآخر الموافقات.
ملاحظة مهمة لينا: Oracle بيبدأ من لحظة ما الفكرة اتحولت لمشروع فعلًا. مفيش عنده مكان نجمع فيه الأفكار ونقيّمها ونوافق على الـ Business Case قبل التصميم. فلو هنبني نظام، الـ Module ده ممكن يكون إضافة حقيقية عندنا.""")

# ---------------------------------------------------------------- 9. Approvals
s = new_slide("04  Workflow & Approvals", "Who approves what, and how the workflow runs")
flow = [("Designer\nconfirms offer", PURPLE), ("Design\nComplete", PURPLE), ("Initiate\napprovals", AMBER),
        ("Approvers notified\n(email + UI)", AMBER), ("Approve / reject\nwith comments", AMBER), ("Ready to\nPublish", GREEN)]
for i, (name, color) in enumerate(flow):
    x = 0.6 + i * 2.07
    box(s, x, 1.65, 1.8, 0.95, name.split("\n"), fill=color, size=13)
    if i < len(flow) - 1:
        arrow(s, x + 1.82, 1.97, 0.23, 0.28)
bullets(s, 0.6, 2.9, 6.0, 3.9, [
    "Two modes: sequential (one approver after another) or parallel (all at once, with a voting %, e.g. 50%)",
    "Approvers set once per initiative, with overrides per offer",
    "Offer is locked while in Approvals Pending",
    "Rejected: back to In Design with comments",
    "Full approval history kept on every offer",
    "Admin can make approval mandatory for any lifecycle status",
], size=15)
rows = [["Stage", "Proposed approver (example)"],
        ["Business case", "Marketing director"],
        ["Offer & price", "Product mgmt + Finance / pricing"],
        ["Terms & regulation", "Legal & regulatory (NTRA, where needed)"],
        ["Technical feasibility", "IT / network / provisioning"],
        ["Go-live", "Commercial head + QA sign-off"]]
text(s, 6.9, 2.85, 5.9, 0.4, "Example approval matrix for our team", size=15, color=PURPLE, bold=True)
table(s, 6.9, 3.3, 5.9, rows, [2.3, 3.6], size=13, row_h=0.55)
notes(s, """(7:45 – 8:45)
الـ Approval Workflow في Oracle ماشي كده: اللي بيصمم بيعمل Confirm، المنتج بيبقى Design Complete، بعدين نبدأ الموافقات، المعتمدين بيوصلهم إيميل، ويوافقوا أو يرفضوا من جوه النظام مع تعليق، ولو كله وافق يبقى Ready to Publish.
فيه نوعين: Sequential، يعني واحد ورا التاني، أو Parallel، يعني كلهم مع بعض، ونقدر نحدد نسبة، مثلًا 50% بس محتاجين يوافقوا.
والمنتج بيتقفل وهو مستني الموافقة، فمحدش يقدر يعدّل فيه، وكل الموافقات بتتسجل في History.
على اليمين اقتراح لمصفوفة الموافقات عندنا. ده مثال مني للنقاش، مش من Oracle. Oracle بيدينا المرونة، وإحنا نحدد مين يوافق على إيه.""")

# ---------------------------------------------------------------- 10. Pricing
s = new_slide("04  Pricing Management", "Flexible pricing, controlled through the same lifecycle")
cards = [
    ("Price types", ["One-time fee (installation)", "Recurring fee (monthly)", "Usage fee (per GB / minute)", "Rollover of allowances"]),
    ("Price lists", ["Standard prices per segment", "e.g. Residential vs Business", "Or per location", "GL ID and tax code per price"]),
    ("Discounts & terms", ["Price alterations on bundles", "Time-limited discounts", "Multi-line / volume tiers", "Commitment terms with penalties"]),
    ("Changing prices", ["Revise the launched offer", "New version goes through approval", "Bulk update by amount or %", "Bundles re-point to new versions"]),
]
for i, (title, lines) in enumerate(cards):
    x = 0.6 + i * 3.08
    box(s, x, 1.65, 2.85, 0.55, title, fill=PURPLE, size=16)
    bullets(s, x, 2.3, 2.9, 2.6, lines, size=14, space=4)
box(s, 0.6, 5.0, 12.2, 1.75,
    ["Example from Oracle's docs, adapted",
     "Multi-line plan: 1 line = 80 per month, 2 lines = 75 per line, 3+ lines = 70 per line.  "
     "Price increase: select several plans and raise them by 10% in one bulk revise, then send them for approval."],
    fill=PURPLE_LIGHT, color=DARK, size=15, align=PP_ALIGN.LEFT)
notes(s, """(8:45 – 9:45)
التسعير في Oracle مرن جدًا:
- أنواع الأسعار: رسوم مرة واحدة زي التركيب، ورسوم شهرية، ورسوم استخدام لكل جيجا أو دقيقة.
- الـ Price Lists: قائمة أسعار لكل شريحة، يعني منزلي غير شركات، أو حسب المكان. وكل سعر ليه كود محاسبي وكود ضريبة.
- الخصومات: خصم لفترة محدودة، أو خصم على عدد الخطوط، أو التزام بمدة وغرامة لو العميل لغى بدري.
- تغيير السعر: مينفعش نغيّر سعر منتج شغال كده على طول. لازم نعمل Revise، وده بيطلع Version جديد، والـ Version ده لازم يعدّي على الموافقات تاني. ولو عايزين نزوّد أسعار كذا باقة مرة واحدة، فيه Bulk Update بمبلغ أو نسبة.
ده بيدّي الإدارة تحكم كامل في أي تعديل سعر.""")

# ---------------------------------------------------------------- 11. Versioning & RBAC
s = new_slide("04  Versioning & Role-Based Access", "Every change is traced; every role sees only what it should")
box(s, 0.6, 1.6, 5.4, 0.5, "Product Versioning", fill=PURPLE, size=16)
bullets(s, 0.6, 2.25, 5.5, 4.5, [
    "Revise: a launched offer gets a new version in In Design",
    "Compare: two offers, or two versions side by side",
    "Clone: start a new offer from an existing one",
    "Bulk revise: many launched offers at once",
    "One launched version at a time; the old one is replaced on publish",
    "Publish history kept per initiative",
], size=15)
roles = [["Role", "View", "Create / edit", "Delete", "Lifecycle & publish"],
         ["Catalog Administrator", "\u2713", "\u2713", "\u2713", "\u2713 (can override status)"],
         ["Catalog Product Manager", "\u2713", "\u2713", "\u2717", "\u2713 manage initiatives"],
         ["Marketing Manager", "\u2713", "\u2713 offers", "\u2717", "\u2717 no initiatives"],
         ["Product Specialist", "\u2713", "\u2713 specs only", "\u2717", "Confirm only"],
         ["Catalog Viewer", "\u2713", "\u2717", "\u2717", "\u2717 read-only"]]
box(s, 6.4, 1.6, 6.4, 0.5, "Role-Based Access (seeded job roles)", fill=PURPLE, size=16)
table(s, 6.4, 2.25, 6.4, roles, [2.2, 0.7, 1.2, 0.7, 1.6], size=12, row_h=0.55)
text(s, 6.4, 5.65, 6.4, 1.2,
     ["Plus: a \"sensitive\" flag on production destinations, so only privileged users can publish to production. "
      "Custom roles can be built in the Security Console."], size=13, color=GREY)
notes(s, """(9:45 – 10:45)
الـ Versioning: أي منتج متطلق لو عايزين نعدّله بنعمل Revise، فبيطلع Version جديد. ونقدر نقارن بين أي Versionين جنب بعض ونشوف الفرق. وفيه Clone لو عايزين منتج جديد شبه منتج قديم، وBulk Revise لكذا منتج مرة واحدة.
الـ Role-Based Access: Oracle جاي فيه خمس أدوار جاهزة:
- الـ Admin: يعمل كل حاجة، حتى إنه يغيّر حالة المنتج يدويًا.
- الـ Product Manager: يعمل ويعدّل بس ميمسحش، ويدير الـ Initiatives.
- الـ Marketing Manager: يعمل عروض بس ميقدرش يدير الـ Initiatives أو ينشر.
- الـ Product Specialist: الجانب الفني بس.
- الـ Viewer: قراءة بس، ودي للإدارة العليا مثلًا.
وفيه حماية إضافية: بيئة الـ Production ممكن نعلّمها Sensitive، فمحدش ينشر عليها غير المسموح لهم.""")

# ---------------------------------------------------------------- 12. Launch / retire / migrate
s = new_slide("03  Launch, Retirement & Migration", "Bringing an offer to market, and taking it out safely")
cols = [
    ("Product Launch", GREEN, [
        "Initiative moves to Active in production",
        "Publish to production CRM, billing and provisioning",
        "Status becomes Launched: sellable",
        "Availability limited by channel, location and segment",
        "Storefront banners to highlight offers",
    ]),
    ("Product Retirement", RED, [
        "Allowed only when all versions are Launched",
        "Link the offer to an initiative: Awaiting Retirement",
        "On publish: Retired, not sellable in any channel",
        "Later: Obsolete (archived)",
        "Existing subscribers are not removed automatically",
    ]),
    ("Product Migration", TEAL, [
        "Migration rules: Upgrade To / Upgrade From",
        "Defined between sellable bundles and packages",
        "Each rule: new commitment start, duration, penalty",
        "CRM uses the rules to offer the customer a valid move",
        "The move itself is executed by CRM and order management",
    ]),
]
for i, (title, color, lines) in enumerate(cols):
    x = 0.6 + i * 4.13
    box(s, x, 1.6, 3.9, 0.55, title, fill=color, size=17)
    bullets(s, x, 2.3, 3.95, 3.6, lines, size=14, space=5)
box(s, 0.6, 5.85, 12.2, 0.95,
    "Takeaway: Launch decides what is allowed (the offers and the rules). "
    "CRM and order management do the work on real customers.",
    fill=PURPLE_LIGHT, color=DARK, size=15)
notes(s, """(10:45 – 11:45)
الإطلاق: لما الـ Initiative تبقى Active في الـ Production، بننشر على الـ CRM والـ Billing والـ Provisioning، والمنتج بيبقى Launched. ونقدر نحدد يتباع فين: أونلاين بس، أو في فروع معينة، أو لشريحة معينة.
الإيقاف (Retirement): مينفعش نوقف منتج لو فيه Version منه لسه في التصميم أو الاختبار. بنربطه بـ Initiative، ولما ننشر بيبقى Retired، يعني مفيش بيع جديد من أي قناة. بس العملاء الحاليين مش بيتشالوا لوحدهم.
التحويل (Migration): هنا بنعرّف قواعد Upgrade وDowngrade: من الباقة القديمة للجديدة، ومعاها مدة الالتزام والغرامة. الـ CRM بيستخدم القواعد دي علشان يعرض على العميل التحويلات المسموحة بس.
الخلاصة: الـ PLC بيحدد المسموح، والـ CRM ونظام الطلبات هما اللي بينفذوا على العميل.""")

# ---------------------------------------------------------------- 13. CRM integration
s = new_slide("05  CRM Integration", "How a product reaches the CRM",
              "Publishing is initiative-based, ordered by sequence, and tracked per destination")
pipeline = [("1. Register\ndestinations", "CRM, billing, etc. per\nlifecycle status (SIT, UAT, Prod)"),
            ("2. Package\nin initiative", "offers, prices,\nrules, promotions"),
            ("3. Publish", "event sent to each\nspoke's API"),
            ("4. Spoke\nconfirms", "status moves on only\nafter all confirm"),
            ("5. Monitor &\nrepublish", "retry only the\nfailed destinations")]
for i, (name, desc) in enumerate(pipeline):
    x = 0.6 + i * 2.5
    box(s, x, 1.9, 2.15, 0.95, name.split("\n"), fill=PURPLE, size=14)
    text(s, x, 2.9, 2.15, 0.75, desc.split("\n"), size=12, color=GREY, align=PP_ALIGN.CENTER)
    if i < len(pipeline) - 1:
        arrow(s, x + 2.18, 2.22, 0.3, 0.28)
bullets(s, 0.6, 3.85, 6.2, 3.0, [
    "Prebuilt Siebel CRM integration: catalog, products, price lists, promotions, rules",
    "One-time migration of the existing CRM catalog into Launch",
    "Any CRM spoke can integrate by implementing one publish-event REST endpoint",
    "Exclusion rules: send each system only what it needs",
], size=14, space=5)
box(s, 7.1, 3.85, 5.7, 2.95,
    ["Golden rule of the integration",
     "Anything published by Launch must not be edited (\"enriched\") in the CRM. "
     "The CRM may only add things Launch does not own.",
     "",
     "So the PLC stays the single source of truth, and the CRM, billing and network "
     "always show the same product."],
    fill=PURPLE_LIGHT, color=DARK, size=15, align=PP_ALIGN.LEFT)
notes(s, """(11:45 – 12:45)
إزاي المنتج بيوصل للـ CRM؟ في 5 خطوات:
1. نسجّل الأنظمة اللي هننشر عليها، ولكل بيئة: اختبار، وUAT، وProduction.
2. نجمّع العروض والأسعار والقواعد جوه Initiative.
3. ننشر: النظام بيبعت Event لكل نظام على API معين.
4. الحالة مش بتتحرك غير لما كل الأنظمة ترد إنها استلمت.
5. لو نظام فشل، نعيد النشر عليه هو بس.
Oracle عنده ربط جاهز مع Siebel CRM، وأي CRM تاني ممكن يتربط لو عمل API واحد بيستقبل الـ Event.
أهم قاعدة: أي حاجة جاية من الـ PLC ممنوع تتعدّل في الـ CRM. ده اللي بيضمن إن المنتج نفسه في كل الأنظمة، وده غالبًا من أكبر المشاكل اللي بتقابل أي شركة اتصالات.""")

# ---------------------------------------------------------------- 14. Use case
s = new_slide("06  Telecom Use Case", "Launching \"Home Fiber 200\" and retiring a legacy ADSL package",
              "Illustrative scenario for our team; names and numbers are examples, not real tariffs")
uc = [
    ["#", "Step", "Who", "In the system", "Status"],
    ["1", "Idea: upsell ADSL customers to fiber", "Marketing", "Create initiative \"Fiber Upsell Q1\"; attach business case; add team and approvers", "In Design"],
    ["2", "Define product", "Product Specialist", "Product spec \"FTTH Access 200M\" linked to its service spec", "In Design"],
    ["3", "Define offer", "Product Manager", "Simple offers (internet, landline, router) packaged in \"Home Fiber 200\" with a 12-month term; channels: shops, call center, app; FTTH-covered areas only", "In Design"],
    ["4", "Price", "Product Mgr + Finance", "Monthly fee, installation fee, 3-month launch discount; residential price list", "In Design"],
    ["5", "Rules", "Product Manager", "Requires router; excludes ADSL on same line; migration rule: Upgrade From \"ADSL 30\"", "In Design"],
    ["6", "Approve", "Marketing, Finance, Legal", "Parallel approval, 100% vote required", "Approvals Pending"],
    ["7", "Test", "QA / IT", "Publish to test CRM, billing, provisioning; end-to-end order test", "In Test \u2192 Active"],
    ["8", "Launch", "Catalog Admin", "Publish to production; agents and app can sell it", "Launched"],
    ["9", "Retire ADSL 30 and migrate", "PM + CRM team", "Retire old package (no new sales); CRM offers the upgrade path to its customers", "Retired"],
]
table(s, 0.45, 1.65, 12.45, uc, [0.35, 2.5, 1.9, 5.85, 1.85], size=11, row_h=0.5)
notes(s, """(12:45 – 14:00)
خلّوني أوضح الكلام ده بمثال من شغلنا. السيناريو ده مثال للتوضيح، والأسماء والأرقام مش أسعار حقيقية.
الفكرة: الماركتنج عايز يحوّل عملاء الـ ADSL للفايبر.
1. الماركتنج بيفتح Initiative ويرفق الـ Business Case.
2. الـ Product Specialist بيعرّف المواصفات الفنية: خدمة FTTH بسرعة 200.
3. الـ Product Manager بيعمل العرض: إنترنت وأرضي وراوتر في Package واحدة بالتزام سنة، متاحة في الفروع وخدمة العملاء والأبلكيشن، وفي المناطق اللي فيها فايبر بس.
4. التسعير: شهري، وتركيب، وخصم أول 3 شهور، ومعاه موافقة المالية.
5. القواعد: لازم راوتر، ومينفعش مع ADSL على نفس الخط، وقاعدة تحويل من ADSL 30.
6. الموافقات: ماركتنج ومالية وقانوني مع بعض.
7. الاختبار على بيئة الاختبار.
8. الإطلاق على الـ Production.
9. ولما الباقة القديمة تخلص، بنعملها Retire، والـ CRM يعرض على عملائها التحويل للفايبر.
كده غطّينا Ideation وDefinition وApproval وPricing وLaunch وRetirement وMigration في سيناريو واحد.""")

# ---------------------------------------------------------------- 15. Conclusion / blueprint
s = new_slide("07  Conclusion", "If we build our own PLC: the modules we need",
              "Answer to the key question, based on what Oracle Launch does well and where it stops")
phase1 = [
    ("Idea & Business Case", "backlog, scoring, go / no-go (gap in Oracle)"),
    ("Product & Offer Catalog", "commercial offer separate from technical spec"),
    ("Pricing & Price Lists", "fees, discounts, segments, approvals on change"),
    ("Lifecycle Engine", "configurable statuses and transitions"),
    ("Approval Workflow", "sequential / parallel, per-stage approvers, audit"),
    ("Roles & Permissions", "view / edit / approve / publish per role"),
    ("Versioning", "revise, compare, clone, history"),
    ("Publishing to CRM & Billing", "TMF620 API, sequence, retry, single source of truth"),
]
phase2 = [
    ("Rules Engine", "eligibility, compatibility, migration"),
    ("Retirement & Migration", "end-of-sale, upgrade paths, customer moves"),
    ("Promotions", "time-bound campaigns"),
    ("Dashboards & Bulk Updates", "time-to-market, bulk price changes"),
    ("AI Validation", "check offers before publish"),
]
box(s, 0.6, 1.75, 7.4, 0.5, "Phase 1: must-have (MVP)", fill=PURPLE, size=16)
for i, (name, desc) in enumerate(phase1):
    y = 2.35 + i * 0.555
    box(s, 0.6, y, 0.42, 0.42, str(i + 1), fill=PURPLE, size=12, shape=MSO_SHAPE.OVAL)
    text(s, 1.12, y + 0.02, 2.75, 0.45, name, size=14, bold=True)
    text(s, 3.85, y + 0.04, 4.2, 0.45, desc, size=13, color=GREY)
box(s, 8.4, 1.75, 4.4, 0.5, "Phase 2: next", fill=TEAL, size=16)
for i, (name, desc) in enumerate(phase2):
    y = 2.35 + i * 0.62
    text(s, 8.45, y, 4.4, 0.32, name, size=14, bold=True)
    text(s, 8.45, y + 0.28, 4.4, 0.32, desc, size=12, color=GREY)
box(s, 8.4, 5.55, 4.4, 1.25,
    ["Design principles", "One source of truth  |  Test before production  |  TM Forum data model  |  Full audit trail"],
    fill=PURPLE_LIGHT, color=DARK, size=14)
notes(s, """(14:00 – 15:00)
ودي إجابة السؤال الأساسي: لو هنبني PLC عندنا، نحط فيه إيه؟
المرحلة الأولى، الأساسيات:
1. Module للأفكار والـ Business Case، وده اللي ناقص في Oracle، فممكن يبقى ميزة عندنا.
2. كتالوج المنتجات مع فصل الجانب التجاري عن الفني.
3. التسعير وقوائم الأسعار.
4. Lifecycle Engine بحالات نقدر نغيّرها.
5. Approval Workflow مع Audit كامل.
6. الصلاحيات حسب الدور.
7. الـ Versioning.
8. النشر على الـ CRM والـ Billing بمعيار TMF620.
المرحلة التانية: محرك القواعد، والإيقاف والتحويل، والعروض الترويجية، والتقارير، والـ AI.
والمبادئ: مصدر واحد للحقيقة، واختبار قبل الـ Production، ونمشي على معايير TM Forum، وكل حاجة متسجلة.""")

# ---------------------------------------------------------------- 16. Comparison scorecard
s = new_slide("Team Comparison", "Scorecard for comparing the three systems",
              "Each team member fills in the column for the system they studied")
score = [
    ["Feature", "Oracle Launch", "System 2 (Mostafa)", "System 3 (Hamdy)"],
    ["Product Ideation", "Partial: initiatives, no idea scoring", "", ""],
    ["Product Definition", "Strong: guided flow, spec vs offer", "", ""],
    ["Approval Workflow", "Strong: sequential / parallel, voting %", "", ""],
    ["Product Catalog", "Strong: simple, bundle, package", "", ""],
    ["Pricing Management", "Strong: price lists, tiers, bulk %", "", ""],
    ["Product Versioning", "Strong: revise, compare, clone", "", ""],
    ["Role-Based Access", "Strong: 5 seeded roles + custom", "", ""],
    ["Product Launch", "Strong: channel, location, segment", "", ""],
    ["CRM Integration", "Strong: prebuilt Siebel; open API", "", ""],
    ["Retirement", "Strong: lifecycle-controlled", "", ""],
    ["Migration", "Rules only; executed in CRM", "", ""],
    ["Standards", "TM Forum TMF620-based APIs", "", ""],
]
table(s, 0.6, 1.65, 12.2, score, [2.6, 4.0, 2.8, 2.8], size=12, row_h=0.4)
notes(s, """(اختياري – للنقاش بعد العرض)
ده جدول المقارنة اللي هنستخدمه إحنا التلاتة. أنا ملّيت عمود Oracle، ومصطفى وحمدي يملوا عمود النظام بتاعهم بنفس الـ Features اللي في التكليف، علشان المقارنة تبقى على نفس المقاييس.""")

# ---------------------------------------------------------------- 17. Sources / Q&A
s = new_slide(None, "Questions?")
text(s, 0.6, 1.5, 12, 0.4, "Sources (Oracle official documentation, releases 25B to 26D)", size=16, color=PURPLE, bold=True)
bullets(s, 0.6, 2.0, 12.3, 4.8, [
    "Launch Cloud Service User's Guide: Launch Overview, Product Offers, Rules, Entity Lifecycle Management",
    ("docs.oracle.com/en/industries/communications/launch-cloud-service/26c/user-guide/", 1),
    "Launch Cloud Service Implementation Guide: Set Up Users and Roles, Predefined Job Roles, Publish Catalogs",
    ("docs.oracle.com/en/industries/communications/launch-cloud-service/26d/implement-guide/", 1),
    "Launch Cloud Service Integration Guide: Integration with Siebel CRM",
    ("docs.oracle.com/en/industries/communications/launch-cloud-service/26b/integration/", 1),
    "Release notes 26A and What's New 26C (Initiatives enhancements, Offer Design Integrity AI agent)",
    "TM Forum: TMF620 Product Catalog Management API and TMFC001 Product Catalog Management component",
], size=14, space=4)
text(s, 0.6, 6.4, 12.2, 0.5,
     "Telecom use case, approval matrix and build blueprint are our own proposals for discussion, not Oracle content.",
     size=12, color=GREY, italic=True)
notes(s, """شكرًا. كل المعلومات اللي عن Oracle مصدرها الـ Documentation الرسمية بتاعتهم، والروابط موجودة هنا لو حد عايز يتأكد أو يتعمق.
المثال العملي ومصفوفة الموافقات وخطة البناء دي اقتراحات مني للنقاش، مش من Oracle.
أي أسئلة؟""")

out = "PLC_Oracle_Launch_Study.pptx"
prs.save(out)
print("saved", out, len(prs.slides), "slides")
