"""Builds the PLC study presentation (DISCOBOLE / ODACAT, open source).

Run:  python3 build_deck_disco.py
"""

from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

HERE = Path(__file__).parent
IMG = HERE / "images" / "disco"
OUT = HERE / "PLC_DISCOBOLE_Presentation.pptx"

ORANGE = RGBColor(0xE8, 0x77, 0x22)
LILAC = RGBColor(0xFD, 0xEB, 0xD8)
DARK = RGBColor(0x26, 0x22, 0x2E)
GREY = RGBColor(0x5F, 0x5A, 0x6B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
NAVY = RGBColor(0x1A, 0x1A, 0x1A)
GREEN = RGBColor(0x2E, 0x8B, 0x57)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
W, H = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]
TOTAL = 17
_page = [0]


def box(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, x, y, w, h)
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(1.25)
    s.shadow.inherit = False
    return s


def text(slide, x, y, w, h, content, size=16, color=DARK, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, shape=None):
    s = shape or slide.shapes.add_textbox(x, y, w, h)
    tf = s.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.08)
    tf.margin_top = tf.margin_bottom = Inches(0.04)
    items = content if isinstance(content, list) else [content]
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if isinstance(it, tuple):
            r = p.add_run()
            r.text = it[0]
            r.font.bold = True
            r.font.size = Pt(size)
            r.font.color.rgb = ORANGE if color == DARK else color
            r2 = p.add_run()
            r2.text = it[1]
            r2.font.size = Pt(size)
            r2.font.color.rgb = color
        else:
            r = p.add_run()
            r.text = it
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.color.rgb = color
        if len(items) > 1:
            p.space_after = Pt(size * 0.55)
    return s


def bullets(slide, x, y, w, h, items, size=17):
    s = slide.shapes.add_textbox(x, y, w, h)
    tf = s.text_frame
    tf.word_wrap = True
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(9)
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(Inches(0.3)))
        pPr.set("indent", str(-Inches(0.25)))
        pPr.append(pPr.makeelement(qn("a:buChar"), {"char": "•"}))
        lead, rest = it if isinstance(it, tuple) else ("", it)
        if lead:
            r = p.add_run()
            r.text = lead
            r.font.bold = True
            r.font.size = Pt(size)
            r.font.color.rgb = ORANGE
        r = p.add_run()
        r.text = rest
        r.font.size = Pt(size)
        r.font.color.rgb = DARK
    return s


def slide(title, subtitle=None):
    _page[0] += 1
    s = prs.slides.add_slide(BLANK)
    box(s, 0, 0, W, Inches(0.12), fill=ORANGE)
    text(s, Inches(0.6), Inches(0.28), Inches(12), Inches(0.7), title, size=28, color=ORANGE, bold=True)
    if subtitle:
        text(s, Inches(0.6), Inches(0.95), Inches(12), Inches(0.4), subtitle, size=15, color=GREY)
    text(s, Inches(0.6), Inches(7.05), Inches(9.5), Inches(0.3),
         "PLC study  |  DISCOBOLE ODACAT (open source): from idea to CRM", size=10, color=GREY)
    text(s, Inches(11.7), Inches(7.05), Inches(1.0), Inches(0.3), f"{_page[0]} / {TOTAL}", size=10,
         color=GREY, align=PP_ALIGN.RIGHT)
    return s


def picture(s, file, x, y, max_w, max_h, caption=None):
    iw, ih = Image.open(IMG / file).size
    scale = min(max_w / iw, max_h / ih)
    w, h = int(iw * scale), int(ih * scale)
    px = x + (max_w - w) // 2
    box(s, px - Inches(0.04), y - Inches(0.04), w + Inches(0.08), h + Inches(0.08),
        line=RGBColor(0xE0, 0xC8, 0xB0))
    s.shapes.add_picture(str(IMG / file), px, y, w, h)
    if caption:
        text(s, x, y + h + Inches(0.04), max_w, Inches(0.4), caption, size=10, color=GREY,
             align=PP_ALIGN.CENTER)


def arrow(s, x1, y1, x2, y2, color=ORANGE):
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    c.line.color.rgb = color
    c.line.width = Pt(2)
    ln = c.line._get_or_add_ln()
    ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle"}))
    return c


def table(s, x, y, w, rows, col_w, size=13, row_h=Inches(0.42)):
    shp = s.shapes.add_table(len(rows), len(rows[0]), x, y, w, row_h * len(rows))
    t = shp.table
    for c, cw in enumerate(col_w):
        t.columns[c].width = Inches(cw)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = t.cell(r, c)
            cell.text = val
            cell.margin_left = cell.margin_right = Inches(0.08)
            p = cell.text_frame.paragraphs[0]
            p.runs[0].font.size = Pt(size)
            cell.fill.solid()
            if r == 0:
                cell.fill.fore_color.rgb = ORANGE
                p.runs[0].font.bold = True
                p.runs[0].font.color.rgb = WHITE
            else:
                cell.fill.fore_color.rgb = LILAC if r % 2 else WHITE
                p.runs[0].font.color.rgb = DARK
                p.runs[0].font.bold = c == 0
    return t


def notes(s, t):
    s.notes_slide.notes_text_frame.text = t


# ------------------------------------------------------------------ 1 title
_page[0] += 1
s = prs.slides.add_slide(BLANK)
box(s, 0, 0, W, H, fill=NAVY)
box(s, Inches(0.8), Inches(2.15), Inches(0.12), Inches(2.5), fill=ORANGE)
text(s, Inches(1.1), Inches(2.05), Inches(11), Inches(0.5), "PLC STUDY  ·  OPEN SOURCE", size=16,
     color=ORANGE, bold=True)
text(s, Inches(1.1), Inches(2.55), Inches(11.5), Inches(1.4),
     "Product Lifecycle Management in DISCOBOLE", size=36, color=WHITE, bold=True)
text(s, Inches(1.1), Inches(4.0), Inches(11), Inches(0.7),
     "From idea to CRM — real screens, architecture and workflows", size=22, color=LILAC)
text(s, Inches(1.1), Inches(5.5), Inches(11), Inches(0.6),
     "ODACAT Product Catalog  ·  TM Forum ODA  ·  OW2 GitLab    |    Prepared by: [Your name]  ·  October 2026",
     size=14, color=GREY)
notes(s, "Open-source reference system. Every screenshot in this deck is from the official ODACAT UI.")

# ------------------------------------------------------------------ 2 agenda
s = slide("Agenda")
items = [
    ("01", "Why a PLC system", "What it gives an operator"),
    ("02", "PLC and TM Forum in brief", "The DISCOBOLE lifecycle"),
    ("03", "ODACAT at a glance", "Architecture, modules, data model"),
    ("04", "Using the system", "Eight stages, with the real UI"),
    ("05", "Integration with CRM", "How the product reaches the salesperson"),
    ("06", "Evidence", "Public code, docs and screens"),
]
for i, (n, t, d) in enumerate(items):
    col, row = i % 2, i // 2
    x, y = Inches(0.8 + col * 6.1), Inches(1.55 + row * 1.7)
    box(s, x, y, Inches(5.7), Inches(1.45), fill=LILAC)
    text(s, x + Inches(0.2), y + Inches(0.25), Inches(1.1), Inches(1.0), n, size=32, color=ORANGE, bold=True)
    text(s, x + Inches(1.35), y + Inches(0.3), Inches(4.1), Inches(0.5), t, size=20, color=DARK, bold=True)
    text(s, x + Inches(1.35), y + Inches(0.8), Inches(4.1), Inches(0.45), d, size=14, color=GREY)

# ------------------------------------------------------------------ 3 takeaways
s = slide("Why a PLC system: what it gives an operator")
cards = [
    ("One source of truth", "Specifications, offerings, prices and policy rules are defined once. CRM and ordering read them."),
    ("A governed journey", "inStudy → inDesign → inTest → active → launched. From active, a new version is required."),
    ("Business users configure", "Guided cards in the Catalog Admin Portal. No development step to add an offer."),
    ("Rules travel with the product", "Channel, market segment and ECA policy rules are evaluated at quote time."),
    ("Standard integration", "TMF620, TMF701, TMF723, TMF679, TMF760, TMF622 — any CRM can consume the catalog."),
    ("Inspectable", "Every screen, process-flow diagram and the Java source are public on OW2 GitLab."),
]
for i, (t, d) in enumerate(cards):
    col, row = i % 3, i // 3
    x, y = Inches(0.6 + col * 4.1), Inches(1.5 + row * 2.6)
    box(s, x, y, Inches(3.85), Inches(2.35), fill=LILAC)
    box(s, x, y, Inches(3.85), Inches(0.1), fill=ORANGE)
    text(s, x + Inches(0.15), y + Inches(0.25), Inches(3.55), Inches(0.55), t, size=18, color=ORANGE, bold=True)
    text(s, x + Inches(0.15), y + Inches(0.85), Inches(3.55), Inches(1.3), d, size=14, color=DARK)

# ------------------------------------------------------------------ 4 lifecycle
s = slide("PLC and the TM Forum standard in brief",
          "DISCOBOLE implements the TMF620 statuses with a dedicated Lifecycle Management process")
picture(s, "lifecycle-product-specification-product-offering.png", Inches(0.5), Inches(1.5),
        Inches(7.4), Inches(5.2), "Official DISCOBOLE Product Offering lifecycle")
bullets(s, Inches(8.1), Inches(1.7), Inches(4.7), Inches(5.0), [
    ("inStudy / inDesign: ", "the idea is being designed."),
    ("inTest: ", "created in the UI; not yet sellable."),
    ("active: ", "tested, ready to be used."),
    ("launched: ", "start of marketing; CRM can sell it."),
    ("retired / obsolete: ", "no new sales / no customers left."),
    ("Gate in the UI: ", "lifecycle icon → system returns only the allowed next statuses."),
], size=16)

# ------------------------------------------------------------------ 5 architecture
s = slide("DISCOBOLE at a glance",
          "Open-source ODA Core Commerce suite. Product Catalog (ODACAT) is the PLC.")
picture(s, "suite-architecture.png", Inches(0.4), Inches(1.45), Inches(8.3), Inches(5.3),
        "Source: DISCOBOLE overall architecture. Bold API numbers are exposed by the suite.")
table(s, Inches(8.85), Inches(1.5), Inches(4.0), [
    ["Component", "Role"],
    ["ODACAT", "Catalog + lifecycle"],
    ["Configurator", "Eligibility, discounts"],
    ["Order Management", "Capture & follow-up"],
    ["COOD", "Delivery orchestration"],
    ["Product Inventory", "What customers own"],
    ["Roles (TMF672)", "Who may approve"],
    ["Admin Portal", "The UI in this deck"],
], [1.8, 2.2], size=12, row_h=Inches(0.48))

# ------------------------------------------------------------------ 6 bottom-up
s = slide("Bottom-up catalog: what the administrator builds",
          "CFS and stock items are imported. Everything above them is mastered in ODACAT.")
picture(s, "frontend-product-catalog-ui-bottom-up-approach.png", Inches(0.5), Inches(1.5),
        Inches(6.2), Inches(5.2), "Official bottom-up approach")
picture(s, "data-model-tmf620.png", Inches(6.9), Inches(1.5), Inches(5.9), Inches(5.2),
        "TMF620 data model: commercial / functional / technical")

# ------------------------------------------------------------------ 7 real homepage
s = slide("The Catalog Admin Portal — the real home screen",
          "After login, the administrator works from one dashboard per entity")
picture(s, "ui-odacat-homepage.png", Inches(0.45), Inches(1.45), Inches(12.4), Inches(5.3),
        "Left: Offering, Price, Policy Rule, Category, Product Specification, CFS, Stock, Lifecycle. "
        "Main: type, status, version, and actions (JSON, hierarchy, lifecycle, edit).")

# ------------------------------------------------------------------ 8 stages 1
s = slide("Using the system (1/2): define the product",
          "Running example: Business Internet 100 — dedicated 100 Mbps + static IP + router + SLA")
picture(s, "ui-product-specification-dashboard.png", Inches(0.4), Inches(1.4), Inches(6.5), Inches(3.55))
picture(s, "ui-ho-po1-create1.png", Inches(7.05), Inches(1.4), Inches(5.8), Inches(5.3),
        "Create Product Offering: Atomic / Bundle / Contract, Visible, Sellable, Channel, Market Segment, term")
text(s, Inches(0.45), Inches(5.05), Inches(6.5), Inches(1.7), [
    ("1. Product Specification. ", "Restrict a CFS (internet) or Stock Item (router). Set characteristics (100 Mbps)."),
    ("2. Product Offering. ", "Atomic items, then a bundle, then the contract the customer signs."),
], size=14)

# ------------------------------------------------------------------ 9 stages 2
s = slide("Using the system (2/2): price, rules, launch",
          "Prices and policy rules stay in the catalog; sales never re-types them")
picture(s, "ui-product-offering-price-dashboard.png", Inches(0.4), Inches(1.4), Inches(6.4), Inches(3.4))
picture(s, "ui-policy-rule-dashboard.png", Inches(6.95), Inches(1.4), Inches(5.9), Inches(3.4))
text(s, Inches(0.5), Inches(5.0), Inches(12.3), Inches(1.8), [
    ("3. Product Offering Price. ", "Charge, Alteration (discount), Tax, Installment. Recurring monthly fee + one-time installation."),
    ("4. Policy Rule. ", "Event–Condition–Action (TMF723). Example: 24-month term unlocks the backup-link add-on."),
    ("5. Lifecycle. ", "inTest → active → launched. Sellable flag on. Only then does CRM see the offer."),
], size=15)

# ------------------------------------------------------------------ 10 hierarchy + lifecycle process
s = slide("What a finished product looks like — and how its status moves")
picture(s, "ui-offering-hierarchy-diagram.png", Inches(0.4), Inches(1.4), Inches(7.4), Inches(5.3),
        "Official example: contract hierarchy with lifecycle status on every node")
picture(s, "lifecycle.png", Inches(8.0), Inches(1.5), Inches(4.9), Inches(3.4),
        "Lifecycle process: system returns allowed next statuses")
bullets(s, Inches(8.05), Inches(5.15), Inches(4.8), Inches(1.6), [
    "Click the lifecycle icon on the dashboard.",
    "Modal: current status + allowed targets.",
    "Docs example: inTest → active.",
], size=14)

# ------------------------------------------------------------------ 11 CRM principle
s = slide("Integration with CRM: the catalog is the single source of truth",
          "Order capture consumes the approved product; it never re-creates it")
cat = box(s, Inches(0.55), Inches(1.7), Inches(3.5), Inches(3.8), fill=ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 0, 0, 0, 0, ["ODACAT", "Product Catalog", "", "Only launched + sellable offerings are published"],
     size=16, color=WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=cat)
arrow(s, Inches(4.1), Inches(3.6), Inches(5.05), Inches(3.6))
pub = ["Offering hierarchy (atomic / bundle / contract)", "Prices, discounts, commitment terms",
       "Channel, market segment, Sellable flag", "Policy rules (ECA)",
       "validFor dates", "Status and version"]
text(s, Inches(5.15), Inches(1.5), Inches(4.0), Inches(0.35), "What gets published", size=14, color=ORANGE, bold=True)
for i, p in enumerate(pub):
    b = box(s, Inches(5.15), Inches(1.95 + i * 0.55), Inches(4.0), Inches(0.48), fill=LILAC,
            shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, 0, 0, 0, 0, [p], size=12, color=DARK, anchor=MSO_ANCHOR.MIDDLE, shape=b)
arrow(s, Inches(9.2), Inches(3.6), Inches(10.1), Inches(3.6))
crm = box(s, Inches(10.2), Inches(1.7), Inches(2.7), Inches(3.8), fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 0, 0, 0, 0, ["CRM / Order Capture", "", "Quote, order and contract use the catalog definition"],
     size=15, color=WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=crm)
text(s, Inches(0.55), Inches(5.75), Inches(12.2), Inches(1.0),
     "TMF701 is how any front end — the admin portal or a CRM — is guided through create, modify and "
     "lifecycle. Kafka events keep Order Management, Inventory and the Configurator in sync.",
     size=14, color=GREY)

# ------------------------------------------------------------------ 12 APIs
s = slide("How the hand-over works: TM Forum Open APIs",
          "These are the APIs DISCOBOLE actually exposes (see architecture diagram)")
table(s, Inches(0.5), Inches(1.5), Inches(12.3), [
    ["API", "Direction", "What it does"],
    ["TMF620 Product Catalog", "Catalog → CRM / OM / Configurator", "Read offerings, prices, specifications"],
    ["TMF701 Process Flow", "Any front end → catalog", "Create, modify, clone, change lifecycle"],
    ["TMF723 Policy Management", "Catalog → Configurator", "ECA rules at configuration time"],
    ["TMF679 Offering Qualification", "CRM → Configurator", "Can this customer get this offer here?"],
    ["TMF760 Product Configuration", "CRM → Configurator", "Allowed combinations and discounts"],
    ["TMF622 Product Ordering", "CRM / Order Management", "Order carries offering id and version"],
    ["Kafka events", "Catalog → the suite", "Create / change / status without polling"],
], [3.3, 3.6, 5.4], size=14, row_h=Inches(0.52))
steps = ["Offering launched", "Event on Kafka", "CRM reads TMF620", "Configurator checks rules", "Order refs version"]
for i, st in enumerate(steps):
    x = Inches(0.5 + i * 2.55)
    b = box(s, x, Inches(5.85), Inches(2.3), Inches(0.8), fill=ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, 0, 0, 0, 0, [st], size=13, color=WHITE, bold=True, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, shape=b)
    if i < 4:
        arrow(s, x + Inches(2.32), Inches(6.25), x + Inches(2.52), Inches(6.25))

# ------------------------------------------------------------------ 13 salesperson
s = slide("What the salesperson inherits from the catalog",
          "Channel, market segment and Sellable are defined on the offering, not in CRM")
picture(s, "ui-ho-po1-create1.png", Inches(0.4), Inches(1.4), Inches(6.6), Inches(5.3),
        "Create Offering: Channel and Market Segment stay disabled until Sellable is on")
bullets(s, Inches(7.2), Inches(1.7), Inches(5.6), Inches(5.0), [
    ("Sellable + Visible ", "decide whether the offer appears in sales."),
    ("Channel / Market Segment ", "are the same eligibility dimensions CRM will filter on."),
    ("Commitment Term ", "12 / 24 / 36 months — prices and policy can differ by term."),
    ("Policy rules ", "are evaluated by the Configurator before a quote is submitted."),
    ("Hierarchy ", "is what the salesperson sees as the contract plus add-ons."),
], size=16)

# ------------------------------------------------------------------ 14 versioning
s = slide("When a product changes after launch", "Live offerings are not edited in place")
flow = [("1", "Clone or new version", "From launched: clone, or minor / major version"),
        ("2", "inTest again", "Version 0.1, name suffixed _CLONE"),
        ("3", "Re-approve", "Prices, policy, lifecycle inTest → launched"),
        ("4", "Event to CRM", "New quotes use v2; existing orders keep v1"),
        ("5", "Retire old", "validFor end, or status retired")]
for i, (n, t, d) in enumerate(flow):
    x = Inches(0.5 + i * 2.55)
    c = box(s, x + Inches(0.7), Inches(1.7), Inches(0.75), Inches(0.75), fill=ORANGE, shape=MSO_SHAPE.OVAL)
    text(s, 0, 0, 0, 0, [n], size=22, color=WHITE, bold=True, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, shape=c)
    if i < 4:
        arrow(s, x + Inches(1.5), Inches(2.07), x + Inches(3.2), Inches(2.07))
    text(s, x, Inches(2.6), Inches(2.35), Inches(0.5), t, size=15, color=ORANGE, bold=True, align=PP_ALIGN.CENTER)
    text(s, x, Inches(3.15), Inches(2.35), Inches(0.9), d, size=13, color=GREY, align=PP_ALIGN.CENTER)
box(s, Inches(0.5), Inches(4.4), Inches(12.3), Inches(2.2), fill=LILAC)
text(s, Inches(0.7), Inches(4.55), Inches(12.0), Inches(1.95), [
    ("Version rules in the Lifecycle Management service: ", "before active, edit in place; from active, a new version is required; rejected and obsolete are locked."),
    ("Minor (1.0 → 1.1): ", "only the latest stays launched; the previous is retired via validFor."),
    ("Major (2.1 → 3.0): ", "starts again in inStudy. Installed-base migration is a separate concern."),
], size=15)

# ------------------------------------------------------------------ 15 Dynamics
s = slide("Example: connecting the catalog to Microsoft Dynamics 365",
          "The same pattern applies to a CRM from another vendor")
b = box(s, Inches(0.55), Inches(1.55), Inches(2.7), Inches(1.15), fill=ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 0, 0, 0, 0, ["ODACAT", "TMF620 + events"], size=15, color=WHITE, bold=True,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=b)
arrow(s, Inches(3.3), Inches(2.12), Inches(4.05), Inches(2.12))
b = box(s, Inches(4.1), Inches(1.55), Inches(2.7), Inches(1.15), fill=GREY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 0, 0, 0, 0, ["Integration layer", "maps and syncs"], size=15, color=WHITE, bold=True,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=b)
arrow(s, Inches(6.85), Inches(2.12), Inches(7.6), Inches(2.12))
b = box(s, Inches(7.65), Inches(1.55), Inches(2.7), Inches(1.15), fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 0, 0, 0, 0, ["Dynamics 365", "Dataverse Web API"], size=15, color=WHITE, bold=True,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=b)
text(s, Inches(10.5), Inches(1.55), Inches(2.4), Inches(1.2),
     "No ready-made connector; built on the published APIs", size=12, color=GREY)
table(s, Inches(0.55), Inches(3.0), Inches(12.2), [
    ["Catalog item (ODACAT / TMF620)", "Dynamics 365 Sales entity"],
    ["Atomic Product Offering", "Product"],
    ["Bundle / Contract Product Offering", "Product bundle"],
    ["Category", "Product family"],
    ["Product Offering Price / Alteration", "Price list item / Discount list"],
    ["reliesOn, add-ons", "Product relationships"],
    ["launched / retired", "Active / Retired"],
], [6.1, 6.1], size=14, row_h=Inches(0.42))
text(s, Inches(0.55), Inches(6.35), Inches(12.2), Inches(0.45),
     "inStudy, inDesign, inTest stay in ODACAT. Dynamics receives only launched, sellable offerings.",
     size=14, color=ORANGE, bold=True)

# ------------------------------------------------------------------ 16 evidence
s = slide("Evidence: why this system can be studied end to end")
for i, (t, d) in enumerate([
    ("Real UI, documented click by click",
     "Login, dashboards, create wizards, hierarchy graph, lifecycle modal — all in the public Catalog UI guide."),
    ("Process flows, not inferred",
     "Create specification, create offering, create price, lifecycle: official BPMN-style diagrams."),
    ("Industry APIs, implemented",
     "TMF620, TMF701, TMF723 in the catalog; TMF679, TMF760, TMF622 in the suite. Architecture diagram labels them."),
    ("Source you can clone",
     "OW2 GitLab: Java 17, Spring Boot, Axon CQRS, Kafka, MongoDB, Keycloak. Designed as Orange's ODA catalog."),
]):
    y = Inches(1.45 + i * 1.3)
    box(s, Inches(0.6), y, Inches(12.1), Inches(1.15), fill=LILAC)
    box(s, Inches(0.6), y, Inches(0.12), Inches(1.15), fill=ORANGE)
    text(s, Inches(0.95), y + Inches(0.12), Inches(11.5), Inches(0.4), t, size=18, color=ORANGE, bold=True)
    text(s, Inches(0.95), y + Inches(0.52), Inches(11.5), Inches(0.5), d, size=14, color=DARK)

# ------------------------------------------------------------------ 17 close
s = slide("Summary and references")
bullets(s, Inches(0.55), Inches(1.4), Inches(6.7), Inches(5.3), [
    ("Idea to launched ", "inside ODACAT: specification, offering, price, policy, lifecycle, version."),
    ("Launched to CRM ", "through TMF620 + events. Configurator enforces the rules."),
    ("Prices and rules ", "are not re-typed in the sales screen."),
    ("Changes are versioned, ", "so existing customers keep what they bought."),
    ("The difference vs a closed vendor: ", "every screen and the source are public."),
], size=16)
refs = [
    "DISCOBOLE Product Catalog introduction",
    "Catalog UI (login, dashboards, wizards)",
    "Product Offering / Specification / Price / Policy UI",
    "Overall architecture (lifecycle, processes, model)",
    "Lifecycle Management service",
    "DISCOBOLE suite architecture",
    "OW2 GitLab: gitlab.ow2.org/discobole",
    "TM Forum TMF620",
    "Microsoft Learn: Dynamics 365 product catalog",
]
box(s, Inches(7.5), Inches(1.4), Inches(5.3), Inches(5.3), fill=LILAC)
text(s, Inches(7.7), Inches(1.5), Inches(4.95), Inches(0.4), "References (full URLs in the document)",
     size=14, color=ORANGE, bold=True)
text(s, Inches(7.7), Inches(2.05), Inches(4.95), Inches(4.4),
     [f"{i}. {r}" for i, r in enumerate(refs, 1)], size=13, color=DARK)

assert _page[0] == TOTAL, _page[0]
prs.save(OUT)
print("saved", OUT.name, "slides", _page[0])
