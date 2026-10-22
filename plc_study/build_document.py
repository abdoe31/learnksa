"""Builds the PLC pre-CRM study document (Netcracker Product Management).

Run:  pip install python-docx  &&  python3 build_document.py
"""

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

HERE = Path(__file__).parent
IMG = HERE / "images"
OUT = HERE / "PLC_Netcracker_PreCRM_Study.docx"

PURPLE = RGBColor(0x5B, 0x2C, 0x83)
GREY = RGBColor(0x5F, 0x5A, 0x6B)

doc = Document()
for section in doc.sections:
    section.left_margin = section.right_margin = Cm(2.2)
    section.top_margin = section.bottom_margin = Cm(2.0)

normal = doc.styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(11)
for name, size in (("Heading 1", 16), ("Heading 2", 13), ("Heading 3", 11.5)):
    st = doc.styles[name]
    st.font.name = "Calibri"
    st.font.size = Pt(size)
    st.font.color.rgb = PURPLE


def shade(cell, hex_fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tc_pr.append(shd)


def h1(t):
    doc.add_heading(t, level=1)


def h2(t):
    doc.add_heading(t, level=2)


def h3(t):
    doc.add_heading(t, level=3)


def para(t, italic=False, color=None, bold_lead=None):
    p = doc.add_paragraph()
    if bold_lead:
        r = p.add_run(bold_lead)
        r.bold = True
    r = p.add_run(t)
    r.italic = italic
    if color:
        r.font.color.rgb = color
    return p


def bullets(items, style="List Bullet"):
    for it in items:
        p = doc.add_paragraph(style=style)
        if isinstance(it, tuple):
            r = p.add_run(it[0])
            r.bold = True
            p.add_run(it[1])
        else:
            p.add_run(it)


def numbered(items):
    for i, it in enumerate(items, 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.8)
        p.paragraph_format.first_line_indent = Cm(-0.6)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(f"{i}.  ")
        r.bold = True
        r.font.color.rgb = PURPLE
        p.add_run(it)


def table(rows, widths=None, header_fill="5B2C83"):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = t.cell(r, c)
            cell.text = ""
            run = cell.paragraphs[0].add_run(val)
            run.font.size = Pt(9.5)
            if r == 0:
                run.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                shade(cell, header_fill)
            elif c == 0:
                run.bold = True
            if widths:
                cell.width = Cm(widths[c])
    if widths:
        t.autofit = False
        for c, w in enumerate(widths):
            t.columns[c].width = Cm(w)
    doc.add_paragraph()
    return t


def note(title, text):
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    cell = t.cell(0, 0)
    shade(cell, "EEE7F5")
    p = cell.paragraphs[0]
    r = p.add_run(title + "  ")
    r.bold = True
    r.font.color.rgb = PURPLE
    p.add_run(text).font.size = Pt(10)
    doc.add_paragraph()


def image(file, caption, width_cm=16):
    doc.add_picture(str(IMG / file), width=Cm(width_cm))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(caption)
    r.italic = True
    r.font.size = Pt(9)
    r.font.color.rgb = GREY


def stage(title, status, purpose, who, inputs, steps, outputs, evidence, gaps=None):
    h2(title)
    table([
        ["Item", "Detail"],
        ["Lifecycle status", status],
        ["Purpose", purpose],
        ["Who (proposed roles) [P]", who],
        ["Inputs", inputs],
        ["Exit criteria / output", outputs],
    ], widths=[4.2, 12.4])
    h3("What you do in the system (walkthrough)")
    numbered(steps)
    h3("Evidence")
    bullets(evidence)
    if gaps:
        h3("Not publicly documented (confirm in a vendor demo)")
        bullets(gaps)


# ------------------------------------------------------------------ cover
t = doc.add_paragraph()
t.alignment = WD_ALIGN_PARAGRAPH.LEFT
r = t.add_run("PLC STUDY  |  TELECOM EGYPT")
r.bold = True
r.font.color.rgb = PURPLE
r.font.size = Pt(11)
title = doc.add_paragraph()
r = title.add_run("Product Lifecycle Management in Netcracker:\nfrom Idea to Launch-Ready (before CRM integration)")
r.bold = True
r.font.size = Pt(22)
para("System studied: Netcracker Product Management (Product Catalog + Product Lifecycle Management), "
     "part of Netcracker Digital BSS / Digital Commerce & Monetization. Vendor: Netcracker Technology, "
     "a wholly owned subsidiary of NEC Corporation.", color=GREY)
para("Prepared by: [Your name]    |    Version: Draft 1 for review    |    October 2026", color=GREY)

note("How to read this document.",
     "Every statement is tagged with its source: [N] = documented by Netcracker; "
     "[S] = defined by the TM Forum TMF620 industry standard, which Netcracker is certified on; "
     "[P] = our proposal or an illustrative example, not vendor content. "
     "Netcracker's detailed user guides and demos are gated, so screen names in the walkthroughs are "
     "illustrative and will differ in a real installation.")

# ------------------------------------------------------------------ 1
h1("1. Scope of this document")
para("This document covers the product journey up to the point where the product is ready to be "
     "published to sales channels. It stops before integration with the CRM.")
table([
    ["In scope (this document)", "Out of scope (later phase)"],
    ["Ideation / concept (In Study)", "Publishing offers to CRM, CPQ and sales channels"],
    ["Technical product definition", "Order capture and order management"],
    ["Commercial offer definition", "Billing / charging activation in production"],
    ["Pricing", "Commercial launch to all customers"],
    ["Business rules and eligibility", "Retirement of old products"],
    ["Approval gates", "Migration of existing customers"],
    ["Versioning during design", ""],
    ["Testing and readiness (In Test, Active)", ""],
], widths=[8.3, 8.3])

# ------------------------------------------------------------------ 2
h1("2. PLC overview")
para("A Product Lifecycle Management (PLC/PLM) system is the single place where an operator turns an idea "
     "into a sellable product, governs every change to it, and finally retires it. Its value is that every "
     "other system (CRM, ordering, billing, network) reads the same product definition instead of keeping "
     "its own copy.")
h2("2.1 The industry-standard lifecycle [S]")
para("TM Forum's Product Catalog Management standard (TMF620) defines the statuses every catalog element "
     "goes through. Netcracker is certified on TMF620 (TM Forum certification list, June 2024).")
table([
    ["Status", "Meaning (TMF620)", "Gate to enter it", "In this document?"],
    ["In Study", "Macro conception of the product has started", "Idea registered", "Yes"],
    ["In Design", "Conception accepted; product is being designed", "Conception approved", "Yes"],
    ["In Test", "Design approved; product is being tested", "Design approved", "Yes"],
    ["Rejected", "Test failed; final status", "Test KO", "Yes"],
    ["Active", "Validated and tested, but not yet available to customers", "Test approved", "Yes (end point)"],
    ["Launched", "Start of marketing; customers can buy it", "Start of marketing", "No (CRM phase)"],
    ["Retired", "No new sales; existing customers keep it", "End of marketing", "No"],
    ["Obsolete", "No customer holds it; can be removed", "No customers left", "No"],
], widths=[2.6, 6.4, 4.0, 3.6])
para("Our scope therefore ends at Active: the product is fully defined, approved and tested, "
     "and is waiting to be published to sales channels.", italic=True)

# ------------------------------------------------------------------ 3
h1("3. Netcracker Product Management: modules used before CRM")
para("Netcracker groups product capabilities under the Product Management domain, with two core "
     "components [N]:")
bullets([
    ("Product Catalog: ", "\"a unified holistic view and easy access to up-to-date product information "
     "(pricing, discounts, fees, loyalty programs, taxes), supporting sales, ordering and billing processes "
     "via industry-standard integration.\" [N]"),
    ("Product Lifecycle Management: ", "\"end-to-end lifecycle management from marketing concept development "
     "and product configuration to product testing, test launch and subsequent mass production. It also "
     "fully supports product versioning, upgrades and retirement.\" [N]"),
])
image("01_tmforum_netcracker_module_map.png",
      "Figure 1. Netcracker module map from TM Forum's certification report. Product Management = Product "
      "Information Management (catalog) + Product Lifecycle Management. Source: TM Forum SID certification "
      "report, Netcracker TOMS Release 8.2 (2011; historical, shows the module split).")

h2("3.1 Module breakdown")
table([
    ["Module / capability", "What it does before CRM", "Source"],
    ["Product Lifecycle Management", "Drives the product through concept, configuration, testing and test "
     "launch; manages versions", "[N]"],
    ["Product Offering Catalog (commercial)", "Holds offering details, business rules, prices and discounts", "[N]"],
    ["Technical Product Catalog", "Holds product/service specifications, provisioning templates and "
     "decomposition rules (how an order is split into technical tasks)", "[N]"],
    ["Product templates & reusable components", "Pre-defined product definition templates and out-of-the-box "
     "components for rapid configuration", "[N]"],
    ["Pricing", "Prices, discounts, fees, taxes, loyalty programs", "[N]"],
    ["Business rules", "Eligibility by market, distribution channel and customer category; dependencies and "
     "constraints", "[N]"],
    ["Catalog federation", "Brings third-party partner catalogs into one product management interface", "[N]"],
    ["Analytics", "Analyzes product consumption patterns to target offers (input to ideation)", "[N]"],
    ["GenAI assistance", "AI-assisted product configuration with reusable templates", "[N]"],
], widths=[4.6, 9.8, 2.2])

image("02_netcracker_catalog_feeds_sales_and_orders.png",
      "Figure 2. Netcracker's split between the Product Offering Catalog (commercial) and the Technical "
      "Product Catalog. The left side (CPQ, Order Management) is the post-CRM phase. "
      "Source: Netcracker CPQ whitepaper (2021), Figure 2.")

h2("3.2 Core data model [S][N]")
para("Netcracker's product data model was certified by TM Forum against the SID information framework "
     "for these entities: Product, Product Price, Product Offering, Product Offering Price, Product Catalog "
     "and Product Specification.")
table([
    ["Entity", "Plain meaning", "Example [P]"],
    ["Product Specification", "What the product technically is, with its characteristics", "FTTH access, speed = 200 Mbps"],
    ["Product Offering", "What is sold: a single offer or a bundle of offers", "\"Home Fiber 200\" bundle"],
    ["Product Offering Price", "The prices attached to an offering", "Monthly fee, installation fee"],
    ["Product Catalog", "A grouping of offerings, e.g. per segment", "Residential catalog"],
    ["Product / Product Price", "What a customer actually owns and pays (after sale)", "Customer X's subscription"],
], widths=[4.2, 7.0, 5.4])

# ------------------------------------------------------------------ 4
doc.add_page_break()
h1("4. Stage-by-stage: how Netcracker applies PLC before CRM")
para("Each stage below lists its purpose, roles, inputs, a walkthrough of what the user does in the system, "
     "exit criteria, evidence, and what still needs confirmation. The running example is an illustrative "
     "home fiber bundle [P].")

stage(
    "4.1 Stage 1: Ideation / marketing concept",
    "In Study [S]",
    "Capture a product idea, check its business value, and decide whether it is worth designing.",
    "Marketing (owner), Product Manager, Finance (business case)",
    "Market need, competitor offers, consumption analytics, target segment",
    [
        "Product Manager opens Product Lifecycle Management and registers a new product concept "
        "(\"Fiber 200 + landline bundle for residential\").",
        "Attaches the business case: target segment, expected uptake, revenue goal.",
        "Reviews product consumption analytics (e.g. heavy-usage ADSL customers) to size the opportunity.",
        "Checks the existing catalog for similar products that can be reused instead of creating a new one.",
        "Submits the concept for the conception approval gate.",
    ],
    "Concept approved, so the status moves to In Design; or the concept is dropped.",
    [
        "Netcracker PLM starts at \"marketing concept development\" [N].",
        "Netcracker analyzes \"product consumption patterns to enable better product targeting\" [N].",
        "TMF620: first status is In Study; \"conception approved\" moves it to In Design [S].",
    ],
    ["Idea scoring / business-case screens and who signs the conception approval."],
)

stage(
    "4.2 Stage 2: Technical product definition",
    "In Design [S]",
    "Define what the product is technically, so the network can deliver it automatically later.",
    "Product Specialist / network & IT (owner), Product Manager (consulted)",
    "Approved concept, network capabilities, existing service specifications",
    [
        "Product Specialist opens the Technical Product Catalog.",
        "Selects an existing Product Specification (or a template) such as \"FTTH Access\" and sets its "
        "characteristics (speed 200 Mbps, technology FTTH).",
        "Links the service specification and the provisioning template that will activate the service.",
        "Defines decomposition rules: which technical tasks an order for this product creates (e.g. activate "
        "port, configure ONT, activate voice line).",
        "Saves; the specification is now available for commercial offers.",
    ],
    "Product specification complete and linked to service and provisioning definitions.",
    [
        "Technical Product Catalog holds \"product/service specifications, provisioning templates, "
        "decomposition rules\" [N] (Figure 2).",
        "Product Specification is a TM Forum-certified entity in Netcracker's data model [S][N].",
        "Template-based configuration \"using out-of-the-box product components\" [N].",
    ],
)

stage(
    "4.3 Stage 3: Commercial offer definition",
    "In Design [S]",
    "Package the technical product into something marketing can sell.",
    "Product Manager / Marketing (owner)",
    "Product specifications from Stage 2, concept from Stage 1",
    [
        "Product Manager opens the Product Offering Catalog and creates a new Product Offering from a template.",
        "Links it to the \"FTTH Access 200\" product specification.",
        "Builds a bundle: internet 200 Mbps + landline + Wi-Fi router.",
        "Adds partner content if needed (e.g. a streaming add-on) through catalog federation.",
        "Places the offering in the Residential product catalog / category.",
    ],
    "Offering structure (single offer or bundle) complete and linked to its specifications.",
    [
        "Product Offering Catalog holds \"offering details, business rules, prices and discounts\" [N].",
        "\"Centralizing the management of products and bundles\" [N].",
        "Federation integrates \"third-party partner catalogs into a single product management interface\" [N].",
    ],
)

stage(
    "4.4 Stage 4: Pricing",
    "In Design [S]",
    "Attach all prices and price-related terms to the offering.",
    "Product Manager (owner), Finance / Pricing (approver)",
    "Offering from Stage 3, pricing strategy, tax rules",
    [
        "Product Manager opens the offering's pricing section.",
        "Adds a recurring monthly fee and a one-time installation fee.",
        "Adds a discount (e.g. a launch discount for the first 3 months) and any fees.",
        "Applies the correct taxes and, if relevant, loyalty points.",
        "Sets different prices per market or customer category if needed.",
    ],
    "Complete price set attached to the offering.",
    [
        "Catalog covers \"pricing, discounts, fees, loyalty programs, taxes\" [N].",
        "Pricing \"based on market, customer category, distribution channel\" (CPQ whitepaper) [N].",
        "Product Offering Price is a TM Forum-certified entity [S][N].",
    ],
    ["The pricing screens themselves and whether price changes need a separate approval."],
)

stage(
    "4.5 Stage 5: Business rules and eligibility",
    "In Design [S]",
    "Control who can buy the offer, where, through which channel, and with what.",
    "Product Manager (owner), Product Specialist (consulted)",
    "Offering and prices; coverage areas; channel strategy",
    [
        "Product Manager opens the offering's rules.",
        "Sets eligibility: market = fiber-covered areas, distribution channel = shops / call center / app, "
        "customer category = residential.",
        "Adds dependencies: the bundle requires a compatible router.",
        "Adds constraints: cannot be combined with ADSL on the same line.",
    ],
    "Rules defined, so the downstream sales tools will only offer the product to eligible customers.",
    [
        "Validation against \"business rules, such as eligibility rules, configured in the Product catalog\" "
        "to check availability \"for the market, distribution channel and customer category\" [N].",
        "Rules cover \"product specifications, dependencies, constraints and business rules\" [N].",
        "Figure 3 shows these eligibility fields on a real Netcracker screen.",
    ],
)

image("03_netcracker_real_screen_create_quote.png",
      "Figure 3. Real Netcracker screen (Create Quote). The fields Customer Category, Distribution Channel "
      "and Market are the same eligibility dimensions defined in the catalog. Source: Netcracker CPQ "
      "whitepaper (2021), Figure 3.", width_cm=13)

stage(
    "4.6 Stage 6: Approval gates",
    "In Design to In Test [S]",
    "Formally approve the design before testing starts.",
    "Product Manager (submits); Marketing head, Finance, Legal / Regulatory (NTRA if required), "
    "IT/Network (approvers) [P]",
    "Complete design from Stages 2 to 5",
    [
        "Product Manager submits the offering for design approval.",
        "Approvers review: commercial fit (Marketing), prices and margin (Finance), terms and regulation "
        "(Legal), technical feasibility (IT/Network).",
        "Each approver approves or rejects with comments.",
        "When approved, the status changes to In Test; if rejected, the product stays In Design for rework.",
    ],
    "Design approved and status In Test.",
    [
        "TMF620 defines the gates: \"conception approved\" to In Design and \"design approved\" to In Test [S].",
        "Netcracker's sales tools use approval statuses (\"Approved\" visible in Figure 3) and "
        "\"streamline approval workflows\" [N].",
        "PLM provides \"automating end-to-end lifecycle management\" [N].",
    ],
    ["The product-approval workflow screens, approver configuration, and sequential vs parallel approval."],
)

stage(
    "4.7 Stage 7: Versioning during design and change",
    "Any status [S]",
    "Keep a controlled history of every change to the product.",
    "Product Manager, Catalog Administrator [P]",
    "An existing product or offering that needs to change",
    [
        "Instead of editing a tested or approved product directly, the user creates a new version.",
        "The new version goes back through design, approval and test.",
        "Previous versions remain visible for comparison and audit.",
    ],
    "New version approved and tested; old version kept in history.",
    [
        "PLM \"fully supports product versioning, upgrades\" [N].",
        "TMF620 catalog entities carry a version and a lifecycle status [S].",
    ],
    ["Version comparison screens and whether parallel versions are allowed."],
)

stage(
    "4.8 Stage 8: Testing and readiness",
    "In Test to Active (or Rejected) [S]",
    "Prove the product works end to end before any customer can buy it.",
    "QA / Test team (owner), IT, Product Manager",
    "Approved design",
    [
        "The approved product is deployed to a test environment.",
        "Test team runs test orders to check that pricing, rules, provisioning templates and decomposition "
        "rules work.",
        "If tests pass, the product moves to Active, ready for a test launch.",
        "If tests fail, the product is Rejected (final), or reworked as a new version.",
        "Optional test launch to a limited pilot group before mass production.",
    ],
    "Status Active: validated and tested, ready to be published to CRM and channels. "
    "This is the hand-off point where this document ends.",
    [
        "PLM covers \"product testing, test launch and subsequent mass production\" [N].",
        "TMF620: test OK gives Active; test KO gives Rejected, a final status [S].",
    ],
)

# ------------------------------------------------------------------ 5
h1("5. Summary: the journey on one page")
table([
    ["#", "Stage", "Status", "Netcracker module", "Owner [P]", "Output"],
    ["1", "Ideation", "In Study", "PLM + analytics", "Marketing / PM", "Approved concept"],
    ["2", "Technical definition", "In Design", "Technical Product Catalog", "Product Specialist", "Product specification"],
    ["3", "Commercial definition", "In Design", "Product Offering Catalog", "Product Manager", "Offering / bundle"],
    ["4", "Pricing", "In Design", "Catalog pricing", "PM + Finance", "Price set"],
    ["5", "Rules & eligibility", "In Design", "Business rules", "Product Manager", "Eligibility rules"],
    ["6", "Approval", "In Design to In Test", "PLM", "Approvers", "Approved design"],
    ["7", "Versioning", "Any", "PLM", "PM / Admin", "Version history"],
    ["8", "Testing & readiness", "In Test to Active", "PLM", "QA / IT", "Launch-ready product"],
], widths=[0.7, 3.2, 2.8, 3.6, 2.9, 3.4])

# ------------------------------------------------------------------ 6
h1("6. Evidence that this works at real operators")
table([
    ["Operator", "What was deployed", "Published result", "Source type"],
    ["Zain KSA (Saudi Arabia)", "Full BSS/OSS including Product Catalog", "50% decrease in product development time; done in under 3 years", "Netcracker case study with operator quotes"],
    ["T-Mobile Netherlands", "Product Catalog + Product Lifecycle Management specifically", "Manage products \"from development to production to retirement\"; fewer incorrect offers", "Netcracker press release with operator quote"],
    ["e& UAE", "Full BSS including product catalog management", "\"Groundwork for catalog-driven offers\"", "Netcracker press release with operator quote"],
    ["Telecom Egypt", "Netcracker BSS/OSS (2013, live 2014)", "None published", "Third-party database only; unconfirmed"],
], widths=[3.4, 4.4, 5.0, 3.8])
image("04_zain_ksa_case_study_page1.png",
      "Figure 4. Zain KSA case study: 50% decrease in product development time. Source: Netcracker.", width_cm=12)

# ------------------------------------------------------------------ 7
h1("7. Open questions for a Netcracker demo or our IT team")
bullets([
    "Does Telecom Egypt still run Netcracker BSS/OSS (2013 contract), and does it include Product Catalog / PLM?",
    "How are ideas captured and scored before In Study? Is there a business-case approval screen?",
    "How is the product-approval workflow configured: approvers per stage, sequential or parallel, delegation?",
    "What roles and permissions exist for catalog users (view, create, edit, approve, publish)?",
    "Can we add our own lifecycle statuses (e.g. a regulatory approval step)?",
    "How are versions compared, and can two versions be in design at the same time?",
    "Which test environments are supported, and how does a product move from test to production?",
])

# ------------------------------------------------------------------ 8
h1("8. References")
refs = [
    "Netcracker, Product Management (Product Catalog + PLM): "
    "https://www.netcracker.com/portfolio/products/digital-bss/product-management",
    "Netcracker, Commerce Management: "
    "https://netcracker.com/portfolio/products/digital-commerce-monetization/commerce-management",
    "Netcracker, Configure-Price-Quote whitepaper (2021), Figures 2 and 3: "
    "https://www.netcracker.com/assets/portfolio/solutions/intelligent-sales-automation/netcracker_cpq-whitepaper.pdf",
    "Netcracker, Zain KSA case study: https://netcracker.com/documents/pdf/netcracker_zain_case-study_260918.pdf",
    "Netcracker, T-Mobile Netherlands press release: "
    "https://www.netcracker.com/news/press-releases/t-mobile-netherlands-optimizes-end-to-end-product-management-with-netcracker-bss/",
    "Netcracker, e& UAE partnership (Oct 2023): "
    "https://netcracker.com/news/press-releases/etisalat-by-e-and-partners-with-netcracker-on-the-largest-full-stack-bss-transformation-project-in-the-middle-east",
    "TM Forum, Netcracker SID certification report (Figure 1): "
    "https://www.tmforum.org/wp-content/uploads/2015/03/Netcrakcer-Conformance-Cert-NetCracker-TOMS-Suite-Release-8.2-SID-Report.pdf",
    "TM Forum, Open API conformance list (Netcracker certified TMF620, June 2024): "
    "https://www.tmforum.org/conformance-certification/open-api-conformance/",
    "TM Forum, TMF620 Product Catalog Management API (lifecycle statuses): "
    "https://www.tmforum.org/open-digital-architecture/open-apis/product-catalog-management-api-TMF620/v5.0",
    "AppsRunTheWorld, Telecom Egypt record (unconfirmed): "
    "https://www.appsruntheworld.com/customers-database/customers/view/telecom-egypt-egypt",
]
numbered(refs)

doc.save(OUT)
print("saved", OUT.name)
