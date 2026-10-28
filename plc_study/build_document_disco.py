"""Builds the PLC study document (DISCOBOLE / ODACAT, open source).

Run:  python3 build_document_disco.py
"""

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

HERE = Path(__file__).parent
IMG = HERE / "images" / "disco"
OUT = HERE / "PLC_DISCOBOLE_Study.docx"

ORANGE = RGBColor(0xE8, 0x77, 0x22)
GREY = RGBColor(0x5F, 0x5A, 0x6B)
HDR = "E87722"

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
    st.font.color.rgb = ORANGE


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
        r.font.color.rgb = ORANGE
        p.add_run(it)


def table(rows, widths=None, header_fill=HDR):
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
            if r < len(rows) - 1 and len(rows) <= 12:
                cell.paragraphs[0].paragraph_format.keep_with_next = True
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
    shade(cell, "FDEBD8")
    p = cell.paragraphs[0]
    r = p.add_run(title + "  ")
    r.bold = True
    r.font.color.rgb = ORANGE
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


def stage(title, status, purpose, who, inputs, steps, outputs, evidence):
    h2(title)
    table([
        ["Item", "Detail"],
        ["Lifecycle status", status],
        ["Purpose", purpose],
        ["Who", who],
        ["Inputs", inputs],
        ["Exit criteria / output", outputs],
    ], widths=[4.2, 12.4])
    h3("What you do in the system")
    numbered(steps)
    h3("Source")
    bullets(evidence)


# ------------------------------------------------------------------ cover
t = doc.add_paragraph()
r = t.add_run("PLC STUDY  |  OPEN SOURCE")
r.bold = True
r.font.color.rgb = ORANGE
r.font.size = Pt(11)
title = doc.add_paragraph()
r = title.add_run("Product Lifecycle Management in DISCOBOLE:\nfrom Idea to CRM")
r.bold = True
r.font.size = Pt(22)
para("System studied: DISCOBOLE Product Catalog (ODACAT), the open-source TM Forum ODA "
     "Product Catalog component. DISCOBOLE stands for Digital & Innovative OpenSource "
     "Platform for Core Commerce Management. Hosted by OW2. Source code and UI documentation "
     "are public.", color=GREY)
para("Prepared by: [Your name]    |    Version: Draft 1 for review    |    October 2026",
     color=GREY)

note("How to read this document.",
     "Every statement is tagged: [D] = DISCOBOLE public documentation or UI; "
     "[S] = TM Forum standard (TMF620, TMF701, TMF679, TMF622, TMF723); "
     "[M] = Microsoft documentation; [E] = illustrative example. "
     "Screenshots are the real ODACAT Product Catalog screens from the official docs, "
     "not reconstructions.")

h1("Key takeaways")
bullets([
    ("One source of truth. ", "Product specifications, offerings, prices and policy rules "
     "are defined once in ODACAT. Order capture, the configurator and billing read them "
     "instead of keeping their own copies [D]."),
    ("A governed journey. ", "Each catalog item moves through inStudy, inDesign, inTest, "
     "active and launched. Direct edits stop at active: a new version is required [D][S]."),
    ("Business users configure in a UI. ", "The Catalog Admin Portal walks the user through "
     "guided cards (Create, Next, Update status). There is no development step to add an "
     "offer [D]."),
    ("Rules travel with the product. ", "Policy rules (event-condition-action) and "
     "eligibility (channel, market segment, sellable flag) are stored on the offering and "
     "evaluated by the Product Configurator at quote time [D]."),
    ("Standard integration. ", "The catalog is exposed as TMF620 (read), TMF701 (create / "
     "lifecycle), TMF723 (policy), TMF679 / TMF760 (qualification and configuration) and "
     "TMF622 (ordering) [D][S]."),
    ("Inspectable. ", "Architecture diagrams, process-flow charts, the data model and every "
     "admin screen are published. The code is on OW2 GitLab [D]."),
])

# ------------------------------------------------------------------ 1
doc.add_page_break()
h1("1. PLC and the TM Forum standard in brief")
para("A Product Lifecycle Management (PLC) system is where an operator turns an idea into a "
     "sellable product, controls every change to it, and publishes it to the systems that "
     "sell, deliver and bill it.")
para("TM Forum defines this lifecycle in TMF620. DISCOBOLE implements those statuses on "
     "Product Offerings and Product Specifications, with a dedicated Lifecycle Management "
     "process (TMF701) that returns only the statuses allowed from the current one [D][S].")
table([
    ["Status", "Meaning in DISCOBOLE", "Gate / action"],
    ["inStudy", "Conception of the offering", "Creation started"],
    ["inDesign", "The offering is being designed", "Conception approved"],
    ["inTest", "The offering is being tested", "Design completed"],
    ["rejected", "Definition abandoned (final)", "Abandoned / cancelled"],
    ["active", "Ready to be used and delivered", "Test is OK"],
    ["launched", "Start of marketing; on sale", "Start of marketing"],
    ["retired / unavailable / obsolete", "No new sales / no customers left", "End of sale"],
], widths=[3.6, 7.4, 5.6])
image("lifecycle-product-specification-product-offering.png",
      "Figure 1. DISCOBOLE Product Offering lifecycle (official process diagram). "
      "Source: DISCOBOLE Product Catalog overall architecture [D].", width_cm=12)
para("This document covers everything up to launched: the stages inside the catalog, and "
     "the hand-over to order capture / CRM.", italic=True)

# ------------------------------------------------------------------ 2
h1("2. DISCOBOLE Product Catalog at a glance")
para("DISCOBOLE is a suite of ODA Core Commerce components covering Order-to-Bill. "
     "The Product Catalog component is nicknamed ODACAT (ODA + CATalogue). It organises "
     "the products and offerings that can be commercialised, including lifecycle, "
     "relationships, packaging rules and pricing conditions [D].")
table([
    ["Component", "Role in the suite"],
    ["Product Catalog (ODACAT)", "Define specifications, offerings, prices, policy rules, lifecycle"],
    ["Product Configurator (CONF)", "Check eligibility, allowed combinations and discounts at quote time"],
    ["Order Management (OM)", "Order capture and follow-up; stores product orders"],
    ["Order Orchestration (COOD)", "Orchestrate delivery of the product order"],
    ["Product Inventory (CPIB)", "Products already assigned to customers"],
    ["User Roles and Permissions", "Who can view, create, approve, publish (TMF672)"],
    ["Catalog Admin Portal", "The UI used in the walkthroughs below"],
], widths=[5.4, 11.2])

image("suite-architecture.png",
      "Figure 2. DISCOBOLE suite architecture. Product Catalog sits at the centre and "
      "exposes TMF620 / TMF701 / TMF723. Order Capture and the Configurator consume it. "
      "Source: DISCOBOLE overall architecture [D].", width_cm=16.5)

h2("2.1 Bottom-up catalog model")
para("New offers are configured bottom-up. Customer Facing Service (CFS) specifications "
     "and Stock Item Types are mastered outside the catalog (service catalog / stock) and "
     "replicated into ODACAT. The catalog administrator then builds Product Specifications, "
     "then Product Offerings and prices, then policy rules [D].")
image("frontend-product-catalog-ui-bottom-up-approach.png",
      "Figure 3. Bottom-up approach used in the Catalog Admin Portal [D].", width_cm=12)
table([
    ["Entity", "Plain meaning", "Example [E]"],
    ["CFS Specification", "Intangible service template, imported from the service catalog",
     "Dedicated Internet Access"],
    ["Stock Item Type", "Tangible article template, imported from stock", "Managed router"],
    ["Product Specification", "What the operator sells at a functional level, restricted from CFS or stock",
     "Business Internet 100 Mbps"],
    ["Atomic Product Offering", "A sellable item on top of one specification",
     "Internet 100 Mbps, Static IP, Router"],
    ["Bundle Product Offering", "Several atomic or bundle offerings with cardinalities",
     "Internet + IP + router + SLA"],
    ["Contract Product Offering", "What the customer subscribes to",
     "Business Internet 100, 24-month contract"],
    ["Product Offering Price", "Charge, alteration (discount), tax or installment",
     "Monthly fee + installation fee"],
    ["Policy Rule", "Event-Condition-Action rule evaluated at configuration time",
     "24-month term unlocks the backup-link add-on"],
    ["Category", "Logical grouping of offerings", "Enterprise connectivity"],
], widths=[4.4, 6.6, 5.6])
image("data-model-tmf620.png",
      "Figure 4. DISCOBOLE TMF620 data model: commercial, functional and technical views [D].",
      width_cm=13)

h2("2.2 How work is executed: process flows, not free editing")
para("Creation, modification, clone and lifecycle change are not free CRUD on a form. "
     "Each is a TMF701 Process Flow: a sequence of Task Flows that the UI renders as "
     "cards with Next / Create / Update status. The back end owns the process; any "
     "front end (admin portal, or a CRM) is guided through the same API [D].")
image("e2e.png",
      "Figure 5. End-to-end process to create a new offering: specification, offering, "
      "price, policy, relationships, validate, change lifecycle to launched [D].",
      width_cm=16.5)

# ------------------------------------------------------------------ 3
h1("3. Using the system: from idea to launch-ready")
para("The running example is \"Business Internet 100\": 100 Mbps dedicated internet, a "
     "static IP block, a managed router and an availability SLA, sold on 12, 24 or "
     "36-month contracts [E]. Screen names and clicks below match the ODACAT UI [D].")

image("ui-odacat-login.png",
      "Figure 6. ODACAT login. Access is role-based (Keycloak / TMF672) [D].", width_cm=14)
image("ui-odacat-homepage.png",
      "Figure 7. Homepage after login. Left: Catalog Entity menu (Offering, Price, Policy "
      "Rule, Category, Product Specification, CFS, Stock Item, Lifecycle Management). "
      "Main: Product Offering dashboard with Type, Status, Version and actions "
      "(JSON, hierarchy, lifecycle, edit) [D].", width_cm=16)

stage(
    "3.1 Stage 1: Ideation and catalog start",
    "inStudy [D][S]",
    "Register that a new product is being conceived. In DISCOBOLE this is the start of "
    "the create-offering / create-specification process, not a separate scoring module.",
    "Catalog administrator (product manager role) [D]",
    "Market need, existing CFS / stock items, existing specifications that can be reused",
    [
        "Sign in to the Catalog Admin Portal (Figure 6).",
        "Search the Product Specification and Product Offering dashboards for something "
        "that can be reused or cloned (clone creates version 0.1 in inTest) [D].",
        "If nothing fits, start a new Product Specification (Stage 2) from an existing "
        "CFS or Stock Item Type. Those support entities cannot be created in ODACAT; "
        "they arrive by event from the service catalog / stock [D].",
        "The new specification is created with lifecycle status inTest [D].",
    ],
    "A specification (or a cloned offering) exists and can be designed.",
    [
        "Clone sets version 0.1 and status inTest, regardless of the source status [D].",
        "CFS and Stock Item Types are read-only replicas in ODACAT [D].",
        "TMF620 first status is In Study; DISCOBOLE's published state machine starts "
        "create in inStudy / inDesign / inTest depending on the entity [D][S].",
    ],
)

stage(
    "3.2 Stage 2: Technical product definition (Product Specification)",
    "inTest (on create) then active / launched [D]",
    "Define what the operator sells at a functional level, as a restriction of a CFS "
    "or a Stock Item Type.",
    "Catalog administrator [D]",
    "Imported CFS Specification or Stock Item Type",
    [
        "Open Product Specification and click Create (Figure 8).",
        "Select support entity type: CFS Specification (intangible, e.g. dedicated "
        "internet) or Stock Item Type (tangible, e.g. managed router) [D].",
        "Pick the CFS / stock item, then Define Identity Data: name, description, "
        "brand, validity dates [D].",
        "Restrict Associated Characteristics from the CFS (e.g. bandwidth = 100 Mbps, "
        "access type = fiber) [D].",
        "For CFS-based specs, Associate Relationships (reliesOn). Example: a backup-link "
        "specification reliesOn the internet specification [D].",
        "Click Create. The specification is stored with status inTest [D].",
    ],
    "Product specification complete, characteristics restricted, relationships set.",
    [
        "Official create-specification card flow (Figure 9) and process diagram "
        "(Figure 10) [D].",
        "Dashboard lists Type (CFSSpec / StockItemType), Status, Version [D].",
    ],
)
image("ui-product-specification-dashboard.png",
      "Figure 8. Product Specification dashboard: search, filter by type and status, "
      "Create / Export, lifecycle and edit actions [D].", width_cm=16)
image("ui-product-specification-create-cards.png",
      "Figure 9. Cards the UI shows when creating a Product Specification from a CFS "
      "versus from a Stock Item Type [D].", width_cm=13)
image("create-product-specification.png",
      "Figure 10. Official process-flow diagram for Create Product Specification [D].",
      width_cm=16)
image("crop-ui-product-specification-details-view.png",
      "Figure 11. Product Specification details in the UI (identity, characteristics, "
      "relationships) [D].", width_cm=16)

stage(
    "3.3 Stage 3: Commercial offer (Product Offering)",
    "inTest on create [D]",
    "Package the specification into something sales can sell: atomic, bundle or contract.",
    "Catalog administrator [D]",
    "Product Specification from Stage 2",
    [
        "Open Product Offering and click Create.",
        "Select type: Atomic (one specification), Bundle (several offerings with "
        "min/max cardinalities) or Contract (what the customer subscribes to) [D].",
        "For an atomic offering, select the Product Specification "
        "(e.g. Business Internet 100 Mbps) [D].",
        "Provide Description: name, description, brand, Installable / Visible / Sellable "
        "flags, validity dates, Channel, Market Segment, Commitment Term "
        "(12 / 24 / 36 months) [D].",
        "Associate Category (Enterprise connectivity), define characteristics and "
        "eligible values, associate prices and policy rules [D].",
        "Click Create. Status is inTest. The offering is not yet visible to sales "
        "because Sellable is typically off until launched [D].",
    ],
    "Offering structure complete: atomic items, bundle and contract.",
    [
        "Create Product Offering screen (Figure 12) [D].",
        "Three offering types and hierarchy graph (Figure 13) [D].",
    ],
)
image("ui-ho-po1-create1.png",
      "Figure 12. Create Product Offering: type (Atomic / Bundle / Contract), linked "
      "specification, name, Visible / Sellable, Channel, Market Segment, Commitment Term [D].",
      width_cm=14)
image("ui-product-offering-dashboard.png",
      "Figure 13. Product Offering dashboard. Type = Contract / Atomic / Bundle; "
      "Status = inTest / active / launched; Version; actions include lifecycle change "
      "and contract hierarchy [D].", width_cm=16)
image("ui-offering-hierarchy-diagram.png",
      "Figure 14. Offering hierarchy graph for a contract (\"Mobile Package Relax\" in "
      "the official example): contract, bundles and atomic items with lifecycle status "
      "on each node [D].", width_cm=15)

stage(
    "3.4 Stage 4: Pricing (Product Offering Price)",
    "Attached to an offering in inTest / active [D]",
    "Attach charges, discounts, taxes and installment plans to the offering.",
    "Catalog administrator, with finance values already agreed [D]",
    "Offering from Stage 3, pricing strategy, tax codes",
    [
        "Open Product Offering Price and click Create.",
        "Select price type: Charge, Alteration (discount), Tax or InstallmentCharge [D].",
        "For a recurring charge: name, price type = Recurring Charge, value excluding "
        "tax, currency, period (1 Month), tax code, proration, charge cycle [D].",
        "Optionally link an alteration (e.g. 10% for 3 months) with relationship "
        "\"Altered by\" [D].",
        "Create a one-time installation charge the same way, Price Type = Non Recurring [D].",
        "Associate the prices to the offering (or to a commitment term) [D].",
    ],
    "Complete price set attached to the offering.",
    [
        "Price dashboard (Figure 15) and Create Charge screen (Figure 16) [D].",
        "TMF620 Product Offering Price entity [S].",
    ],
)
image("ui-product-offering-price-dashboard.png",
      "Figure 15. Product Offering Price dashboard: Charge, Alteration, Tax; recurring "
      "and non-recurring; percentage or amount [D].", width_cm=16)
image("ui-product-offering-price-create-popc-example1.png",
      "Figure 16. Create Product Offering Price, Charge type: recurring value, currency, "
      "period, tax, relationship to an alteration [D].", width_cm=14)

stage(
    "3.5 Stage 5: Policy rules and eligibility",
    "Defined on the offering; evaluated at configuration time [D]",
    "Control which combinations, prices and terms apply, and who can buy the offer.",
    "Catalog administrator [D]",
    "Offering and prices; channel and market-segment strategy",
    [
        "On the offering, set Channel and Market Segment (greyed until Sellable is on) "
        "and the Sellable / Visible flags [D].",
        "Open Policy Rule and click Create.",
        "Define Event (what triggers the rule, e.g. offering selected), Condition "
        "(logical test, e.g. commitment term = 24 months) and Action (apply a price, "
        "include an add-on, block a combination) [D].",
        "Associate the policy rule to the Product Offering [D].",
        "At quote time the Product Configurator evaluates these rules (TMF679 / TMF760) [D].",
    ],
    "Eligibility and configuration rules stored with the offering.",
    [
        "Policy Rule dashboard (Figure 17). Rules are Event-Condition-Action (TMF723) [D][S].",
        "Policy rules are not cloned with an offering; they must be recreated on the clone [D].",
    ],
)
image("ui-policy-rule-dashboard.png",
      "Figure 17. Policy Rule dashboard. Example names encode term and volume "
      "(A9 512GB_12 Month_SGB), type = Pricing, status = active [D].", width_cm=16)

stage(
    "3.6 Stage 6: Lifecycle change (the approval gate in the UI)",
    "inDesign / inTest to active / launched [D][S]",
    "Move the offering only to a status that the lifecycle service allows.",
    "Catalog administrator with lifecycle permission; User Roles component (TMF672) [D]",
    "Complete design from Stages 2 to 5",
    [
        "On the Product Offering dashboard, click the lifecycle icon in the Actions column [D].",
        "A modal shows the current status and the allowed target statuses returned by "
        "the Lifecycle Management process (TMF701). Example in the docs: change "
        "\"FFTH Demo Contract\" from inTest to active [D].",
        "Select the target status and Save changes. The dashboard refreshes [D].",
        "The same action exists on the offering details page [D].",
        "The process is: Provide Entity ID and Entity Type, system returns possible "
        "statuses, user provides the new status, system updates [D].",
    ],
    "Offering in active (ready) then launched (on sale).",
    [
        "Official lifecycle process diagram (Figure 18) and UI description [D].",
        "Version rules: before active, direct edit is allowed; from active onward a "
        "new version is required; rejected and obsolete are not editable [D].",
    ],
)
image("lifecycle.png",
      "Figure 18. Lifecycle Management process: the back end returns the possible next "
      "statuses for the entity type (Contract / Bundle / Atomic / Product Specification) [D].",
      width_cm=16)

stage(
    "3.7 Stage 7: Versioning and clone",
    "Any status [D][S]",
    "Keep a controlled history. Live entities are not edited in place.",
    "Catalog administrator [D]",
    "An existing specification or offering",
    [
        "From active or launched, create a new version (minor 1.0 to 1.1 for compatible "
        "changes; major 2.1 to 3.0 starts again in inStudy) [D].",
        "Or Clone: limited clone (the entity only) or full hierarchy (down to atomic) [D].",
        "The clone is version 0.1, status inTest, name suffixed _CLONE [D].",
        "Take the clone through design, policy and lifecycle again. Previous versions "
        "stay in the dashboard for audit [D].",
    ],
    "New version in test; old version retained, later retired via validFor dates.",
    [
        "Lifecycle Management service: \"Once Active or beyond, changes require a new "
        "version.\" Minor: only the latest stays active/launched, old one is retired. "
        "Major: starts in inStudy [D].",
    ],
)

stage(
    "3.8 Stage 8: Testing and hand-over to sales",
    "inTest to active to launched [D][S]",
    "Prove the offering configures and prices correctly, then make it sellable.",
    "Catalog administrator, test team [D]",
    "Offering in inTest",
    [
        "Filter the dashboard on status = inTest [D].",
        "Open the offering, check hierarchy, prices and policy rules.",
        "Use the Product Configurator (or a test quote in Order Capture) to confirm "
        "eligibility, combinations and discounts [D].",
        "If tests pass, change lifecycle inTest to active, then active to launched, "
        "and set Sellable = on so Channel and Market Segment apply [D].",
        "If tests fail, change status to rejected (final) or clone and rework [D].",
    ],
    "Status launched and Sellable: the offering is visible to order capture / CRM.",
    [
        "Dashboard default list is inTest / active / launched only [D].",
        "TMF620: test OK gives Active; start of marketing gives Launched [S].",
    ],
)

h2("3.9 The journey on one page")
table([
    ["#", "Stage", "Status", "ODACAT module / screen", "Output"],
    ["1", "Ideation / reuse", "inStudy", "Dashboards, Clone", "Spec or clone to design"],
    ["2", "Technical definition", "inTest", "Product Specification", "Restricted CFS / stock spec"],
    ["3", "Commercial offer", "inTest", "Product Offering (A/B/C)", "Atomic, bundle, contract"],
    ["4", "Pricing", "inTest", "Product Offering Price", "Charge, discount, tax"],
    ["5", "Rules", "inTest", "Policy Rule + Channel / Segment", "ECA rules, eligibility"],
    ["6", "Lifecycle change", "to active", "Lifecycle Management", "Allowed next status applied"],
    ["7", "Versioning", "Any", "Clone / new version", "History retained"],
    ["8", "Test and launch", "launched", "Configurator + Sellable flag", "Visible to CRM / order capture"],
], widths=[0.7, 3.2, 2.2, 5.4, 5.1])

# ------------------------------------------------------------------ 4
h1("4. Integration with CRM: how the catalog hands the product over")

h2("4.1 The principle: CRM consumes, it does not re-create")
para("In DISCOBOLE the catalog is the single source of truth. Order Capture, the Shop UI "
     "and any external CRM do not create their own product records. They read TMF620, "
     "ask the Configurator whether a given customer can buy a given offer, then place a "
     "TMF622 product order that references the offering id and version [D].")
para("The Catalog Admin Portal is one front end on TMF701. The same process-flow API is "
     "the documented way \"to implement a centralized order capture process, guiding any "
     "front-end seamlessly through the process flow interface TMF701\" [D].")

h2("4.2 What gets published")
table([
    ["Item", "Why the CRM / order capture needs it", "API"],
    ["Product Offering (atomic / bundle / contract)", "What the salesperson selects", "TMF620"],
    ["Product Offering Price and alterations", "Quote uses the approved price", "TMF620"],
    ["Channel, Market Segment, Sellable / Visible", "Only eligible offers appear", "TMF620"],
    ["Policy rules (ECA)", "Combinations, discounts, terms", "TMF723 + Configurator"],
    ["Validity dates (validFor)", "Appear and disappear on time", "TMF620"],
    ["Lifecycle status and version", "Only launched + sellable is sold", "TMF620 / TMF701"],
    ["Hierarchy / relationships", "Bundles, reliesOn, requires (e.g. shipment)", "TMF620"],
], widths=[5.6, 7.0, 3.8])

h2("4.3 When it gets published")
bullets([
    "The TMF620 query API can read any status, but sales front ends are expected to "
    "filter on launched (and Sellable = true) [D][S].",
    "inStudy, inDesign, inTest, rejected stay in the catalog UI for designers [D].",
    "Start of marketing (active to launched) is a lifecycle change that can be timed "
    "with validFor start date [D][S].",
    "Kafka events notify Order Management, Product Inventory and the Configurator when "
    "catalog entities change (Figure 2, green Event bars) [D].",
])

h2("4.4 How it is published: the Open APIs")
table([
    ["API", "Direction", "What it does"],
    ["TMF620 Product Catalog", "Catalog to CRM / OM / Configurator",
     "Read catalogs, offerings, prices, specifications, categories"],
    ["TMF701 Process Flow", "Any front end to catalog",
     "Create, modify, clone, change lifecycle (the admin UI uses this)"],
    ["TMF723 Policy Management", "Catalog to Configurator",
     "Policy rules evaluated at configuration time"],
    ["TMF679 Product Offering Qualification", "CRM / OM to Configurator",
     "\"Can this customer get this offer here?\""],
    ["TMF760 Product Configuration", "CRM / OM to Configurator",
     "Allowed combinations, characteristics and discounts"],
    ["TMF622 Product Ordering", "CRM / OM",
     "Order carries offering id and version"],
    ["Kafka events", "Catalog to the rest of the suite",
     "Create / change / status change without polling"],
], widths=[4.6, 5.0, 7.0])

h2("4.5 What the salesperson sees")
para("In the DISCOBOLE suite the salesperson works in Order Capture / Shop UI, not in "
     "ODACAT. The catalog still drives that screen: Channel, Market Segment and Sellable "
     "on the offering (Figure 12) are the same dimensions the quote uses. The Configurator "
     "applies policy rules before the quote can be submitted [D].")
para("The offering hierarchy graph (Figure 14) is the catalog-side view of what the "
     "salesperson will be offered as a contract with its bundles and atomic items.")

h2("4.6 When a product changes after launch")
numbered([
    "The catalog administrator clones the launched offering, or creates a new version [D].",
    "The clone is inTest (version 0.1). It goes through prices, policy, test and lifecycle "
    "again [D].",
    "When the new version becomes launched, a catalog event is published on Kafka [D].",
    "New quotes and orders use the new version. Existing customers keep the version on "
    "their product order / inventory, because the order stored the offering id and "
    "version [D][S].",
    "The old version is retired (validFor end, or lifecycle retired) and disappears from "
    "sales screens [D].",
])

h2("4.7 Example: connecting the catalog to Microsoft Dynamics 365")
para("The same pattern applies to a CRM from another vendor. Dynamics 365 Sales has its "
     "own product entities [M]. There is no ready-made ODACAT connector for Dynamics, so "
     "the link is an integration layer: TMF620 / events on one side, Dataverse Web API on "
     "the other [E].")
table([
    ["Catalog item (DISCOBOLE / TMF620)", "Dynamics 365 Sales entity [M]"],
    ["Atomic Product Offering", "Product"],
    ["Bundle / Contract Product Offering", "Product bundle"],
    ["Category", "Product family"],
    ["Product Offering Price / Alteration", "Price list item / Discount list"],
    ["Related offerings, add-ons, reliesOn", "Product relationships (accessory, cross-sell)"],
    ["Status launched / retired", "Product status Active / Retired"],
    ["Channel, Market Segment", "Price list / territory / customer group"],
], widths=[8.3, 8.3])
para("Dynamics statuses are Draft, Active, Under Revision and Retired [M]. inStudy, "
     "inDesign, inTest and approval stay in ODACAT. Dynamics receives only launched, "
     "sellable offerings.", italic=True)

# ------------------------------------------------------------------ 5
h1("5. Evidence")
table([
    ["Fact", "What it shows", "Source"],
    ["Public Catalog Admin UI, with login, dashboards and create wizards",
     "The walkthroughs are the real screens, not an illustration",
     "DISCOBOLE Product Catalog UI docs"],
    ["Published lifecycle and create-offering process diagrams",
     "The workflow is specified, not inferred",
     "Overall architecture pages"],
    ["TMF620, TMF701, TMF723 implemented; TMF679 / TMF760 / TMF622 in the suite",
     "The catalog speaks the industry APIs a CRM already expects",
     "Suite architecture (Figure 2)"],
    ["OW2 GitLab source (Java 17, Spring Boot, Axon CQRS, Kafka, MongoDB)",
     "The system can be cloned, read and run",
     "gitlab.ow2.org/discobole"],
    ["Built as Orange's ODA Core Commerce catalog, aimed at a catalog-driven solution",
     "It is an operator-originated design, not a generic e-commerce catalog",
     "DISCOBOLE introduction"],
], widths=[5.6, 5.8, 5.2])
para("Unlike a closed vendor, DISCOBOLE does not publish a \"50% faster time to market\" "
     "case study. The evidence here is stronger in a different way: every screen, every "
     "process flow and the source code are inspectable.", italic=True)

h1("6. Points to explore further")
bullets([
    "How lifecycle permissions are bound to TMF672 roles (who may move inTest to active).",
    "Whether extra statuses (e.g. regulatory approval) can be added to the state machine.",
    "How Kafka catalog events are consumed by an external CRM, and retry on failure.",
    "Multi-site and commitment-term pricing for enterprise contracts.",
    "Running ODACAT locally (Java 17, Maven, MongoDB, Kafka, Keycloak) to replay the "
    "walkthrough on a live instance.",
])

h1("7. References")
refs = [
    "DISCOBOLE Product Catalog introduction: "
    "https://discobole.ow2.io/disco-oda-components/disco-catalog/doc/about/introduction/",
    "DISCOBOLE Catalog UI (screens used in this document): "
    "https://discobole.ow2.io/disco-oda-components/disco-catalog/doc/frontend/introduction/",
    "Product Offering UI (create, clone, lifecycle, hierarchy): "
    "https://discobole.ow2.io/disco-oda-components/disco-catalog/doc/frontend/entities/product-offering/",
    "Product Specification UI: "
    "https://discobole.ow2.io/disco-oda-components/disco-catalog/doc/frontend/entities/product-specification/",
    "Product Offering Price UI: "
    "https://discobole.ow2.io/disco-oda-components/disco-catalog/doc/frontend/entities/product-offering-price/",
    "Policy Rule UI: "
    "https://discobole.ow2.io/disco-oda-components/disco-catalog/doc/frontend/entities/policy-rule/",
    "Overall architecture (lifecycle, processes, data model): "
    "https://discobole.ow2.io/disco-oda-components/disco-catalog/doc/architecture/overall-architecture/",
    "Lifecycle Management service: "
    "https://discobole.ow2.io/disco-oda-components/disco-catalog/doc/architecture/detailed-architecture/catalog-product-lifecycle-management/",
    "DISCOBOLE suite architecture: "
    "https://discobole.ow2.io/doc/architecture/overall-architecture/",
    "OW2 GitLab (source): https://gitlab.ow2.org/discobole",
    "TM Forum TMF620 Product Catalog Management API: "
    "https://www.tmforum.org/open-digital-architecture/open-apis/product-catalog-management-api-TMF620/v5.0",
    "Microsoft Learn, Dynamics 365 Sales product catalog: "
    "https://learn.microsoft.com/en-us/dynamics365/sales/set-up-product-catalog-walkthrough",
]
numbered(refs)

doc.save(OUT)
print("saved", OUT.name)
