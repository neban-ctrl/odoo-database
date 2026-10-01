#!/usr/bin/env python3
"""
Build the "Ruskin ES K1 & K2" demo project for Best Contracting Services (BCSI)
inside the Odoo Project app, replicating the client's documents:

  * PROCESS_FLOW.pptx              -> task stages (PH1-PH4), process tasks, role/department tags,
                                      dependencies and milestones
  * WR_23009_01_Weekly_Status_Report.xlsm
                                   -> project info / team, custom task properties that mirror the
                                      Submittal Log, Incoming Docs, Outgoing RFI Log, Material
                                      Procurement and Action Item logs, and a Weekly Status Report
                                      (project update + recurring task)
  * SUB - 07 90 05 - Sealants.pdf  -> submittal sub-task (transmittal data, spec checklist, PDF)
  * RFI 001 - Built-Up Roof Assembly - SEI Response.pdf
                                   -> RFI sub-task (question, engineer response, PDF) and the
                                      follow-up change-order evaluation

Everything is created through the Project app models (project.project, project.task,
project.task.type, project.project.stage, project.tags, project.milestone, project.update),
plus the contacts/users those records need.

Usage:
    export ODOO_URL=https://best-contracting-services-demo-msel.odoo.com
    export ODOO_DB=best-contracting-services-demo-msel      # optional, auto-detected
    export ODOO_LOGIN=you@example.com
    export ODOO_API_KEY=xxxxxxxx
    python3 build_ruskin_project.py            # build / update (idempotent)
    python3 build_ruskin_project.py --dry-run  # validate the data without connecting

Source PDFs / Office files are attached when found in ./source_docs (see README).
"""
import argparse
import base64
import datetime as dt
import json
import os
import sys
import urllib.request
import xmlrpc.client
from pathlib import Path

HERE = Path(__file__).resolve().parent
DOCS_DIR = Path(os.environ.get("SOURCE_DOCS", HERE / "source_docs"))

# Context used for every write: no e-mails to the (real) people named in the documents.
QUIET = {
    "mail_create_nosubscribe": True,
    "mail_auto_subscribe_no_notify": True,
    "mail_notify_force_send": False,
    "no_reset_password": True,
    "tracking_disable": False,
}

# ---------------------------------------------------------------------------
# Source data (taken from the client documents)
# ---------------------------------------------------------------------------
PROJECT_NAME = "Ruskin ES K1 & K2 - Roof Replacement (Kinder Wing)"

CUSTOMER = {
    "name": "Berryessa Union School District",
    "is_company": True,
    "street": "1401 Turlock Lane",
    "city": "San Jose",
    "zip": "95132",
    "state_code": "CA",
    "country_code": "US",
    "comment": "Owner / Customer - Ruskin Elementary School (project address).",
}

COMPANIES = [
    {
        "key": "gc",
        "name": "Strawn Construction",
        "is_company": True,
        "comment": "General Contractor (GC) - Ruskin ES K1 & K2.",
    },
    {
        "key": "sei",
        "name": "Steelhead Engineers (SEI)",
        "is_company": True,
        "comment": "Roofing consultant / Architect of record for the roof replacement. "
                   "SEI Job Number 26006 - Bid Set dated 3/12/26.",
    },
    {
        "key": "sika",
        "name": "Sika Corporation",
        "is_company": True,
        "website": "https://usa.sika.com",
        "comment": "Manufacturer - Sikaflex-1A elastomeric joint sealant (Spec 07 90 05).",
    },
]

CONTACTS = [
    {
        "key": "donny",
        "name": "Donny Durham",
        "parent": "gc",
        "function": "GC Project Contact",
        "email": "ddurham@scmdinc.com",
        "phone": "408-286-1299",
    },
]

# BEST project team (Weekly Status Report "BEST PROJECT TEAM" + documents).
# Odoo requires e-mail logins: they use the reserved, non-deliverable ".example" domain
# (RFC 2606) and users get Odoo-inbox notifications, so the demo never e-mails real people.
LOGIN_DOMAIN = "bcsi-demo.example"
TEAM = [
    {"key": "maria", "name": "Maria Alimagno", "login": "maria.alimagno", "role": "Project Manager",
     "note": "Author of RFI 001. malimagno@bestcontracting.com / 310-328-6969"},
    {"key": "alona", "name": "Alona Bowman", "login": "alona.bowman", "role": "Project Executive"},
    {"key": "wes", "name": "Wes Fleming", "login": "wes.fleming", "role": "Quality Control Manager"},
    {"key": "kris", "name": "Kris Lauengco", "login": "kris.lauengco", "role": "Project Engineer - Roofing"},
    {"key": "raymond", "name": "Raymond Lobato", "login": "raymond.lobato", "role": "Project Engineer - Waterproofing"},
    {"key": "chris", "name": "Chris Wilson", "login": "chris.wilson", "role": "Superintendent"},
    {"key": "joseph", "name": "Joseph Gonzales", "login": "joseph.gonzales", "role": "Foreman"},
    {"key": "kassandra", "name": "Kassandra Solis", "login": "kassandra.solis", "role": "Submittals / Project Administration",
     "note": "Signs submittal transmittals. 310-328-6969"},
]

# Who covers each swim-lane of the process flow.
ROLE_USERS = {
    "Sales": ["alona"],
    "Admin": ["kassandra"],
    "Legal": ["alona"],
    "Operations": ["maria"],
    "Project Manager": ["maria"],
    "Project Engineer": ["kris"],
    "Superintendent": ["chris"],
    "Field Manager": ["chris", "joseph"],
    "QC": ["wes"],
}
OPS_TEAM = ["maria", "kris", "chris", "joseph"]

PROJECT_STAGES = [  # project.project.stage (portfolio kanban)
    ("PH1 - Contract", False),
    ("PH2 - Preconstruction", False),
    ("PH3 - Construction", False),
    ("PH4 - Closeout", False),
    ("Complete / Archived", True),
]
CURRENT_PROJECT_STAGE = "PH2 - Preconstruction"

TASK_STAGES = [  # project.task.type (columns of the project kanban = process-flow phases)
    ("PH1 - Contract", False),
    ("PH2 - Preconstruction", False),
    ("PH3 - Construction", False),
    ("PH4 - Closeout", False),
]

TAG_COLORS = {
    # departments (functional map, slide 1 & 3)
    "Sales": 2, "Admin": 3, "Legal": 9, "Operations": 4,
    # roles (slides 4 & 5)
    "Project Manager": 1, "Project Engineer": 10, "Superintendent": 5, "Field Manager": 6, "QC": 7,
    # log types (weekly status report tabs)
    "Submittal": 11, "RFI": 8, "Incoming Doc": 8, "Procurement": 2, "Action Item": 9,
    "Change Order": 1, "Meeting": 4, "Weekly Report": 10, "Field Activity": 6, "Safety & Quality": 7,
    # spec sections
    "07 51 05 Built-Up Roofing": 3, "07 60 05 Flashing & Sheet Metal": 5, "07 90 05 Sealants": 11,
}

# Task properties = the columns of the weekly status report logs.
PROPERTIES = [
    {"name": "log_ref", "string": "Log No. (SUB / RFI / CO)", "type": "char"},
    {"name": "spec_section", "string": "Spec Section", "type": "char"},
    {"name": "spec_par", "string": "Spec Par. No.", "type": "char"},
    {"name": "drawing_ref", "string": "Drawing Reference", "type": "char"},
    {"name": "doc_type", "string": "Submittal / Doc Type", "type": "selection", "selection": [
        ["product_data", "Product Data"], ["shop_drawings", "Shop Drawings"], ["samples", "Samples"],
        ["warranty", "Warranty"], ["closeout", "Closeout"], ["rfi", "RFI"], ["co", "Change Order"],
        ["other", "Other"]]},
    {"name": "review_status", "string": "Review Status", "type": "selection", "selection": [
        ["not_started", "Not Started"], ["open", "Open"], ["submitted", "Submitted"],
        ["in_progress", "In-Progress"], ["answered", "Answered / Reviewed"],
        ["revise", "Revise & Resubmit"], ["approved", "Approved"], ["closed", "Closed"]]},
    {"name": "date_sent", "string": "Date Sent to GC", "type": "date"},
    {"name": "review_days", "string": "Review Period (Days)", "type": "integer"},
    {"name": "date_due_client", "string": "Due from Client", "type": "date"},
    {"name": "date_response", "string": "Response Received", "type": "date"},
    {"name": "responded_by", "string": "Responded By", "type": "char"},
    {"name": "lead_time_days", "string": "Matl Lead Time (Days)", "type": "integer"},
    {"name": "manufacturer", "string": "Manufacturer", "type": "char"},
    {"name": "po_number", "string": "PO #", "type": "char"},
    {"name": "scope_impact", "string": "Scope / Schedule Impact", "type": "boolean"},
    {"name": "cost_impact", "string": "Cost Impact", "type": "boolean"},
    {"name": "ball_in_court", "string": "Ball in Court (BIC)", "type": "char"},
]

D = dt.date


def d(s):
    return dt.date.fromisoformat(s)


# Milestones (project.milestone)
MILESTONES = [
    ("M1 - Contract Executed", "2026-05-22", True),
    ("M2 - PH1 Meeting / Milestones Set", "2026-05-27", True),
    ("M3 - Submittals Approved & Materials Released", "2026-10-02", False),
    ("M4 - Mobilization / Start of Roofing", "2026-10-05", False),
    ("M5 - Roof Complete K1 & K2 (4-ply BUR + Cap Sheet)", "2026-11-20", False),
    ("M6 - Punchlist Complete", "2026-12-18", False),
    ("M7 - Project Closeout & Files Archived", "2027-01-29", False),
]

# State codes (Odoo 17+)
DONE, PROG, WAIT, APPR, CHG = "1_done", "01_in_progress", "04_waiting_normal", "03_approved", "02_changes_requested"


def html_table(headers, rows):
    th = "".join(f"<th>{h}</th>" for h in headers)
    trs = "".join("<tr>" + "".join(f"<td>{c if c not in (None, '') else '&nbsp;'}</td>" for c in r) + "</tr>" for r in rows)
    return f'<table class="table table-bordered table-sm"><thead><tr>{th}</tr></thead><tbody>{trs}</tbody></table>'


def ul(items):
    return "<ul>" + "".join(f"<li>{i}</li>" for i in items) + "</ul>"


def checklist(items):
    # Odoo html checklist markup
    out = ""
    for text, checked in items:
        cls = "o_checked" if checked else ""
        out += f'<ul class="o_checklist"><li class="{cls}">{text}</li></ul>'
    return out


# ---------------------------------------------------------------------------
# Tasks. key: unique reference; parent: key of parent task; deps: keys this task waits on.
# ---------------------------------------------------------------------------
def build_tasks():
    T = []

    def add(key, name, stage, tags, users, start, end, state, desc="", parent=None, deps=(),
            milestone=None, props=None, priority="0", hours=0, attach=(), sequence=None):
        T.append(dict(key=key, name=name, stage=stage, tags=tags, users=users, start=start, end=end,
                      state=state, desc=desc, parent=parent, deps=list(deps), milestone=milestone,
                      props=props or {}, priority=priority, hours=hours, attach=list(attach),
                      sequence=sequence if sequence is not None else len(T) + 1))

    P1, P2, P3, P4 = (s for s, _ in TASK_STAGES)
    M = [m[0] for m in MILESTONES]

    # ---------------- PH1 - CONTRACT (slide 3: Sales / Admin / Legal / Operations) ----------------
    add("p1_bid", "Bid Results / Accepted Proposal", P1, ["Sales"], ROLE_USERS["Sales"],
        "2026-04-13", "2026-04-14", DONE,
        "<p>Bid results received and proposal accepted for the <b>Roof Replacement - Kinder Wing, "
        "Ruskin Elementary School</b> (Bid Set, SEI Job No. 26006, dated 3/12/26).</p>")
    add("p1_80", "80% Report", P1, ["Admin"], ROLE_USERS["Admin"], "2026-04-15", "2026-04-15", DONE,
        deps=["p1_bid"])
    add("p1_loi", "Receive LOI or Notice of Award", P1, ["Admin"], ROLE_USERS["Admin"],
        "2026-04-16", "2026-04-16", DONE, deps=["p1_80"])
    add("p1_overview", "Provide Project Overview", P1, ["Sales"], ROLE_USERS["Sales"],
        "2026-04-17", "2026-04-20", DONE, deps=["p1_loi"])
    add("p1_90", "90% Report", P1, ["Admin"], ROLE_USERS["Admin"], "2026-04-20", "2026-04-20", DONE,
        deps=["p1_overview"])
    add("p1_sharepoint", "Create SharePoint Link", P1, ["Admin"], ROLE_USERS["Admin"],
        "2026-04-20", "2026-04-21", DONE, deps=["p1_loi"])
    add("p1_jobpkg", "Prepare Job Package, Fill Out Routing Form, Upload Files in SharePoint, Send Link to Ops Team",
        P1, ["Sales"], ROLE_USERS["Sales"], "2026-04-21", "2026-04-24", DONE,
        checklist([("Prepare job package", True), ("Fill out routing form", True),
                   ("Upload files in SharePoint", True), ("Send link to Ops team", True)]),
        deps=["p1_90", "p1_sharepoint"])
    add("p1_assign", "Assign Operations Team (PM / Super / PE)", P1, ["Operations"], ROLE_USERS["Operations"],
        "2026-04-27", "2026-04-27", DONE,
        "<p>Operations team assigned:</p>" + ul([
            "Project Executive - Alona Bowman", "Project Manager - Maria Alimagno",
            "Project Engineer (Roofing) - Kris Lauengco", "Project Engineer (Waterproofing) - Raymond Lobato",
            "Superintendent - Chris Wilson", "Foreman - Joseph Gonzales", "QC Manager - Wes Fleming",
            "Submittals / Admin - Kassandra Solis"]),
        deps=["p1_jobpkg"])
    add("p1_review_pkg", "Review Job Package (1-2 weeks)", P1, ["Operations"], ROLE_USERS["Operations"],
        "2026-04-27", "2026-05-08", DONE, deps=["p1_assign"], hours=16)
    add("p1_receive_contract", "Receive Contract", P1, ["Admin"], ROLE_USERS["Admin"],
        "2026-05-04", "2026-05-04", DONE, deps=["p1_jobpkg"])
    add("p1_review_contract", "Review Contract", P1, ["Legal"], ROLE_USERS["Legal"],
        "2026-05-05", "2026-05-12", DONE, deps=["p1_receive_contract"])
    add("p1_ph1mtg", "Set-Up Phase 1 (PH1) Meeting", P1, ["Operations", "Meeting"], ROLE_USERS["Operations"],
        "2026-05-11", "2026-05-11", DONE, deps=["p1_review_pkg"])
    add("p1_vista", "Project Set Up in Vista", P1, ["Admin"], ROLE_USERS["Admin"],
        "2026-05-13", "2026-05-14", DONE, deps=["p1_review_contract"])
    add("p1_signoff", "Sign Off on Routing Form and Scope of Work", P1, ["Operations"], ROLE_USERS["Operations"],
        "2026-05-15", "2026-05-15", DONE, deps=["p1_review_pkg", "p1_review_contract"])
    add("p1_execute", "Execute Contract", P1, ["Legal"], ROLE_USERS["Legal"], "2026-05-22", "2026-05-22", DONE,
        deps=["p1_signoff", "p1_vista"], milestone=M[0], priority="1")

    # ---------------- PH2 - PRECON (slide 4: PM / Field Mgr / Super / PE) ----------------
    add("p2_review_pkg", "Review Job Package (Ops Team)", P2, ["Project Manager", "Superintendent", "Project Engineer"],
        OPS_TEAM, "2026-05-25", "2026-05-26", DONE, deps=["p1_execute"])
    add("p2_master_sched", "Enter Project in Master Field Resource Schedule", P2, ["Project Manager"],
        ROLE_USERS["Project Manager"], "2026-05-26", "2026-05-26", DONE, deps=["p2_review_pkg"])
    add("p2_ph1_attend", "Attend PH1 Meeting / Strategize Right After Meeting / Set Milestones", P2,
        ["Project Manager", "Superintendent", "Project Engineer", "Field Manager", "Meeting"], OPS_TEAM,
        "2026-05-27", "2026-05-27", DONE,
        "<p>Milestones set for the project - see the <b>Milestones</b> of this project.</p>",
        deps=["p2_review_pkg"], milestone=M[1])
    add("p2_create_sublog", "Create Submittal Log", P2, ["Project Engineer", "Submittal"], ROLE_USERS["Project Engineer"],
        "2026-05-27", "2026-05-29", DONE,
        "<p>Submittal log created from the Bid Set specifications (SEI Job No. 26006). "
        "Individual submittals are tracked as sub-tasks of <i>Complete Submittal Process</i>.</p>",
        deps=["p2_review_pkg"])
    add("p2_kickoff", "Project Kick-Off with Client / Owner, Manufacturer, Lower Tier Sub (Intro Mtg, Jobsite Visit)",
        P2, ["Project Manager", "Superintendent", "Project Engineer", "Field Manager", "Meeting"], OPS_TEAM,
        "2026-06-01", "2026-06-02", DONE,
        "<p>Attendees: Berryessa Union School District (Owner), Strawn Construction (GC - Donny Durham), "
        "Steelhead Engineers (SEI), manufacturer reps, lower-tier subs.</p>"
        "<p>Project address: 1401 Turlock Lane, San Jose, CA 95132 (Ruskin Elementary School - Kinder Wing, "
        "Buildings K1 &amp; K2).</p>",
        deps=["p2_ph1_attend"])
    add("p2_sov", "Generate SOV", P2, ["Project Manager"], ROLE_USERS["Project Manager"],
        "2026-06-03", "2026-06-04", DONE, deps=["p2_kickoff"])
    add("p2_submittals", "Complete Submittal Process", P2, ["Project Engineer", "Submittal"],
        ["kris", "kassandra"], "2026-06-01", "2026-10-02", PROG,
        "<p>Parent task for the <b>Submittal Log</b> (tab of the Weekly Status Report). "
        "Each sub-task is one submittal package with its spec section, type, dates and review status "
        "stored in the task properties.</p>",
        deps=["p2_create_sublog"], milestone=M[2], priority="1")
    add("p2_proc_sched", "Generate Procurement Schedule", P2, ["Project Manager", "Procurement"],
        ROLE_USERS["Project Manager"], "2026-06-08", "2026-06-10", DONE, deps=["p2_sov"])
    add("p2_budget", "Finalize Budget (Material, Labor, Equipment)", P2,
        ["Project Manager", "Field Manager", "Superintendent"], ["maria", "chris"],
        "2026-06-15", "2026-09-30", PROG,
        "<p>Budget finalization re-opened after <b>RFI 001</b>: SEI confirmed a <b>4-ply built-up membrane "
        "plus cap sheet</b> (Spec 07 51 05 par. 3.06.B and detail 1/A10.40), not 3-ply + cap sheet per the "
        "par. 1.01 header.</p>" + checklist([("Material budget", False), ("Labor budget", False),
                                              ("Equipment budget", True)]),
        deps=["p2_proc_sched"])
    add("p2_binder", "Provide Field Binder", P2, ["Project Engineer"], ROLE_USERS["Project Engineer"],
        "2026-09-28", "2026-10-02", PROG,
        checklist([("Approved submittals (07 51 05, 07 60 05, 07 90 05)", False),
                   ("RFI 001 response - 4-ply BUR + cap sheet", True),
                   ("Detail 1/A10.40 and roof plans", True),
                   ("Manufacturer installation instructions / PDS / SDS", False),
                   ("Site safety plan & QC plan", False)]),
        deps=["p2_submittals"])
    add("p2_pep", "Project Execution Plan (PH1A Meeting)", P2,
        ["Project Manager", "Superintendent", "Project Engineer", "Field Manager", "Meeting"], OPS_TEAM,
        "2026-10-01", "2026-10-01", WAIT, deps=["p2_budget", "p2_binder"])
    add("p2_ph2_sales", "Request PH2 Meeting with Sales if Budget Needs to be Re-Adjusted", P2,
        ["Project Manager", "Sales", "Meeting"], ["maria", "alona"], "2026-09-29", "2026-10-01", PROG,
        "<p>Budget re-adjustment required: RFI 001 response changes the roof assembly to 4-ply + cap sheet. "
        "Review with Sales against the accepted proposal.</p>", deps=["p2_budget"], priority="1")

    # --- Submittal log (sub-tasks of Complete Submittal Process) ---
    sealant_desc = (
        "<h3>Submittal Transmittal</h3>"
        + html_table(["Field", "Value"], [
            ["Date", "06/15/26"], ["Project", "Ruskin ES K1 &amp; K2"],
            ["Customer", "Berryessa Union School District"],
            ["Address", "1401 Turlock Lane, San Jose, CA 95132 (Project address)"],
            ["Attention", "Donny Durham - 408-286-1299 - ddurham@scmdinc.com"],
            ["Submitted by", "Kassandra Solis - Best Contracting Services, Inc. - 310-328-6969"]])
        + "<h3>Contents</h3>"
        + html_table(["Spec #", "Description", "Page #"], [
            ["07 90 05", "<b>Sealants</b>", "-"], ["", "Transmittal Sheet", "1"], ["", "Spec", "2-3"],
            ["1.04.A", "Product Data", "4-25"], ["2.01.A", "Sika - 1A - PDS/SDS *Color Chart*", "4-25"]])
        + "<p><i>Note on transmittal: a pre-job meeting will be required approximately 2 weeks prior to the "
          "anticipated start date. Attached general design criteria and standard manufacturer details should "
          "be forwarded to any trades interfacing with our scope of work.</i></p>"
        + "<h3>Spec 07 90 05 - Sealants (key requirements)</h3>"
        + checklist([
            ("1.04.A Submit manufacturer literature, specifications and color charts for sealants and primers", True),
            ("2.01.A Metal-to-metal joints: one-part polyurethane, ASTM C920 (e.g. NP-1) - <b>Sikaflex-1A proposed</b>; color by Owner", True),
            ("2.01.B Concealed metal-to-metal joints: one-part butyl sealant, ASTM C1085", False),
            ("2.01.C Hot pipes: non-corrosive one-part silicone, -60F to +400F (Dow Corning 999-A or equal)", False),
            ("2.02.A Backer rod: round closed-cell polyethylene, compressed 25% to fit joint", False),
            ("2.02.B Primers / cleaners as recommended by sealant manufacturer", False)])
        + "<h3>Proposed product - Sikaflex-1A</h3>"
        + ul(["Premium 1-component, moisture-cured polyurethane elastomeric sealant / adhesive",
              "ASTM C920 Type S, Grade NS, Class 35, Use T, NT, O, M, A, I; TT-S-00230C Type II Class A",
              "Movement capability +/-35%; service temp -40F to +170F",
              "Packaging: 10.1 oz cartridge, 20 oz sausage, 4.5 gal pail, 52 gal drum",
              "Colors: white, colonial white, aluminum gray, limestone, black, dark bronze, capitol tan, stone, medium bronze",
              "Tack free 3-6 hrs; cure 4-7 days (73F / 50% RH)",
              "Limitation: do not use in contact with bituminous / asphaltic materials"])
        + "<p>Related sections: 07 51 05 Built-Up Bituminous Roofing; 07 60 05 Flashing and Sheet Metal.</p>")
    add("sub01", "SUB-01 | 07 90 05 Sealants - Product Data (Sikaflex-1A PDS/SDS & Color Chart)", P2,
        ["Submittal", "07 90 05 Sealants", "Project Engineer"], ["kassandra", "kris"],
        "2026-06-15", "2026-06-29", WAIT, sealant_desc, parent="p2_submittals", priority="1",
        props={"log_ref": "SUB-01", "spec_section": "07 90 05", "spec_par": "1.04.A / 2.01.A",
               "doc_type": "product_data", "review_status": "submitted", "date_sent": "2026-06-15",
               "review_days": 14, "date_due_client": "2026-06-29", "manufacturer": "Sika (Sikaflex-1A)",
               "ball_in_court": "Strawn Construction / SEI"},
        attach=["SUB - 07 90 05 - Sealants.pdf"])
    add("sub02", "SUB-02 | 07 51 05 Built-Up Bituminous Roofing - Product Data (4-ply + Cap Sheet)", P2,
        ["Submittal", "07 51 05 Built-Up Roofing", "Project Engineer"], ["kris"],
        "2026-06-22", "2026-09-25", PROG,
        "<p>Prepare product data for the built-up roof assembly per RFI 001 response: "
        "<b>4-ply built-up membrane and a cap sheet</b> (par. 3.06.B, detail 1/A10.40).</p>",
        parent="p2_submittals",
        props={"log_ref": "SUB-02", "spec_section": "07 51 05", "spec_par": "3.06.B", "drawing_ref": "1/A10.40",
               "doc_type": "product_data", "review_status": "in_progress", "review_days": 14,
               "ball_in_court": "BCSI - Kris Lauengco"})
    add("sub03", "SUB-03 | 07 60 05 Flashing and Sheet Metal - Product Data & Shop Drawings", P2,
        ["Submittal", "07 60 05 Flashing & Sheet Metal", "Project Engineer"], ["kris"],
        "2026-06-22", "2026-09-25", PROG, parent="p2_submittals",
        props={"log_ref": "SUB-03", "spec_section": "07 60 05", "doc_type": "shop_drawings",
               "review_status": "in_progress", "review_days": 14, "ball_in_court": "BCSI - Kris Lauengco"})

    # ---------------- PH3 - CONSTRUCTION (slide 5) ----------------
    add("p3_weekly", "Weekly Monitoring / Tracking: Labor, Materials, Equipment, RFI, Change Orders, Field Issues", P3,
        ["Project Manager", "Superintendent", "Project Engineer", "Field Manager", "Weekly Report"], OPS_TEAM,
        "2026-06-01", "2026-12-11", PROG,
        "<p>Weekly status tracking - mirrors the <b>Weekly Status Report</b> workbook "
        "(Project Info, Submittal Log, Incoming Docs, Outgoing RFI Log, BCSI Activities-Schedule, "
        "Material Procurement, Action Item / Issue Log). Published every week as a <b>Project Update</b>.</p>",
        deps=["p2_ph1_attend"])
    add("p3_procurement_pm", "Procurement (PM) - Release POs", P3, ["Project Manager", "Procurement"],
        ROLE_USERS["Project Manager"], "2026-09-21", "2026-10-02", PROG,
        "<p>Material Procurement log (tab of the Weekly Status Report). One sub-task per material, "
        "with manufacturer, PO #, lead time and delivery dates in the task properties.</p>",
        deps=["p2_proc_sched"])
    add("p3_procurement_super", "Procurement (Super) - Receive & Stock Materials On Site", P3,
        ["Superintendent", "Procurement"], ROLE_USERS["Superintendent"], "2026-10-01", "2026-10-05", WAIT,
        deps=["p3_procurement_pm"])
    add("p3_labor", "Labor Resource Allocation for All On-Going Jobs and 4-Week Look Ahead Schedule", P3,
        ["Field Manager"], ROLE_USERS["Field Manager"], "2026-09-21", "2026-12-11", PROG,
        deps=["p2_master_sched"])
    add("p3_rfis", "RFIs", P3, ["Project Engineer", "RFI"], ROLE_USERS["Project Engineer"],
        "2026-06-01", "2026-12-11", PROG,
        "<p>Outgoing RFI Log (tab of the Weekly Status Report). One sub-task per RFI; "
        "date sent, due date (sent + 7), response, status and scope/cost impact tracked in task properties.</p>",
        deps=["p2_create_sublog"])
    add("p3_co", "Change Order", P3, ["Project Manager", "Change Order"], ROLE_USERS["Project Manager"],
        "2026-06-15", "2026-12-11", PROG,
        "<p>Change order log. Potential change orders are raised from RFIs / incoming documents with a "
        "scope, schedule or cost impact.</p>", deps=["p3_rfis"])
    add("p3_qc", "Oversee Quality & Safety in the Field", P3, ["Field Manager", "QC", "Safety & Quality"],
        ["chris", "wes"], "2026-10-05", "2026-12-11", WAIT, deps=["p2_pep"])

    rfi_desc = (
        html_table(["Field", "Value"], [
            ["RFI #", "001"], ["Date", "6/2/26"], ["Subject", "Built-Up Roof Assembly"],
            ["To", "Donny Durham - Strawn Construction (ddurham@scmdinc.com)"],
            ["From", "Maria Alimagno - Best Contracting Services, Inc. (310-328-6969, malimagno@bestcontracting.com)"],
            ["Drawing reference", "1/A10.40"], ["Spec section", "07 51 05"], ["Respond by", "ASAP"],
            ["Potential impact", "Cost Increase [ ]  Time Increase [ ]  Unknown Cost [ ]  Unknown Time [ ]"]])
        + "<h3>Question</h3><p>Spec paragraph 1.01 clearly calls out 3-ply and cap sheet. However, it is unclear on "
          "detail 1/A10.40 if the cap sheet is part of the 4-ply system. Please confirm roof system to be per specs, "
          "3-ply and cap sheet.</p>"
        + "<h3>Response - Steelhead Engineers</h3><blockquote><b>Disregard the header on paragraph 1.01. Follow "
          "paragraph 3.06.B and 1/A10.40 which specify a 4-ply built-up membrane and a cap sheet.</b></blockquote>"
          "<p>Signed: Steelhead Engineers - dated 6/10/26 (per SEI response file).</p>"
        + "<h3>Actions</h3>" + checklist([
            ("Distribute response to field (Super / Foreman) and update field binder", True),
            ("Update SUB-02 (07 51 05) to 4-ply + cap sheet", False),
            ("Evaluate cost / time impact vs. accepted proposal -> PCO-01", False)]))
    add("rfi001", "RFI 001 | Built-Up Roof Assembly (3-ply vs 4-ply + Cap Sheet) - ANSWERED by SEI", P3,
        ["RFI", "Incoming Doc", "07 51 05 Built-Up Roofing", "Project Engineer"], ["maria", "kris"],
        "2026-06-02", "2026-06-10", DONE, rfi_desc, parent="p3_rfis", priority="1",
        props={"log_ref": "RFI 001", "spec_section": "07 51 05", "spec_par": "1.01 / 3.06.B",
               "drawing_ref": "1/A10.40", "doc_type": "rfi", "review_status": "answered",
               "date_sent": "2026-06-02", "review_days": 7, "date_due_client": "2026-06-09",
               "date_response": "2026-06-10", "responded_by": "Steelhead Engineers (SEI)",
               "scope_impact": True, "cost_impact": True, "ball_in_court": "BCSI - Maria Alimagno"},
        attach=["RFI 001 - Built-Up Roof Assembly - SEI Response 2026-06-10.pdf"])
    add("pco01", "PCO-01 | 4-Ply Built-Up Membrane + Cap Sheet per RFI 001 - Evaluate Cost / Schedule Impact", P3,
        ["Change Order", "07 51 05 Built-Up Roofing", "Project Manager"], ["maria"],
        "2026-06-11", "2026-10-02", PROG,
        "<p>RFI 001 response directs a <b>4-ply</b> built-up membrane plus cap sheet (par. 3.06.B / 1/A10.40) "
        "instead of 3-ply + cap sheet (par. 1.01 header). Compare against the accepted proposal and, "
        "if the bid was based on 3-ply, submit a change order request to Strawn Construction.</p>"
        + checklist([("Quantify added ply: material, labor, equipment", False),
                     ("Confirm with Sales what the proposal was based on", False),
                     ("Submit COR to GC", False)]),
        parent="p3_co", deps=["rfi001"], priority="1",
        props={"log_ref": "PCO-01", "spec_section": "07 51 05", "drawing_ref": "1/A10.40", "doc_type": "co",
               "review_status": "in_progress", "scope_impact": True, "cost_impact": True,
               "ball_in_court": "BCSI - Maria Alimagno"})

    # Material procurement sub-tasks
    add("mat_sealant", "MATL | 07 90 05 Sikaflex-1A Polyurethane Sealant (Owner color TBD)", P3,
        ["Procurement", "07 90 05 Sealants"], ["maria"], "2026-09-28", "2026-10-02", WAIT,
        "<p>Release after SUB-01 approval and Owner color selection (spec 2.01.A). "
        "Store at 40-95F; cartridge shelf life 15 months.</p>",
        parent="p3_procurement_pm", deps=["sub01"],
        props={"spec_section": "07 90 05", "manufacturer": "Sika Corporation", "lead_time_days": 7,
               "review_status": "open"})
    add("mat_bur", "MATL | 07 51 05 Built-Up Roofing Materials - 4-Ply Felts + Cap Sheet", P3,
        ["Procurement", "07 51 05 Built-Up Roofing"], ["maria"], "2026-09-21", "2026-10-02", WAIT,
        parent="p3_procurement_pm", deps=["sub02", "pco01"],
        props={"spec_section": "07 51 05", "lead_time_days": 14, "review_status": "open"})
    add("mat_metal", "MATL | 07 60 05 Flashing & Sheet Metal", P3, ["Procurement", "07 60 05 Flashing & Sheet Metal"],
        ["maria"], "2026-09-21", "2026-10-09", WAIT, parent="p3_procurement_pm", deps=["sub03"],
        props={"spec_section": "07 60 05", "lead_time_days": 21, "review_status": "open"})

    # Action item / issue log (sub-tasks of weekly monitoring)
    add("ai01", "ACTION | Obtain Owner Color Selection for Sikaflex-1A (Spec 07 90 05 - 2.01.A)", P3,
        ["Action Item", "07 90 05 Sealants"], ["kris"], "2026-09-28", "2026-10-02", PROG,
        "<p>Spec 2.01.A: color to be selected by Owner. Sikaflex-1A color chart included in SUB-01.</p>",
        parent="p3_weekly", priority="1",
        props={"review_status": "open", "ball_in_court": "Owner / GC (Donny Durham)"})
    add("ai02", "ACTION | Schedule Pre-Job Meeting (~2 Weeks Before Start) per Submittal Transmittal", P3,
        ["Action Item", "Meeting"], ["maria"], "2026-09-21", "2026-09-21", DONE,
        parent="p3_weekly",
        props={"review_status": "closed", "ball_in_court": "BCSI - Maria Alimagno"})
    add("ai03", "ACTION | Forward Sealant Design Criteria & Manufacturer Details to Interfacing Trades", P3,
        ["Action Item", "07 90 05 Sealants"], ["kassandra"], "2026-09-28", "2026-10-02", PROG,
        parent="p3_weekly", props={"review_status": "in_progress", "ball_in_court": "BCSI - Kassandra Solis"})
    add("ai04", "ACTION | Confirm Sealant Compatibility at Bituminous Surfaces (Sikaflex-1A limitation)", P3,
        ["Action Item", "07 90 05 Sealants", "QC"], ["wes"], "2026-09-28", "2026-10-02", PROG,
        "<p>Sikaflex-1A PDS: <i>do not use in contact with bituminous / asphaltic materials</i>. "
        "Confirm locations with SEI / Sika rep (sheet-metal joints only).</p>",
        parent="p3_weekly", priority="1",
        props={"review_status": "open", "ball_in_court": "BCSI - Wes Fleming"})

    # BCSI Activities - schedule (field activities, per spec sections)
    add("act", "BCSI Activities - Schedule (Field Work K1 & K2)", P3, ["Field Activity", "Superintendent"],
        ["chris", "joseph"], "2026-10-05", "2026-11-20", WAIT,
        "<p>Field activities (BCSI Activities-Sched tab). Buildings <b>K1</b> and <b>K2</b>, Kinder Wing.</p>",
        deps=["p2_pep", "p3_procurement_super"], milestone=M[3])
    for bldg, s0 in (("K1", 5), ("K2", 26)):
        base = d("2026-10-01") + dt.timedelta(days=s0 - 1)
        seq = [
            (f"Tear-Off Existing Roofing - {bldg}", 4, ["Field Activity"], ""),
            (f"Install 4-Ply Built-Up Membrane + Cap Sheet - {bldg}", 8, ["Field Activity", "07 51 05 Built-Up Roofing"],
             "07 51 05"),
            (f"Flashing & Sheet Metal - {bldg}", 4, ["Field Activity", "07 60 05 Flashing & Sheet Metal"], "07 60 05"),
            (f"Sealants at Sheet Metal Joints & Hot Pipes - {bldg}", 2, ["Field Activity", "07 90 05 Sealants"], "07 90 05"),
        ]
        prev = None
        cur = base
        for i, (nm, days, tags, spec) in enumerate(seq, 1):
            key = f"act_{bldg}_{i}"
            end = cur + dt.timedelta(days=days - 1)
            add(key, nm, P3, tags, ["chris", "joseph"], cur.isoformat(), end.isoformat(), WAIT,
                parent="act", deps=[prev] if prev else ["mat_bur"] if i == 2 else [],
                milestone=M[4] if i == 4 else None,
                props={"spec_section": spec} if spec else {})
            prev = key
            cur = end + dt.timedelta(days=1)
            while cur.weekday() >= 5:
                cur += dt.timedelta(days=1)
    add("act_sampling", "Representative Sampling of Sealant Joints w/ Owner (Spec 07 90 05 - 3.04)", P3,
        ["Field Activity", "07 90 05 Sealants", "QC"], ["wes", "chris"], "2026-11-19", "2026-11-20", WAIT,
        parent="act", deps=["act_K2_4"], props={"spec_section": "07 90 05", "spec_par": "3.04.A"})

    # ---------------- PH4 - CLOSEOUT (slide 5) ----------------
    add("p4_pos", "Clean-Up POs", P4, ["Project Manager", "Procurement"], ROLE_USERS["Project Manager"],
        "2026-12-14", "2026-12-18", WAIT, deps=["act"])
    add("p4_punch", "Punchlist (Owner / GC / Manufacturer)", P4, ["Superintendent", "QC"], ["chris", "wes"],
        "2026-12-14", "2026-12-18", WAIT, deps=["act"], milestone=M[5])
    add("p4_co", "Finalize Change Orders", P4, ["Project Manager", "Change Order"], ROLE_USERS["Project Manager"],
        "2026-12-21", "2027-01-08", WAIT, deps=["p4_pos"])
    add("p4_closeout_log", "Closeout Log to Admin", P4, ["Project Engineer", "Admin"], ["kris", "kassandra"],
        "2026-12-21", "2027-01-15", WAIT,
        checklist([("Warranties (roofing manufacturer & contractor)", False), ("As-builts", False),
                   ("O&M / product data", False), ("Final lien releases", False)]),
        deps=["p4_punch"])
    add("p4_lessons", "PH3 Meeting - Lessons Learned", P4,
        ["Project Manager", "Superintendent", "Project Engineer", "Field Manager", "Meeting"], OPS_TEAM,
        "2027-01-20", "2027-01-20", WAIT, deps=["p4_co", "p4_closeout_log"])
    add("p4_archive", "Archive Project Files", P4, ["Project Manager", "Admin"], ["maria", "kassandra"],
        "2027-01-25", "2027-01-29", WAIT, deps=["p4_lessons"], milestone=M[6])
    return T


def weekly_report_html(tasks):
    by = {t["key"]: t for t in tasks}
    status_lbl = {s[0]: s[1] for s in PROPERTIES[5]["selection"]}
    sub_rows = [[by[k]["props"].get("log_ref"), by[k]["props"].get("spec_section"), by[k]["name"].split("| ")[-1],
                 by[k]["props"].get("date_sent", ""), by[k]["props"].get("date_due_client", ""),
                 status_lbl.get(by[k]["props"].get("review_status"), "")] for k in ("sub01", "sub02", "sub03")]
    r = by["rfi001"]["props"]
    return (
        "<h2>Weekly Status Report #01 - 09/30/26</h2>"
        "<h3>Project Info</h3>"
        + html_table(["", ""], [
            ["Project / Job", "Ruskin ES K1 &amp; K2 - Roof Replacement, Kinder Wing"],
            ["Customer", "Berryessa Union School District - 1401 Turlock Lane, San Jose, CA 95132"],
            ["GC", "Strawn Construction - Donny Durham"], ["Consultant", "Steelhead Engineers (SEI Job 26006)"],
            ["Scope", "Built-up bituminous roofing (4-ply + cap sheet), flashing &amp; sheet metal, sealants"],
            ["Phase", CURRENT_PROJECT_STAGE]])
        + "<h3>Submittal Log</h3>"
        + html_table(["No.", "Spec", "Description", "Sent to GC", "Due from Client", "Status"], sub_rows)
        + "<h3>Incoming Docs / Outgoing RFI Log</h3>"
        + html_table(["Ref", "Subject", "Sent", "Due", "Response", "Status", "Scope/Sched", "Cost"], [
            ["RFI 001", "Built-Up Roof Assembly", r["date_sent"], r["date_due_client"], r["date_response"],
             "Answered - 4-ply + cap sheet", "Yes", "Yes (PCO-01)"]])
        + "<h3>Material Procurement</h3>"
        + html_table(["Spec", "Material", "Manufacturer", "Status"], [
            ["07 90 05", "Sikaflex-1A sealant", "Sika", "Pending SUB-01 approval + Owner color"],
            ["07 51 05", "BUR 4-ply felts + cap sheet", "", "Pending SUB-02 / PCO-01"],
            ["07 60 05", "Flashing &amp; sheet metal", "", "Pending SUB-03"]])
        + "<h3>Action Items / Issue Log</h3>"
        + html_table(["Item", "Priority", "BIC", "Status"], [
            ["Owner color selection - Sikaflex-1A", "High", "Owner / GC", "Open"],
            ["Pre-job meeting ~2 weeks before start", "Medium", "Maria", "Done"],
            ["Forward sealant details to interfacing trades", "Medium", "Kassandra", "In-progress"],
            ["Sealant compatibility at bituminous surfaces", "High", "Wes", "Open"]])
        + "<h3>Risks</h3>" + ul([
            "SUB-01 (Sealants) review outstanding since 06/15/26.",
            "RFI 001 changes roof assembly to 4-ply + cap sheet - budget/CO impact under review (PCO-01).",
            "Mobilization 10/05/26 depends on submittal approvals and material release."]))


# ---------------------------------------------------------------------------
# Odoo client
# ---------------------------------------------------------------------------
class Odoo:
    def __init__(self, url, db, login, key):
        self.url = url.rstrip("/")
        common = xmlrpc.client.ServerProxy(f"{self.url}/xmlrpc/2/common", allow_none=True)
        self.version = common.version()
        self.db = db or self._guess_db()
        self.uid = common.authenticate(self.db, login, key, {})
        if not self.uid:
            sys.exit(f"Authentication failed for {login} on database {self.db}")
        self.key = key
        self.models = xmlrpc.client.ServerProxy(f"{self.url}/xmlrpc/2/object", allow_none=True)
        self._fields = {}

    def _guess_db(self):
        try:
            req = urllib.request.Request(f"{self.url}/web/database/list",
                                         data=json.dumps({"jsonrpc": "2.0", "method": "call", "params": {}}).encode(),
                                         headers={"Content-Type": "application/json"})
            res = json.loads(urllib.request.urlopen(req, timeout=30).read()).get("result")
            if res and len(res) == 1:
                return res[0]
        except Exception:
            pass
        return self.url.split("//")[1].split(".")[0]

    def x(self, model, method, *args, **kw):
        return self.models.execute_kw(self.db, self.uid, self.key, model, method, list(args), kw)

    def fields(self, model):
        if model not in self._fields:
            self._fields[model] = self.x(model, "fields_get", attributes=["type", "selection"])
        return self._fields[model]

    def has(self, model, field):
        return field in self.fields(model)

    def clean(self, model, vals):
        return {k: v for k, v in vals.items() if self.has(model, k)}

    def ref(self, xmlid):
        mod, name = xmlid.split(".")
        r = self.x("ir.model.data", "search_read", [("module", "=", mod), ("name", "=", name)],
                   fields=["res_id"], limit=1)
        return r[0]["res_id"] if r else False

    OPTIONAL = ("task_properties", "task_properties_definition", "planned_date_begin", "milestone_id",
                "allocated_hours", "stage_id", "recurring_task", "repeat_interval", "repeat_unit", "repeat_type",
                "repeat_until", "favorite_user_ids", "notification_type", "website", "comment", "progress")

    def upsert(self, model, domain, vals, ctx=None):
        ids = self.x(model, "search", domain, limit=1, context={"active_test": False})
        vals = self.clean(model, vals)

        def save(v):
            if ids:
                self.x(model, "write", ids, v, context=ctx or QUIET)
                return ids[0]
            return self.x(model, "create", v, context=ctx or QUIET)
        try:
            return save(vals)
        except xmlrpc.client.Fault as e:
            dropped = [k for k in self.OPTIONAL if k in vals]
            if not dropped:
                raise
            print(f"  note: {model} '{vals.get('name', '')[:40]}' saved without {', '.join(dropped)} "
                  f"({e.faultString.strip().splitlines()[-1][:120]})")
            return save({k: v for k, v in vals.items() if k not in dropped})


class DryRun(Odoo):
    """Offline stand-in that accepts every call - used to validate the script's data."""

    def __init__(self):
        self.version = {"server_version": "dry-run"}
        self.db, self.uid, self._next, self.url = "dry-run", 2, 100, "https://dry-run"
        self.calls = []
        self._fields = {}

    def x(self, model, method, *args, **kw):
        self.calls.append((model, method))
        if method in ("search",):
            return []
        if method == "search_read":
            return [{"id": 1, "res_id": 1}]
        if method == "create":
            self._next += 1
            return self._next
        if method == "fields_get":
            return {}
        return True

    def has(self, model, field):
        return True


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------
def build(o, create_users=True):
    log = print
    log(f"Connected: db={o.db} uid={o.uid} version={o.version.get('server_version')}")

    # 1. Enable Project features: milestones, dependencies, project stages, recurring tasks
    feats = {k: True for k in ("group_project_milestone", "group_project_task_dependencies",
                               "group_project_stages", "group_project_recurring_tasks")
             if o.has("res.config.settings", k)}
    if feats:
        try:
            sid = o.x("res.config.settings", "create", feats)
            o.x("res.config.settings", "execute", [sid])
            log(f"Enabled Project settings: {', '.join(feats)}")
        except Exception as e:
            log(f"  settings not changed: {str(e).strip().splitlines()[-1][:120]}")

    # 2. Contacts
    us = o.x("res.country", "search", [("code", "=", "US")], limit=1)
    ca = o.x("res.country.state", "search", [("code", "=", "CA"), ("country_id", "in", us)], limit=1) if us else []
    cust = dict(CUSTOMER)
    cust.pop("state_code"), cust.pop("country_code")
    cust.update({"country_id": us[0] if us else False, "state_id": ca[0] if ca else False})
    partner = {"customer": o.upsert("res.partner", [("name", "=", CUSTOMER["name"])], cust)}
    for c in COMPANIES:
        vals = {k: v for k, v in c.items() if k != "key"}
        partner[c["key"]] = o.upsert("res.partner", [("name", "=", c["name"])], vals)
    for c in CONTACTS:
        vals = {k: v for k, v in c.items() if k not in ("key", "parent")}
        vals["parent_id"] = partner[c["parent"]]
        partner[c["key"]] = o.upsert("res.partner", [("name", "=", c["name"])], vals)
    log(f"Contacts ready: {len(partner)}")

    # 3. Users (BEST team)
    users = {}
    grp_user = o.ref("project.group_project_user")
    for m in TEAM:
        found = o.x("res.users", "search", [("name", "=", m["name"])], limit=1, context={"active_test": False})
        if found:
            users[m["key"]] = found[0]
            continue
        if not create_users:
            continue
        login = f"{m['login']}@{LOGIN_DOMAIN}"
        found = o.x("res.users", "search", [("login", "=", login)], limit=1, context={"active_test": False})
        if found:
            users[m["key"]] = found[0]
            continue
        vals = {"name": m["name"], "login": login, "function": m["role"], "notification_type": "inbox"}
        try:
            uid = o.x("res.users", "create", o.clean("res.users", vals), context=QUIET)
        except Exception as e:
            log(f"  could not create user {m['name']} ({str(e)[:120]}) - tasks will be left unassigned for them")
            continue
        if grp_user:
            try:
                gfield = "group_ids" if o.has("res.users", "group_ids") and not isinstance(o, DryRun) else "groups_id"
                o.x("res.users", "write", [uid], {gfield: [(4, grp_user)]}, context=QUIET)
            except Exception:
                pass
        users[m["key"]] = uid
    log(f"Team users ready: {len(users)} / {len(TEAM)}")

    def uids(keys):
        return [users[k] for k in keys if k in users]

    # 4. Tags
    tags = {n: o.upsert("project.tags", [("name", "=", n)], {"name": n, "color": c}) for n, c in TAG_COLORS.items()}

    # 5. Project stages
    pstage = {}
    for i, (n, fold) in enumerate(PROJECT_STAGES, 1):
        pstage[n] = o.upsert("project.project.stage", [("name", "=", n)], {"name": n, "sequence": i, "fold": fold})

    # 6. Project
    tasks = build_tasks()
    desc = (
        "<h3>Ruskin Elementary School - Roof Replacement, Kinder Wing (Buildings K1 &amp; K2)</h3>"
        + html_table(["", ""], [
            ["Customer", "Berryessa Union School District"],
            ["Project address", "1401 Turlock Lane, San Jose, CA 95132"],
            ["General Contractor", "Strawn Construction - Donny Durham, 408-286-1299, ddurham@scmdinc.com"],
            ["Roofing Consultant", "Steelhead Engineers (SEI) - Job No. 26006, Bid Set 3/12/26"],
            ["Scope", "07 51 05 Built-Up Bituminous Roofing (4-ply + cap sheet per RFI 001); "
                      "07 60 05 Flashing &amp; Sheet Metal; 07 90 05 Sealants"],
            ["BCSI", "Best Contracting Services, Inc. - 310-328-6969"]])
        + "<h4>BEST Project Team</h4>"
        + html_table(["Name", "Position"], [[m["name"], m["role"]] for m in TEAM])
        + "<h4>Client / CM Collaboration Tools</h4>"
        + html_table(["App", "Purpose"], [["BIM 360", "Plans"], ["Smartsheet", "Triage / QC Reports"],
                                          ["Extracker", "Change orders"], ["Smartsheet", "Delivery Calendar"]])
        + "<h4>Contract</h4>"
        + html_table(["", ""], [["Original contract amount", ""], ["Executed change orders to date", ""],
                                ["Contract amount to date", ""], ["Projected GP%", ""]])
        + "<p><i>Workflow follows the BCSI Project Management process map: PH1 Contract &rarr; PH2 Preconstruction "
          "&rarr; PH3 Construction &rarr; PH4 Closeout.</i></p>")
    pvals = {
        "name": PROJECT_NAME,
        "partner_id": partner["customer"],
        "user_id": users.get("maria", o.uid),
        "label_tasks": "Tasks",
        "description": desc,
        "date_start": "2026-04-13",
        "date": "2027-01-29",
        "allow_milestones": True,
        "allow_task_dependencies": True,
        "allow_recurring_tasks": True,
        "privacy_visibility": "employees",
        "stage_id": pstage[CURRENT_PROJECT_STAGE],
        "tag_ids": [(6, 0, [tags["Operations"], tags["07 51 05 Built-Up Roofing"]])],
        "task_properties_definition": PROPERTIES,
    }
    project = o.upsert("project.project", [("name", "=", PROJECT_NAME)], pvals)
    log(f"Project ready: id={project}")
    fav = o.clean("project.project", {"favorite_user_ids": [(4, o.uid)]})
    if fav:
        o.x("project.project", "write", [project], fav)

    # 7. Task stages for this project (phases)
    tstage = {}
    for i, (n, fold) in enumerate(TASK_STAGES, 1):
        tstage[n] = o.upsert("project.task.type", [("name", "=", n), ("project_ids", "in", [project])],
                             {"name": n, "sequence": i, "fold": fold, "project_ids": [(4, project)]})

    # 8. Milestones
    mil = {}
    for n, dl, reached in MILESTONES:
        mil[n] = o.upsert("project.milestone", [("name", "=", n), ("project_id", "=", project)],
                          {"name": n, "project_id": project, "deadline": dl, "is_reached": reached})

    # 9. Tasks (two passes: create/update, then parents + dependencies)
    tid = {}
    for t in tasks:
        vals = {
            "name": t["name"],
            "project_id": project,
            "stage_id": tstage[t["stage"]],
            "tag_ids": [(6, 0, [tags[x] for x in t["tags"]])],
            "user_ids": [(6, 0, uids(t["users"]))],
            "planned_date_begin": f"{t['start']} 08:00:00",
            "date_deadline": f"{t['end']} 17:00:00" if o.fields("project.task").get("date_deadline", {}).get("type") == "datetime" or isinstance(o, DryRun) else t["end"],
            "priority": t["priority"],
            "sequence": t["sequence"],
            "milestone_id": mil.get(t["milestone"]) if t["milestone"] else False,
            "partner_id": partner["customer"],
        }
        if t["desc"]:
            vals["description"] = t["desc"]
        if t["hours"]:
            vals["allocated_hours"] = t["hours"]
        if t["props"]:
            vals["task_properties"] = t["props"]
        tid[t["key"]] = o.upsert("project.task", [("name", "=", t["name"]), ("project_id", "=", project)], vals)
    for t in tasks:
        vals = {}
        if t["parent"]:
            vals["parent_id"] = tid[t["parent"]]
        if t["deps"]:
            vals["depend_on_ids"] = [(6, 0, [tid[k] for k in t["deps"]])]
        vals = o.clean("project.task", vals)
        if vals:
            try:
                o.x("project.task", "write", [tid[t["key"]]], vals, context=QUIET)
            except Exception as e:
                log(f"  links not saved on {t['name'][:50]}: {str(e).strip().splitlines()[-1][:120]}")
    # states last (dependencies can force "waiting")
    if o.has("project.task", "state"):
        for t in tasks:
            try:
                o.x("project.task", "write", [tid[t["key"]]], {"state": t["state"]}, context=QUIET)
            except Exception as e:  # blocked tasks can refuse "in progress"
                log(f"  state {t['state']} not applied on {t['name'][:50]}: {str(e)[:80]}")
    log(f"Tasks ready: {len(tid)}")

    def safe(label, fn):
        try:
            fn()
        except Exception as e:
            log(f"  skipped {label}: {str(e).strip().splitlines()[-1][:150]}")

    # 10. Recurring weekly status report task
    rec = {
        "name": "Weekly Status Report (WR) - Prepare & Publish Project Update",
        "project_id": project,
        "stage_id": tstage["PH3 - Construction"],
        "tag_ids": [(6, 0, [tags["Weekly Report"], tags["Project Manager"]])],
        "user_ids": [(6, 0, uids(["maria", "kris"]))],
        "date_deadline": "2026-10-02",
        "description": "<p>Every Friday: update Submittal Log, Incoming Docs, Outgoing RFI Log, BCSI Activities, "
                       "Material Procurement and Action Item log, then publish a Project Update.</p>",
        "recurring_task": True,
        "repeat_interval": 1,
        "repeat_unit": "week",
        "repeat_type": "until",
        "repeat_until": "2027-01-29",
    }
    if not o.has("project.task", "date_deadline") or o.fields("project.task").get("date_deadline", {}).get("type") == "datetime":
        rec["date_deadline"] = "2026-10-02 17:00:00"
    try:
        tid["weekly_wr"] = o.upsert("project.task", [("name", "=", rec["name"]), ("project_id", "=", project)], rec)
    except Exception as e:
        rec.pop("recurring_task"), rec.pop("repeat_interval"), rec.pop("repeat_unit"), rec.pop("repeat_type"), rec.pop("repeat_until")
        tid["weekly_wr"] = o.upsert("project.task", [("name", "=", rec["name"]), ("project_id", "=", project)], rec)
        log(f"  recurrence not applied: {str(e)[:80]}")

    # 11. Weekly Status Report as a Project Update
    upd = {
        "name": "Weekly Status Report #01 - 09/30/26",
        "project_id": project,
        "status": "at_risk",
        "progress": 35,
        "date": "2026-09-30",
        "user_id": users.get("maria", o.uid),
        "description": weekly_report_html(tasks),
    }
    safe("project update", lambda: o.upsert("project.update", [("name", "=", upd["name"]), ("project_id", "=", project)], upd))
    log("Project update (Weekly Status Report) ready")

    # 12. Activities (follow-ups) on key items
    todo = o.ref("mail.mail_activity_data_todo")
    task_model = o.x("ir.model", "search", [("model", "=", "project.task")], limit=1)
    acts = [
        ("sub01", "Follow up with Strawn Construction on SUB-01 Sealants review (sent 06/15/26)", "kassandra", "2026-10-01"),
        ("ai01", "Get Owner color selection for Sikaflex-1A", "kris", "2026-10-02"),
        ("pco01", "Price 4-ply BUR + cap sheet and send COR to GC", "maria", "2026-10-02"),
    ]
    if todo and task_model:
        for key, summary, who, due in acts:
            if who not in users:
                continue
            exists = o.x("mail.activity", "search", [("res_model", "=", "project.task"), ("res_id", "=", tid[key]),
                                                     ("summary", "=", summary)], limit=1)
            if not exists:
                safe("activity", lambda key=key, summary=summary, who=who, due=due: o.x("mail.activity", "create", {"res_model_id": task_model[0], "res_id": tid[key],
                                                "activity_type_id": todo, "summary": summary,
                                                "user_id": users[who], "date_deadline": due}, context=QUIET))

    # 13. Attach source documents
    def attach(res_model, res_id, fname):
        p = DOCS_DIR / fname
        if not p.exists():
            log(f"  (skip attachment, not found: {p})")
            return
        if o.x("ir.attachment", "search", [("res_model", "=", res_model), ("res_id", "=", res_id), ("name", "=", fname)], limit=1):
            return
        o.x("ir.attachment", "create", {"name": fname, "res_model": res_model, "res_id": res_id,
                                        "datas": base64.b64encode(p.read_bytes()).decode()})
        log(f"  attached {fname}")

    for t in tasks:
        for f in t["attach"]:
            safe("attachment", lambda t=t, f=f: attach("project.task", tid[t["key"]], f))
    for f in ("PROCESS_FLOW.pptx", "WR_23009_01_Weekly_Status_Report.xlsm"):
        safe("attachment", lambda f=f: attach("project.project", project, f))

    # chatter note on the RFI (logged note, no e-mail)
    def rfi_note():
        note = o.ref("mail.mt_note")
        already = o.x("mail.message", "search", [("model", "=", "project.task"), ("res_id", "=", tid["rfi001"]),
                                                 ("body", "ilike", "RFI 001 answered")], limit=1)
        if not already:
            o.x("project.task", "message_post", [tid["rfi001"]],
                body="RFI 001 answered by Steelhead Engineers: disregard header of par. 1.01 - provide 4-ply built-up "
                     "membrane + cap sheet per par. 3.06.B and detail 1/A10.40. PCO-01 opened to evaluate cost impact.",
                message_type="comment", subtype_id=note)
    safe("RFI chatter note", rfi_note)

    log(f"\nDone. Open: {o.url}/odoo/project/{project}  (or /web#model=project.project&id={project})")
    return project


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="validate data without connecting")
    ap.add_argument("--no-users", action="store_true", help="do not create the BEST team users")
    a = ap.parse_args()
    if a.dry_run:
        o = DryRun()
        build(o, create_users=not a.no_users)
        tasks = build_tasks()
        keys = {t["key"] for t in tasks}
        for t in tasks:
            assert t["parent"] in (None, *keys), t["key"]
            assert all(k in keys for k in t["deps"]), t["key"]
            assert d(t["start"]) <= d(t["end"]), t["key"]
            assert all(x in TAG_COLORS for x in t["tags"]), (t["key"], t["tags"])
            assert all(u in {m["key"] for m in TEAM} for u in t["users"]), t["key"]
        print(f"dry-run OK: {len(tasks)} tasks, {len(o.calls)} RPC calls simulated")
        return
    url = os.environ.get("ODOO_URL", "https://best-contracting-services-demo-msel.odoo.com")
    login, key = os.environ.get("ODOO_LOGIN"), os.environ.get("ODOO_API_KEY")
    if not (login and key):
        sys.exit("Set ODOO_LOGIN and ODOO_API_KEY (see README.md)")
    build(Odoo(url, os.environ.get("ODOO_DB"), login, key), create_users=not a.no_users)


if __name__ == "__main__":
    main()
