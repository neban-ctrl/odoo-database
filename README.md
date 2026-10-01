# BCSI demo – Ruskin ES K1 & K2 in Odoo Project

`build_ruskin_project.py` builds one project in the Odoo **Project** app, based on the client's documents:

| Client document | Where it lands in Odoo Project |
|---|---|
| PROCESS_FLOW.pptx | Task stages **PH1 Contract → PH2 Preconstruction → PH3 Construction → PH4 Closeout**, one task per process-flow box, department/role tags (Sales, Admin, Legal, Operations, PM, PE, Super, Field Mgr), task dependencies, milestones, project stages |
| Weekly Status Report (.xlsm) | Project description (info, BEST team, collaboration tools, contract), **task properties** that mirror the log columns (Spec, Par., Review Status, Date Sent, Due, Response, Scope/Cost Impact, BIC…), sub-task logs (Submittals, RFIs, Change Orders, Material Procurement, Action Items, BCSI Activities), a recurring weekly-report task and a **Project Update** "Weekly Status Report #01" |
| SUB – 07 90 05 – Sealants.pdf | Task **SUB-01** under *Complete Submittal Process* (transmittal data, spec checklist, Sikaflex-1A data, PDF attached) and procurement/action items |
| RFI 001 – Built-Up Roof Assembly.pdf | Task **RFI 001** under *RFIs* (question, SEI response, PDF attached) and **PCO-01** under *Change Order* |

## Run

```bash
export ODOO_URL=https://best-contracting-services-demo-msel.odoo.com
export ODOO_LOGIN=<login e-mail>
export ODOO_API_KEY=<api key>
# optional: export ODOO_DB=<database name>   (auto-detected otherwise)
python3 build_ruskin_project.py --dry-run   # offline validation
python3 build_ruskin_project.py             # build (safe to re-run; updates in place)
```

Put the four source files in `./source_docs/` (git-ignored) with the names used in the script to have them attached.

Notes
* Creates 8 internal users for the BEST team (use `--no-users` to skip). Logins use the reserved non-deliverable domain `@bcsi-demo.example` and Odoo-inbox notifications, so nobody named in the documents is e-mailed.
* Schedule dates beyond the document dates (RFI 6/2/26, SEI response 6/10/26, transmittal 6/15/26) are illustrative for the demo Gantt.

## Workflow configuration (`--workflow`)

Run after the base build. Adds what makes the BCSI process flow move in a demo:

1. **Phase gates**: tasks tagged *Phase Gate* (Execute Contract, PH1A meeting, field work, archive) advance the project PH1 → PH2 → PH3 → PH4 → Complete when marked done (automation rule).
2. **Ruskin settings**: timesheets, portal sharing, e-mail alias `ruskin-k1k2@…` (incoming e-mails become tasks), Documents folder.
3. **Activity types**: RFI Response Due (+7), Notify Client (+5), Submittal Follow-up (+14), Pre-Job Meeting.
4. **Activity plans**: PH1 Contract Handoff, PH2 Preconstruction, Submittal Package, RFI, PH4 Closeout.
5. **Project template** "BCSI Job Template": the full PH1–PH4 process without job data.
6. **Automation rules**: phase gates; new RFI / submittal / incoming doc tasks get their follow-up activity automatically.
7. **CRM**: bidding pipeline (Bid Invitation → Estimating → 80% → 90% → Proposal → Won), Ruskin opportunity (won), a demo bid.
8. **Sales**: "BCSI Roofing Contract" product that creates a project from the template, invoiced by milestones.
9. **Purchase**: Sikaflex-1A, BUR felts, cap sheet, sheet metal products; draft RFQ to Sika linked to Ruskin.
10. **Employees & Planning**: planning roles, employees for the team, Ruskin shifts (master field resource schedule).
11. **Documents**: Ruskin folder structure (replaces SharePoint).

```bash
wget -qO- https://raw.githubusercontent.com/neban-ctrl/odoo-database/<commit>/build_ruskin_project.py | python3 - --workflow
```
