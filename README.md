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
* Creates 8 internal users for the BEST team (use `--no-users` to skip). Users get **no e-mail address** and Odoo-inbox notifications, so nobody named in the documents is e-mailed.
* Schedule dates beyond the document dates (RFI 6/2/26, SEI response 6/10/26, transmittal 6/15/26) are illustrative for the demo Gantt.
