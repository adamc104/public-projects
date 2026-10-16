# Medicaid enrollment tracker

Pulls monthly Medicaid enrollment from 11 state websites into one Excel workbook, with MoM and QoQ growth.

## Run it

Open a terminal in this folder and run:

```
.venv\Scripts\python tracker.py
```

Or just ask Claude Code: "run the Medicaid tracker".

- A Chrome window opens and clicks through some state sites. **Leave it alone until it closes.** Kansas, New Hampshire, Minnesota, Kentucky and North Carolina block hidden browsers.
- Takes about 1–2 minutes. Only new files get downloaded; everything already fetched is kept in `downloads/`.
- To run only some states: `.venv\Scripts\python tracker.py Florida "New York"`

## What you get

- **`Medicaid Enrollment Tracker.xlsx`**: four sheets, Enrollment, MoM growth, QoQ growth and Notes.
- **`data/enrollment.csv`**: every data point with its source. This is the master record.
- The terminal prints any **new months** since the last run, plus a **reference check** against the reference historicals (`reference.csv`). If a state's website changes and the tracker starts pulling the wrong row, the check fails loudly.

## Where each number comes from

| State | Source | Row used |
|---|---|---|
| Florida | Monthly enrollment Excel (AHCA) | "FLORIDA_MEDICAID ENROLLMENT TOTAL" |
| Kansas | Medical Assistance Report PDF, one per fiscal year | First "Total" row |
| Kentucky | Monthly Membership Counts by County PDF | "Sum:" |
| Minnesota | Managed care enrollment Excel | "Total Enrollment" |
| Mississippi | **Manual**: site down, use `manual_entries.csv` | |
| Missouri | Caseload Counter history Excel | "MO HealthNet Enrollees" |
| New Hampshire | Medicaid Enrollment Data PDF | "Grand Total Full Medicaid" |
| New York | Monthly managed care enrollment Excel | Medicaid Managed Care Table, "STATE TOTAL" |
| North Carolina | Monthly Enrollees Excel, then enrollment dashboard (OCR) for newer months | "MEDICAID + MED. EXPANSION" / "Total Enrollment" tile |
| West Virginia | MCO monthly enrollment PDF, one per year | Last "Total" row |
| Wisconsin | BadgerCare Plus statewide PDF | "Total" |

## Manual numbers

Add a row to `manual_entries.csv` (state, month as `YYYY-MM`, number, note). Manual numbers only fill gaps; a number from a website always wins.

## If something breaks

The terminal names the state that failed. Details go to `data/errors.log`. Other states still update. Ask Claude Code to "fix the <state> reader in the Medicaid tracker".
