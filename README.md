# AVW Electrical Schedule Generator — Conservative Pilot

This pilot converts a text-based AVW / QuickBooks quote PDF and the AVW Master List workbook into an engineer-reviewable Electrical Schedule draft. A bundled QuickBooks item catalog supplies a controlled second-pass item/description fallback.

Version 0.6.0 keeps the v0.5.1 canonical-code and hierarchy rules, then adds a bundled 4,233-item QuickBooks catalog. Every quote line still checks the uploaded engineering Master first. Only when that Master result is unmatched or unsafe does the app try one unique exact QuickBooks item-number match. The editor also includes a persistent **Add from QuickBooks** search panel with the same placement workflow and two-second description tooltip used by Add from Master.

## Safety rules in this build

- The uploaded Master List always has first priority.
- If the Master cannot produce one safe row, the app tries one unique exact QuickBooks item-number match; description similarity is never used to guess.
- QuickBooks maps `Item Number` to schedule `Part #` and `Description` to schedule `Description`.
- QuickBooks fallback rows are standalone and visibly marked as needing engineering requirements; QuickBooks never supplies HP, phase, volts, amps, breaker, air, or water requirements.
- A QuickBooks fallback keeps the exact quote quantity on one row instead of expanding per-foot or bulk quantities into hundreds of browser/Excel rows.
- Engineering-approved old-to-new code aliases are exact and built in: `CN1 → CN1-3524`, `WA1-SK → WA1-SK-2018`, `WA1-SKP → WA1-SKP-2018`, and `WA1P → WA1-120V`.
- A direct `WA1-SK-2018` call always produces `WA1-SKP-2018` as its direct child, exactly once.
- `PSH25` and its descendants are excluded from the `PWB1` package; `PSH25` remains separately searchable in the Master List.
- An explicitly approved alias file may be supported later; there is no generic before-dash match.
- Truncated item numbers auto-match only when the visible code prefix resolves to one safe Master part number; ambiguous cases stay in Review.
- Replacements and special transformations are never applied automatically.
- Unmatched items always appear in Review.
- Known non-schedule items create no schedule rows and still appear in Review.
- Whole-number equipment quantities repeat the complete parent/child group.
- Different motor specifications share one Motor sequence in the `#` column.
- Extended Amps is intentionally excluded from calculations until engineering defines it.

## Engineer editing

The web app allows engineers to:

- Select one or many rows directly in the schedule grid.
- See parent and standalone rows with a blue background and bold text.
- Switch between all rows and a parent-only view without changing the schedule.
- Move selected rows directly before, after, under, or to the bottom in one action.
- Move a full subtree one step up or down among siblings when a small adjustment is needed.
- Nest or outdent any row while keeping all of its descendants together.
- Insert a complete stored Master assembly, one exact Master component, or a project-only custom row before, after, under, or at the bottom.
- Automatically regenerate project group numbers and nested codes after every structural change.
- Search and replace a row from the exact Master row without losing its tree position or children.
- Edit all engineering requirement fields.
- See whether each requirement row is inherited from the Master, overridden by engineering, or custom; restore inherited values with one click.
- Apply bold, italic, underline, or yellow highlighting to several selected rows at once; formatting is preserved in Excel.
- Delete several selected rows with a confirmation that counts all affected descendants.
- Save editable grid values automatically to the draft.
- Undo the last 20 schedule changes during the current session.
- Record decisions and notes for Review items.
- Select one or more Review rows and add only their missing exact item/description quantity at the top, bottom, before/after a row, or as children of a row.
- Identify Review-added schedule rows with light-red highlighting in both the web draft and exported Excel.
- Hover over draft or Master descriptions for two seconds to see the complete description and available electrical/utility requirements.
- Keep the selected row after a one-step move so repeated up/down movement needs no reselection.
- Keep the Master search panel open through searches, selection, and insertion.
- Search the bundled QuickBooks catalog and insert an exact item before, after, under, at the top, or at the bottom without closing the panel.
- Hover over a QuickBooks result for two seconds to see its complete description and the engineering-requirements warning.
- Export an editable `.xlsx` workbook with four visible sheets.

The editor is deliberately wide and 760 pixels tall. Its action toolbar stays above the grid so engineers can move, add, find, replace, format, delete, or nest rows without moving between separate editing areas.

## Known non-schedule CSV

The optional CSV uses these columns:

```csv
item,reason,enabled
VAC-HOSE-150,Vacuum hose sold by length,true
```

Keep entries exact. Prefixes and wildcards are intentionally unsupported in this pilot.

## Run locally

On Windows, right-click `run_windows.ps1` and choose **Run with PowerShell**, or run:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\run_windows.ps1
```

The manual commands are:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Open the address displayed by Streamlit. Upload the quote PDF and Master List; the QuickBooks catalog is already bundled and indexed by the app. Review the draft, then download the Excel schedule.

## Output workbook

The exported workbook contains:

- Summary
- Electrical Schedule
- Quote Extract
- Review Items

The Electrical Schedule uses the same A:Q column structure, grouped headers, parent bolding, nested codes, and print-oriented layout as the supplied engineering schedule. Part-number cells include an Excel comment showing the exact Master source row or QuickBooks source row and the requirement status.

## Current pilot limits

- The engineering Master List remains the only source of electrical and utility requirements; QuickBooks fallback rows require engineering review.
- Items that are unresolved by both Master and QuickBooks, plus replacement/special quote items, remain in Review instead of being guessed.
- Extended Amps stays intentionally blank until engineering confirms its definition.
- The PDF reader expects text-based QuickBooks PDFs; scanned/image-only quotes require OCR before production rollout.
