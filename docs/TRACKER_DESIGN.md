# Production & QC tracker design

## One-screen workflow
Keep the chapter/unit plan on the left, day-wise team output on the right, and a small visual summary within the working view. Core information must remain readable at normal zoom on the agreed display size.

## Single target control
Provide just one editable strip containing Comp SB, Comp WS, QC SB, QC WS, JA/day and working hours/day. All dependent targets must reference those cells—not duplicate hardcoded rates.

## Correct accounting
- Manuscript pages, proof scope and completed pages are distinct.
- Daily employee output must not be silently assigned to chapters without evidence.
- Remaining pages = sum of each unit's MAX(0, planned - completed). Overproduction elsewhere cannot hide backlog.
- Total corrections and verified PE-affected pages are different.
- Flag output above plan, a Done status with outstanding pages and invalid PE counts.

## Usability
Only input cells should look editable; calculated cells should be visually distinguished. Label Comp, QC, Quality %, Pages Left in full. New chapters, staff and dates must not require manual formula repair.

## Visuals
Prefer compact coverage by section, top unfinished chapters and daily Comp/QC trend charts. Avoid unnecessary dashboards. Charts must refer to actual calculated values and update over their advertised input range.

## Compatibility
Target Excel 2010 using supported formulas and styles. Avoid dynamic arrays and modern-only Excel functions. Test the resulting workbook in desktop Excel 2010 before release.
