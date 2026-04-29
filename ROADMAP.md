# Roadmap

This roadmap is public-facing and intentionally split into two groups:

- `Planned`: things that fit the current product direction and are realistic next steps
- `Ideas`: promising directions that should stay open, but are not commitments yet

## Product Principles

- Keep workbook processing local in the browser
- Prefer features that are safe for source files
- Build a workbench around Excel, not a full Excel clone
- Prioritize workflows that are painful in normal Excel, especially on tablet/PWA
- Reuse existing modules instead of growing a pile of special-case tools
- Draw light inspiration from Excel Pivot Tables for aggregation workflows
- Aim to replace common macro/VBA-heavy Excel workflows with lighter, browser-native alternatives

## Current Foundation

Already in the product:

- local file loading for `.xlsx` and `.xlsm`
- text and date filtering
- sorting and working-view workflow
- workbook and sheet inspection
- section navigation
- repeated block detection
- `Wide-to-Long`
- duration analysis
- aggregation workbench `v1`
- formula workbench
- offline-first PWA shell

## Planned

### 1. Better Everyday Filtering And Working Views

- stronger multi-column text filtering
- smoother date filtering workflows
- faster restoring of working views
- continued cleanup of quick-search ergonomics

Why:
- this is core daily value
- this is where the app can beat normal Excel fastest

### 1.5. Puste/Niepuste Filtering Enhancement (FUTURE)

- opcja "sprawdzaj we WSZYSTKICH kolumnach" niezależnie od wybranych kolumn
- jako dodatkowy checkbox: "we wszystkich kolumnach" albo osobne pole wyboru
- pozwoli na scenariusze gdzie użytkownik chce filtrować wiersze całkowicie puste/niepuste w całym arkuszu

Why:
- rozszerza elastyczność filtrowania o nowe scenariusze użycia

### 2. Multi-Sort And Reusable Presets

### 2. Multi-Sort And Reusable Presets

- multi-sort with priority order
- saved filter/sort presets
- quick switching between common working states
- later: export/import of presets

Why:
- it turns filtering into a real workbench workflow instead of a one-off action

### 3. Inspector Refinement

- keep simplifying the sidebar information architecture
- make section navigation, block detection, and column signals feel like one coherent inspector flow
- reduce duplicated signals and unclear labels
- improve header-row guidance and recovery paths

Why:
- this supports every advanced feature added later

### 4. Multi-Level Grouping (Aggregation v2)

- grouping by multiple columns (e.g., Country > City in hierarchy)
- result sorting options
- better result browsing

Note: HAVING filter and distinct count already implemented in v1.5.

Why:
- makes aggregation workbench more universal for various sheet types
- distinct count is a common business question ("how many unique customers?")
- multi-level grouping solves more complex analytical scenarios

### 5. Wide Workbook Support

- improve repeated-block detection across more workbook shapes
- make `Wide-to-Long` clearer and more robust
- improve long-view record labels and structure hints
- support more real-world repeated patterns

Why:
- wide operational spreadsheets are one of the strongest use cases for this project

### 6. Aggregation Workbench `v2`

- stronger manual customization of:
  - grouping
  - measures
  - aggregation rules
  - source/header selection
- better result browsing and result-to-table workflows
- more flexible handling of workbook-specific structures

Why:
- current version is a strong `v1`
- extends the pivot-table-inspired workflow further

### 7. Formula Workbench Improvements

- better grouping of similar formulas
- stronger filters by function and error type
- clearer detection of outlier formulas in a column
- more useful formula troubleshooting views

Why:
- formulas are a painful Excel area and a good fit for workbench tooling

### 8. Lightweight Macro-Substitute Workflows

- reusable analysis setups for specific workbook patterns
- saved workbench scenarios per file type
- lightweight local transformations that avoid VBA dependence

Why:
- this is one of the most important long-term product opportunities
- especially for PWA and tablet use

### 9. Future Localization

- add English as an optional UI language
- make labels, helper text, and workbench panels available in both Polish and English
- improve locale-aware formatting so Polish mode uses Polish-friendly date presentation instead of English month names

Why:
- the project is easier to share publicly when it can be shown in English
- local formatting should feel natural in the language the user actually chose

## Ideas

These are valuable directions, but they are not promises yet.

### KPI Extractor Expansion

- more flexible extraction from dashboards and summary-heavy sheets
- better understanding of labels, aliases, and summary placement

### Cross-Sheet Dependency Explorer

- inspect helper sheets, lookups, and cross-sheet structure
- surface workbook dependencies in a more visual way

### Compare Tools

- lightweight `Key Compare`
- lightweight `Formula Compare`
- only if real usage proves the need

### Workbook Structure Diagnostics

- named ranges
- freeze panes - sticky header when scrolling (like Excel/Google Sheets freeze panes)
- table objects
- hidden / very hidden sheet signals
- structural heatmaps

### Advanced Filters

- regex filtering
- filtering by text length
- filtering by cell type
- filtering by color or formatting hints

### Shared Words Grouping (FUTURE)

- intelligent grouping of similar text values with variations (e.g., "Julian Olsz", "Gr1 Julian Olsz", "Julian Olasz")
- use shortest record as canonical pattern per group
- compare by shared words count with configurable minimum threshold
- noise removal option (numbers, prefixes like "Gr1", "gr.4")
- case-sensitive option
- goal: count "Julian Olsz" variants together while keeping "Julian Olasz" separate

Why:
- handles common data entry variations and typos
- enables accurate counting of entities with inconsistent naming
- currently blocked by complexity issues (needs more robust implementation)

### Local Analytical Sessions

- save local analysis sessions
- restore investigative state later
- keep a light history of workbench actions per file type

### Localization Quality

- stronger locale-aware formatting for dates, labels, and summaries
- more consistent Polish and English wording across the UI

### Python Integration (Exploratory)

- research potential for bridging this PWA application with Python-based processing capabilities
- explore direct or indirect file editing workflows that leverage Python libraries (e.g., pandas, openpyxl)
- potential for running Python scripts locally or via lightweight backend services
- could unlock: advanced data transformations, custom formula processing, automated cleanup routines, cross-file operations, and integration with Python-powered analytical pipelines
- longer-term vision: position the workbench as a bridge between accessible browser-based UI and powerful Python-driven data workflows

Why:
- expands the tool from a workbench into a potential integration hub
- leverages Python's ecosystem strengths while keeping the browser interface accessible
- opens doors for enterprise automation, data science workflows, and custom processing pipelines
- aligns with the macro-substitute opportunity — Python could replace VBA-heavy workflows more elegantly

Status: early exploration, requires technical feasibility assessment.

### Real Workbook Intelligence

Based on the real workbook that inspired this project (`RODO_Obieg_terenow_2026_V5.xlsx`), the app should gradually become better at recognizing multi-layer Excel systems, not just flat tables.

- detect workbook roles automatically:
  data sheet, helper sheet, dashboard sheet, analysis sheet, chart sheet
- detect real Excel tables and expose:
  table name, range, columns, linked formulas, dependent sheets
- detect named ranges and warn about broken names or dead references
- detect chart-bearing sheets and unsupported/partially supported Excel features
- highlight conditional formatting presence as a workbook signal

Why:
- the workbook uses a true multi-sheet architecture with helper logic, dashboarding, and summary layers
- better structural awareness would make the app much more useful for real-world operational Excel files

### Formula Pattern Intelligence

The latest workbook version uses many repeated formula families (`IF`, `COUNTIF`, `COUNTIFS`, `IFERROR`, duration calculations, helper transformations), which makes it a strong source of inspiration for a more advanced formula workbench.

- group repeated formulas by pattern, not only by exact text
- show formula families with counts, affected columns, and sample addresses
- detect outliers inside repeated formula blocks
- surface long / complex formulas as "high-maintenance" candidates
- identify helper formulas that look like business rules:
  status derivation, overdue flags, date-range duration logic, name normalization

Why:
- real files often rely on hundreds or thousands of copied formulas
- understanding formula structure is often more valuable than reading formulas cell by cell

### Process Sheet / SLA Workbench

The inspiring workbook behaves like a process tracker with status logic and aging rules. This suggests a dedicated analysis mode for operational sheets.

- detect status columns such as:
  `W trakcie`, `Zamknięte`, `PRZETERMINOWANY`
- detect "open vs closed" lifecycle rules from paired date columns (`od` / `do`)
- auto-build SLA and aging summaries:
  in progress, closed, overdue, average closure time, longest open items
- add workload views per employee / owner / assignee
- add top bottleneck and longest-duration ranking panels

Why:
- this would match the actual business use case that started the project
- the workbench could become especially strong for operational, legal, administrative, and tracking-style spreadsheets

### Repeated Block Templates

The main sheet in the inspiring workbook uses repeated cycle blocks (`Imię i Nazwisko`, `od`, `do`, `Długość`, then suffixed variants like `2`, `3`, etc.), which should directly inform future upgrades.

- improve repeated-block detection for suffixed headers
- suggest a canonical block schema automatically
- show how many cycles/blocks were found and which columns belong to each block
- offer one-click "analyze as repeated process cycles"
- improve Wide-to-Long suggestions for cyclical Excel layouts

Why:
- this structure appears in real files, not only in synthetic examples
- repeated operational cycles are one of the strongest areas where the workbench can outperform basic spreadsheet viewers

### Workbook Inspiration Loop

Keep using real workbooks that motivated the project as a design source for future features.

- review new workbook versions for:
  new formulas, helper columns, reporting patterns, and sheet architectures
- treat workbook evolution as product research
- when a new manual Excel workaround appears, consider whether it should become a workbench feature

Why:
- the strongest roadmap ideas are coming from real pain points, not abstract brainstorming
- this keeps the product aligned with genuine spreadsheet work instead of drifting into generic BI tooling

## Not In Scope

Things this project should avoid for now:

- pretending to offer full Excel compatibility
- promising macro/VBA parity
- introducing heavy BI-style complexity too early
- adding browser features that weaken trust in local file safety

## Related Notes

Deeper internal notes still live in:

- [NOTES-priority-plan.md](./docs/notes/NOTES-priority-plan.md)
- [NOTES-workbook-patterns-from-real-files.md](./docs/notes/NOTES-workbook-patterns-from-real-files.md)
- [NOTES-module-overlap-audit.md](./docs/notes/NOTES-module-overlap-audit.md)
- [NOTES-filter-press-workbench.md](./docs/notes/NOTES-filter-press-workbench.md)
