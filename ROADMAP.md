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
- distinct count (number of unique values)
- group filtering with HAVING (e.g., "only categories with >10 products")
- result sorting options
- better result browsing

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
- freeze panes
- table objects
- hidden / very hidden sheet signals
- structural heatmaps

### Advanced Filters

- regex filtering
- filtering by text length
- filtering by cell type
- filtering by color or formatting hints

### Local Analytical Sessions

- save local analysis sessions
- restore investigative state later
- keep a light history of workbench actions per file type

### Localization Quality

- stronger locale-aware formatting for dates, labels, and summaries
- more consistent Polish and English wording across the UI

## Not In Scope

Things this project should avoid for now:

- pretending to offer full Excel compatibility
- promising macro/VBA parity
- introducing heavy BI-style complexity too early
- adding browser features that weaken trust in local file safety

## Related Notes

Deeper internal notes still live in:

- [NOTES-priority-plan.md](./NOTES-priority-plan.md)
- [NOTES-workbook-patterns-from-real-files.md](./NOTES-workbook-patterns-from-real-files.md)
- [NOTES-module-overlap-audit.md](./NOTES-module-overlap-audit.md)
- [NOTES-filter-press-workbench.md](./NOTES-filter-press-workbench.md)
