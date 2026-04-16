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

### 4. Wide Workbook Support

- improve repeated-block detection across more workbook shapes
- make `Wide-to-Long` clearer and more robust
- improve long-view record labels and structure hints
- support more real-world repeated patterns

Why:
- wide operational spreadsheets are one of the strongest use cases for this project

### 5. Aggregation Workbench `v2`

- stronger manual customization of:
  - grouping
  - measures
  - aggregation rules
  - source/header selection
- better result browsing and result-to-table workflows
- more flexible handling of workbook-specific structures

Why:
- current version is a strong `v1`
- the long-term goal is a browser-native substitute for some macro-based Excel workflows

### 6. Formula Workbench Improvements

- better grouping of similar formulas
- stronger filters by function and error type
- clearer detection of outlier formulas in a column
- more useful formula troubleshooting views

Why:
- formulas are a painful Excel area and a good fit for workbench tooling

### 7. Lightweight Macro-Substitute Workflows

- reusable analysis setups for specific workbook patterns
- saved workbench scenarios per file type
- lightweight local transformations that avoid VBA dependence

Why:
- this is one of the most important long-term product opportunities
- especially for PWA and tablet use

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

## Specific Direction For Aggregation

The current aggregation module should evolve carefully:

- keep the easy `v1` flow for fast questions
- expand manual control without making the default UX heavy
- allow more powerful layouts only when the sidebar stops being enough

That means:

- some controls may stay in the sidebar
- richer result views may eventually move into a wider or dedicated workspace
- UI decisions should follow clarity and comfort, not a rigid layout rule

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
