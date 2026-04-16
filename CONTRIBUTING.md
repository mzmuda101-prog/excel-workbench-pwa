# Contributing

Thanks for considering contributing to Excel Workbench PWA.

At this stage, the most valuable contributions are not only code. They are also:

- bug reports
- UX feedback
- strange real-world workbook examples
- edge cases that break assumptions about headers, blocks, formulas, or filters

## Best Ways To Help

- open an issue with a clear reproduction path
- describe what you expected and what actually happened
- share a sanitized workbook sample if possible
- suggest workflow improvements, especially for tablet/PWA use

## Before Opening A Pull Request

- keep changes aligned with the product direction in [ROADMAP.md](./ROADMAP.md)
- prefer small, focused changes over broad rewrites
- avoid adding backend requirements unless explicitly discussed first
- keep the app local-first and browser-friendly

## UX Guidelines

This project values:

- human-readable UI
- low-friction workflows
- safe handling of source files
- features that feel useful on both desktop and tablet

If a change is powerful but makes the app harder to understand, it should probably be simplified first.

## Local Run

Serve the repo with any static server, for example:

```bash
python3 -m http.server 8001
```

Then open:

```text
http://127.0.0.1:8001/
```

## Pull Request Scope

Good PRs for this project usually:

- solve one problem clearly
- avoid unrelated formatting churn
- preserve existing workflows
- explain user impact, not only code changes

## If You Are Unsure

Opening an issue first is totally fine.

That is often the best way to check whether an idea fits the project before spending time on a larger contribution.
