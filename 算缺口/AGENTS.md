# Repository Guidelines

## Project Structure & Module Organization
- `src/main.py` is the CLI entry point for shipment planning runs.
- Core business logic lives in `src/shipment_planner/`:
  - `cli.py` (argument parsing and orchestration)
  - `parsers.py` and `xlsx_reader.py` (input validation/parsing)
  - `engine.py` (allocation and recommendation logic)
  - `constraints.py` and `reports.py` (rule loading and exports)
- UI code lives in `src/planner_ui/` with launcher `src/ui_main.py`.
- Runtime data paths are `data/input/` (xlsx/json inputs) and `data/output/` (generated csv/json reports).
- Packaging assets are `build.bat` and `发货建议工具.spec`; binaries are emitted to `dist/`.

## Build, Test, and Development Commands
- `python3 src/main.py --input-dir data/input --out-dir data/output`  
  Runs the planner with auto-detected orders/sales files.
- `python3 src/main.py --orders data/input/<orders>.xlsx --sales data/input/<sales>.xlsx --out-dir data/output`  
  Runs with pinned input files.
- `python3 src/ui_main.py` or `PYTHONPATH=src python3 -m planner_ui`  
  Starts the PyQt6 desktop UI.
- `python3 -m compileall src`  
  Fast syntax validation before opening a PR.
- `build.bat` / `build.bat clean` (Windows)  
  Builds the executable with PyInstaller, optionally cleaning prior artifacts first.

## Coding Style & Naming Conventions
- Use Python with 4-space indentation and readable, PEP 8-aligned formatting.
- Keep type hints on public functions and dataclasses; follow existing `@dataclass(slots=True)` patterns.
- Naming: `snake_case` for variables/functions/files, `PascalCase` for classes, `UPPER_CASE` for constants.
- Preserve Chinese output headers/messages and field names unless a migration is explicitly planned.

## Testing Guidelines
- Use `python3 -m pytest` for the committed automated test suite.
- Use repeatable CLI smoke tests when a change affects file detection, report exports, or end-to-end workflows.
- For each logic change, verify:
  - `data/output/发货建议明细.csv`
  - `data/output/数据质量报告.csv`
  - `data/output/运行摘要.json`
- Add tests under `tests/` with `test_*.py` naming, prioritizing parser normalization and engine allocation edge cases.

## Commit & Pull Request Guidelines
- Recent history mixes terse and descriptive commits; prefer descriptive messages going forward.
- Recommended style: `type(scope): imperative summary` (example: `fix(engine): enforce order-level threshold after keep rule`).
- PRs should include purpose, business impact, commands used for validation, and any output/schema changes.
- For UI or report format changes, attach screenshots or sample output snippets and call out backward-compatibility risks.

## Data & Configuration Hygiene
- Do not commit sensitive production exports or machine-local paths.
- Keep configurable rules in `data/input/shipment_constraints.json` instead of hardcoding constraints.
- Treat `build/`, `dist/`, and `data/output/` as generated artifacts.
