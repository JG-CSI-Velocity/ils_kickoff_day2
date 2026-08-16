# feat: Improve ILS Kickoff Tool Usability and Output Quality

## Enhancement Summary

**Deepened on:** 2026-02-06
**Sections enhanced:** 10
**Research agents used:** kieran-python-reviewer, architecture-strategist, performance-oracle, security-sentinel, code-simplicity-reviewer, pattern-recognition-specialist, best-practices-researcher (x4: Pydantic config, Plotly PPTX, Streamlit UX, Excel formatting)

### Key Improvements

1. **URGENT: Security remediation** -- Public GitHub repo contains 2.3MB of client banking data; `.gitignore` and repo privacy must be Phase 0
2. **Architecture corrections** -- Rename `src/` to `ils_kickoff/`; add `pipeline.py` orchestrator; decouple charts from `AnalysisResult`; separate chart creation from export
3. **Performance-critical fixes** -- Pin `kaleido==0.2.1` (50x faster than v1); single ExcelWriter context (saves 3-7s); Streamlit lazy rendering to avoid 15-chart UI freeze
4. **Code duplication elimination** -- Extract `grouped_summary()` and `binned_summary()` templates to replace 15 copy-pasted analysis blocks
5. **Simplified config** -- Use plain Pydantic `BaseModel` + `yaml.safe_load()` instead of `pydantic-settings`; drop 4 unnecessary dependencies

### New Considerations Discovered

- Notebook charts are matplotlib/seaborn, NOT Plotly -- chart code is new development, not a port
- Grand Total is a presentation concern; keep it out of analysis DataFrames
- `st.tabs` does NOT lazy-load in Streamlit -- use `st.segmented_control` + `st.fragment` for conditional rendering
- SHA-1 hashed account numbers in CSV are brute-forceable (8-12 digit keyspace); data is not safely de-identified
- `add_grand_total()` function at line 104-115 is defined but never called; 15 inline copies exist instead

### Simplicity Considerations

The simplicity review flagged that this plan proposes significant architectural expansion for what is currently a single-purpose batch script. The recommendations below represent a balanced position: enough structure to make the tool maintainable and extensible for new clients, without over-engineering into an enterprise framework. Key simplifications incorporated:

- Drop `pydantic-settings`, `jinja2`, `matplotlib` from dependencies (7 deps instead of 13)
- Skip `AnalysisResult` base class hierarchy -- use plain functions with dataclass returns
- Config needs only 3 user-facing fields; advanced settings are commented-out defaults
- Streamlit (Phase 4) remains optional and should only be built if CLI proves insufficient

---

## Overview

Restructure the ILS Kickoff banking data analysis project from a hardcoded, developer-only Python script into a configurable, non-technical-user-friendly tool with improved chart/table output and a foundation for template-based PowerPoint generation.

**Current state:** Two competing implementations (798-line script + 147-cell Jupyter notebook) with hardcoded file paths, no dependency management, no documentation, mismatched filenames, and tables-only output.

**Target state:** A single consolidated tool with config-file-driven execution, CLI and optional Streamlit web UI, Plotly charts alongside styled tables, formatted Excel reports, and improved PowerPoint decks.

## Problem Statement / Motivation

Non-technical analysts cannot run the tool today because:
1. The script filename at `ils_kickoff.py:47` (`1774_INB_OD_Tran_Combo_20260203__1_.csv`) does not match the actual CSV (`1774_INB_OD.Tran.Combo_20260203 (1).csv`) -- the script crashes before any analysis begins
2. There is no `requirements.txt` -- users get `ModuleNotFoundError` with no guidance
3. There is no README, no setup instructions, no CLI interface
4. All configuration (file path, client ID, output names) is hardcoded in source code
5. The script produces tables-only PowerPoint and unformatted Excel -- the notebook has charts and formatting but is broken (column name mismatches, undefined variables, syntax errors)

When a consultant needs to run the same analysis for a different client, they must edit Python source code and rename files manually.

## Proposed Solution

### Architecture

```
ils_kickoff_day2/
  config.yaml                # User-editable configuration (primary interface)
  config.example.yaml        # Committed template with documentation
  requirements.txt           # Pinned dependencies (base)
  requirements-dev.txt       # Dev deps: pytest, ruff, mypy
  pyproject.toml             # Project metadata, tool configs, entry points
  .gitignore                 # Exclude data/, output/, *.csv, __pycache__
  README.md                  # Setup + usage for non-technical users
  Makefile                   # make setup, make run, make test, make lint, make clean
  setup.bat                  # Windows double-click setup
  run.bat                    # Windows double-click run
  ils_kickoff/               # <-- renamed from src/ for meaningful imports
    __init__.py
    __main__.py              # Enables `python -m ils_kickoff`
    settings.py              # Pydantic BaseModel + yaml.safe_load (NOT pydantic-settings)
    data_loader.py           # CSV/Excel loading + schema validation
    column_map.py            # Canonical column naming + alias resolution
    formatting.py            # Shared value formatting (column-name heuristic)
    exceptions.py            # Custom exception hierarchy
    pipeline.py              # Orchestrator shared by CLI and Streamlit
    analyses/
      __init__.py            # run_all_analyses() with registry
      base.py                # AnalysisResult dataclass, safe_pct, add_grand_total
      templates.py           # grouped_summary(), binned_summary() -- shared logic
      account_status.py      # Analyses 1-2
      deposits.py            # Analyses 3-4
      nsf_stratification.py  # Analyses 5-10
      od_status.py           # Analyses 11-12
      reg_e.py               # Analyses 13, 15
      od_limits.py           # Analysis 14
    charts/                  # <-- subpackage, not single file (exceeds 300 lines)
      __init__.py            # create_charts() dispatcher + CHART_REGISTRY
      account_status.py
      deposits.py
      nsf.py
      od_status.py
      reg_e.py
    excel_report.py          # Formatted Excel output with NamedStyles
    pptx_report.py           # PowerPoint generation with tables + chart images
    cli.py                   # Typer CLI entry point
  app.py                     # Streamlit web UI (optional, secondary)
  data/                      # Client data files (gitignored)
  output/                    # Generated reports (gitignored)
  templates/                 # Future: PPTX templates
  legacy/                    # Archived notebook
  tests/
    conftest.py              # sample_df, sample_config, edge case fixtures
    test_settings.py
    test_data_loader.py
    test_column_map.py
    test_cli.py              # Typer CliRunner tests
    analyses/
      test_account_status.py
      test_deposits.py
      test_nsf_stratification.py
      test_od_status.py
      test_reg_e.py
      test_od_limits.py
    test_excel_report.py
    test_pptx_report.py
    data/
      sample.csv             # Synthetic 100-row test fixture
```

### Research Insights: Architecture

**From architecture-strategist and pattern-recognition-specialist:**

- **Package naming:** `src/` as an importable package name is an anti-pattern. `from src.settings import Config` is meaningless; `from ils_kickoff.settings import Config` is self-documenting. Also avoids collision with the "src layout" convention where `src/` is non-importable.
- **Pipeline orchestrator:** Both CLI and Streamlit need the same load -> analyze -> chart -> export sequence. Without `pipeline.py`, this gets duplicated. The pipeline also provides the hook for Streamlit progress callbacks.
- **Chart/analysis decoupling:** `AnalysisResult` must NOT contain `go.Figure`. Analysis logic should be testable without Plotly installed. Charts are built in a second pass by mapping `analysis_name -> chart_builder_fn`.
- **Grand Total as presentation concern:** Grand Total rows should not be in analysis DataFrames. They confuse chart functions (which must filter them out) and break percentage calculations. Store as a separate field or compute at render time.
- **Chart export separation:** Separate `create_charts()` (pure Figure creation, no I/O) from `export_chart_png()` (materialization to bytes). Use a `temp_chart_images` context manager for PPTX/Excel embedding.
- **Custom exceptions:** Define `DataLoadError`, `ColumnMismatchError`, `AnalysisError` in `exceptions.py` so both CLI and Streamlit can catch and display them appropriately.
- **Analysis templates:** The 15 analyses decompose into 2-3 parameterized templates (`grouped_summary`, `binned_summary`, `pivot_summary`) that eliminate 440+ lines of duplicated groupby/agg/percent/grandtotal code.

### Source of Truth Decision

Use the **script's analysis logic** (cleaner, consolidated, all 15 analyses working) and **create new Plotly charts from scratch** (the notebook uses matplotlib/seaborn, not Plotly -- this is new development, not a port). Merge in the **notebook's Excel formatting** (cells 13-14). The notebook is archived to `legacy/`.

### Config Schema

```yaml
# config.yaml
# ===========================================================
# ILS Kickoff Analysis Configuration
# ===========================================================
# HOW TO USE:
#   1. Copy config.example.yaml to config.yaml
#   2. Change "data_file" to point to your CSV
#   3. Run: make run  (or: python -m ils_kickoff)
# ===========================================================

# ---- REQUIRED ----
data_file: "data/REPLACE_WITH_YOUR_FILENAME.csv"

# ---- OPTIONAL (defaults work for most clients) ----
# client_id: "1774"          # Auto-extracted from filename if omitted
# client_name: "Client 1774"  # Used in report titles
output_dir: "output/"

outputs:
  excel: true
  powerpoint: true
  html_charts: false

# ---- ADVANCED (you probably don't need to change these) ----
# pptx_template: "templates/company_template.pptx"
# charts:
#   theme: "plotly_white"
#   colors: ["#1B365D", "#4A90D9", "#7BC67E", "#F5A623", "#D0021B", "#8B572A"]
#   width: 900
#   height: 500
# nsf_bins: [0, 1, 2, 3, 4, 5, 10, 15, 20, 50, 100]
# deposit_bins: [0, 500, 1000, 2500, 5000, 10000, 25000, 50000]
```

### Research Insights: Configuration

**From best-practices-researcher (Pydantic config):**

- **Use plain `BaseModel` + `yaml.safe_load()`, NOT `pydantic-settings`.** pydantic-settings was designed for env-var-driven apps. Its YAML support requires overriding `settings_customise_sources()`, has open bugs around silently ignored values (issue #660), and adds hidden env-var resolution that confuses non-technical users. A `Settings.from_yaml(**cli_overrides)` classmethod is simpler, explicit, and debuggable.
- **Config layering:** `CLI args > config.yaml > hardcoded defaults`. No environment variables (analysts don't use them).
- **Three-section config layout:** REQUIRED / OPTIONAL / ADVANCED -- users scan for what they must change.
- **Sentinel value:** Use `REPLACE_WITH_YOUR_FILENAME.csv` as default -- makes it obvious what needs changing, fails validation with a clear message.
- **Friendly error messages:** Catch `ValidationError` and translate to plain English. Never show raw Pydantic tracebacks to non-technical users.
- **Model pattern:** Use `@field_validator` for `data_file` existence check, `@model_validator(mode="after")` for `client_id` derivation from filename.
- **Mature tools reference:** dbt, Great Expectations, Cookiecutter Data Science all use YAML for user-facing config with example files committed and actual config gitignored.

## Technical Approach

### Phase 0: Security Remediation (URGENT -- Do Before Any Other Work)

**From security-sentinel: The repository is PUBLIC on GitHub with client banking data committed. This is an active data exposure that must be addressed immediately.**

- [ ] **Make the GitHub repository PRIVATE** via Settings > Danger Zone > Change visibility
- [ ] **Remove CSV from git history** using `git filter-repo --path "1774_INB_OD.Tran.Combo_20260203 (1).csv" --invert-paths` then `git push --force origin main`
- [ ] **Contact GitHub Support** to purge CDN caches of the CSV
- [x] **Create `.gitignore`** with comprehensive exclusions (see below)
- [ ] **Assess breach notification obligations** -- CSV contains SHA-hashed account numbers (brute-forceable), financial balances, deposit amounts, OD limits, and account open dates. Consult with client's compliance team re: GLBA and state notification laws.

```gitignore
# Client data -- NEVER commit
data/
*.csv
*.tsv

# Generated reports -- contain client data
output/
*.xlsx
*.pptx
*.png
*.html

# Runtime config (contains file paths and client identifiers)
config.yaml

# Python
__pycache__/
*.pyc
.venv/
venv/
*.egg-info/
dist/
build/

# Environment and secrets
.env
.env.*

# OS artifacts
.DS_Store
Thumbs.db

# IDE
.idea/
.vscode/
*.swp
```

**Files:** `.gitignore`

### Phase 1: Foundation (Fix, Structure, Configure)

Tasks that unblock everything else. No new features, just making the existing analysis runnable by non-technical users.

#### 1.1 Fix critical bugs and add project scaffolding

- [x] Fix filename mismatch at `ils_kickoff.py:47` -- change `DATA_PATH` to match actual CSV name
- [x] Create `requirements.txt` with pinned versions (base): `pandas~=2.2`, `numpy~=1.26`, `python-pptx~=1.0`, `openpyxl~=3.1`, `plotly~=6.0`, `kaleido==0.2.1`, `pyyaml~=6.0`, `pydantic~=2.6`, `typer[all]~=0.12`
- [x] Create `requirements-dev.txt`: `pytest~=8.0`, `pytest-cov~=5.0`, `ruff~=0.4`
- [x] Create `pyproject.toml` with project metadata, `[project.scripts]` entry point, and tool configs (ruff, pytest, mypy)
- [x] Create `README.md` with Quick Start (3 steps), Detailed Setup, Configuration Guide
- [x] Create `Makefile` with targets: `setup`, `run`, `test`, `lint`, `fmt`, `clean`
- [x] Create `setup.bat` and `run.bat` for Windows users
- [x] Move CSV to `data/` directory, archive notebook to `legacy/`
- [x] Create synthetic `tests/data/sample.csv` (100 rows -- see fixture requirements below)

### Research Insights: Dependencies

**From kieran-python-reviewer and code-simplicity-reviewer:**

**Removed from dependencies:**
- `matplotlib` -- Plan replaces it with Plotly. Don't carry two charting libraries.
- `pydantic-settings` -- Plain Pydantic + `yaml.safe_load()` is sufficient (see config research above).
- `jinja2` -- Never referenced in any task description. Add in Phase 5 if needed.
- `streamlit` -- Phase 4 dependency only. Not in base requirements.

**Added to dependencies:**
- `lxml` (optional) -- openpyxl uses it for 2x faster XML processing if available.

**Pinning strategy:** Use `~=` (compatible release) for reproducible installs. Exception: `kaleido==0.2.1` exact pin due to v1.0+ performance regression.

**Kaleido v1.0+ has a 50x performance regression** (2000ms+ per chart vs 120-400ms for v0.2.1). Pin `kaleido==0.2.1` until the upstream regression is resolved. See [plotly/Kaleido#400](https://github.com/plotly/Kaleido/issues/400).

**Test fixture `sample.csv` must include:**
- At least one row where `Total Items = 0` (division-by-zero in pay ratio)
- Rows with both `Business Flag = "P"` and `"B"`
- At least one row with `Account Status != "O"` (closed account)
- A `DepositCount` value with embedded tabs
- A row with `NaN` in `Returned Items`, `Open Date`, and `Reg E Flag`
- A `Year Opened` value that is current year (to test dynamic year-bin fix)

**Files:** `.gitignore`, `requirements.txt`, `requirements-dev.txt`, `pyproject.toml`, `README.md`, `Makefile`, `setup.bat`, `run.bat`

#### 1.2 Build configuration system

- [x] Create `ils_kickoff/settings.py` -- Pydantic `BaseModel` with `from_yaml()` classmethod
  - `data_file: Path` (required, validated to exist with supported extension)
  - `client_id: str | None` (auto-extracted from filename via regex `r'^(\d+)'`)
  - `client_name: str | None` (defaults to `"Client {client_id}"`)
  - `output_dir: Path` (default `"output/"`)
  - `outputs: OutputConfig` (excel, powerpoint, html_charts booleans)
  - `charts: ChartConfig` (theme, colors, width, height)
  - `nsf_bins`, `deposit_bins` with current hardcoded values as defaults
- [x] Use `@field_validator` for `data_file` existence + extension check
- [x] Use `@model_validator(mode="after")` for `client_id` derivation from filename
- [x] Wrap `ValidationError` with user-friendly error messages (no raw tracebacks)
- [x] Create `config.example.yaml` with three-section layout (REQUIRED/OPTIONAL/ADVANCED)
- [x] Auto-create `output_dir` if it does not exist

**Files:** `ils_kickoff/settings.py`, `config.example.yaml`

#### 1.3 Build data loading and validation

- [x] Create `ils_kickoff/column_map.py` -- canonical column names + alias resolution
  - Define `REQUIRED_COLUMNS` set with canonical names (from script's `RENAME_MAP` target names)
  - Define `COLUMN_ALIASES` dict merging script's `RENAME_MAP` and notebook's `to_canonical()` aliases
  - Normalize column name "# of Paid Items" vs "# of Items Paid" to single convention
  - `resolve_columns(df) -> df` function that renames columns to canonical names
- [x] Create `ils_kickoff/data_loader.py`
  - Load CSV/Excel based on file extension
  - Handle `DepositCount` tab-embedded data (split on tab, take first value, log warning)
  - Apply column resolution from `column_map.py`
  - Validate all required columns present -- raise `ColumnMismatchError` with available vs expected column names
  - Apply `pd.to_numeric(errors="coerce")` to numeric columns
  - Use `.loc[]` assignments (never chained assignment) to avoid `SettingWithCopyWarning`
  - Validate resolved path is within allowed base directory (path traversal protection)
  - Return validated DataFrame
- [x] Create `ils_kickoff/exceptions.py` -- `DataLoadError`, `ColumnMismatchError`, `AnalysisError`, `ConfigError`

**Files:** `ils_kickoff/column_map.py`, `ils_kickoff/data_loader.py`, `ils_kickoff/exceptions.py`

#### 1.4 Build CLI entry point and pipeline

- [x] Create `ils_kickoff/pipeline.py` -- single orchestrator for CLI and Streamlit:
  ```python
  @dataclass
  class PipelineResult:
      settings: Settings
      df: pd.DataFrame
      analyses: list[AnalysisResult]
      charts: dict[str, go.Figure]  # populated lazily

  def run_pipeline(settings, on_progress=None) -> PipelineResult: ...
  def export_outputs(result: PipelineResult) -> list[Path]: ...
  ```
- [x] Create `ils_kickoff/cli.py` using Typer
  - Default command (no subcommand needed for single-purpose tool):
    `python -m ils_kickoff data/client.csv`
    `python -m ils_kickoff data/client.csv --output ./reports/ --client-id 1774`
  - Use `typer.Option(default=None)` for all overridable fields; merge non-None values into Settings
  - CLI args override config.yaml values via `Settings.from_yaml(**overrides)`
  - `--help` shows clear descriptions for each option
  - `--verbose` enables DEBUG logging
  - Catch custom exceptions and print friendly messages (not tracebacks)
- [x] Create `ils_kickoff/__main__.py` so `python -m ils_kickoff` works
- [x] Replace all `print()` with `logging` from Phase 1 onward

**Files:** `ils_kickoff/pipeline.py`, `ils_kickoff/cli.py`, `ils_kickoff/__main__.py`

### Phase 2: Analysis Refactoring

Extract 15 analyses from monolithic script into modular functions. No behavior changes -- same logic, same outputs, just organized.

#### 2.1 Create base analysis infrastructure and templates

- [x] Create `ils_kickoff/analyses/base.py`
  - `AnalysisResult` dataclass: `name: str`, `title: str`, `df: pd.DataFrame` (NO chart field)
  - Optional: `grand_total: dict | None` -- precomputed Grand Total as separate field (presentation concern)
  - Optional: `error: str | None` -- for graceful degradation when an analysis is skipped
  - Helper: `safe_percentage(numerator, denominator) -> float` -- division with zero guard
  - Helper: `add_grand_total(df, label_col, sum_cols, pct_cols) -> df` -- working version that returns DataFrame (not dict)
  - Helper: `format_number(val, fmt) -> str` -- consistent number formatting
- [x] Create `ils_kickoff/analyses/templates.py` -- parameterized analysis templates:
  - `grouped_summary(df, group_col, agg_specs, pct_of, pay_ratio_cols)` -- handles 11 of 15 analyses
  - `binned_summary(df, value_col, bins, labels, bin_name, agg_specs)` -- handles binned analyses 3-10, 14
  - Each template encapsulates: filter -> groupby -> agg -> compute percentages -> add grand total
- [x] Create `ils_kickoff/formatting.py` -- shared value formatting used by both PPTX and Excel reports
  - Column-name-based type inference (`"%" in col_name`, `"Ratio" in col_name`, etc.)
  - Extracted from current `format_ppt_table` at `ils_kickoff.py:118-168`

### Research Insights: Code Patterns

**From pattern-recognition-specialist:**

Three major duplication patterns in the current script affecting 15 analyses:

| Pattern | Occurrences | Lines Affected |
|---|---|---|
| Filter, copy, bin, groupby, agg | 10 times | 296-547 |
| Compute totals, percentages, pay ratio | 13 times | 239-600 |
| Build Grand Total row as dict, concat | 15 times | 249-649 |

The `grouped_summary()` and `binned_summary()` templates reduce ~440 lines of analysis code to ~100 lines of template logic + 15 small configuration wrappers (~15 lines each) = ~325 lines total.

**`add_grand_total` at line 104-115 is defined but never called.** Every analysis builds Grand Total inline. The function also returns a dict (wrong) instead of appending to the DataFrame. The new version must return the DataFrame with the total row appended.

**Column name inconsistency:** "# of Items Paid" (analyses 1, 7) vs "# of Paid Items" (analysis 2). Normalize to "# of Items Paid" everywhere.

**Files:** `ils_kickoff/analyses/base.py`, `ils_kickoff/analyses/templates.py`, `ils_kickoff/formatting.py`

#### 2.2 Extract analyses into modules

For each analysis module, extract the logic from `ils_kickoff.py` at the specified line ranges. Each function takes a DataFrame and config, returns an `AnalysisResult`. Use the templates from 2.1 to eliminate duplication.

- [x] `ils_kickoff/analyses/account_status.py` -- Analysis 1 (lines 229-258), Analysis 2 (lines 260-291)
- [x] `ils_kickoff/analyses/deposits.py` -- Analysis 3 (lines 294-316), Analysis 4 (lines 319-339)
- [x] `ils_kickoff/analyses/nsf_stratification.py` -- Analyses 5-10 (lines 342-513)
- [x] `ils_kickoff/analyses/od_status.py` -- Analyses 11-12 (lines 516-563)
- [x] `ils_kickoff/analyses/reg_e.py` -- Analysis 13 (lines 566-576), Analysis 15 (lines 615-666)
- [x] `ils_kickoff/analyses/od_limits.py` -- Analysis 14 (lines 579-612)
- [x] `ils_kickoff/analyses/__init__.py` -- `run_all_analyses(df, config) -> list[AnalysisResult]`
  - Create filtered DataFrames once (personal_open, business_open) and pass to analysis functions
  - Do NOT `.copy()` unless the analysis mutates the DataFrame (e.g., `pd.cut()`)
  - Wrap each analysis in try/except; log error and continue on failure
  - Support optional `include`/`exclude` sets for selective execution
  - Consider decorator-based registry pattern for extensibility

**Bug fixes during extraction:**
- Fix year bin upper bound at `ils_kickoff.py:624-625`: change `2025` to `datetime.now().year`
- Add division-by-zero guards to ALL percentage calculations (currently inconsistent)
- Remove `warnings.filterwarnings("ignore")` -- ban global suppression; use targeted `warnings.catch_warnings()` only where explicitly needed
- Remove `pd.options.mode.chained_assignment = None` -- fix code to use `.loc[]` instead
- Fix column name inconsistency: normalize "# of Paid Items" to "# of Items Paid"

**Tests:** Write tests for each analysis module in Phase 2. Each extraction should have corresponding test coverage before being considered complete.

**Files:** `ils_kickoff/analyses/*.py`

### Phase 3: Improved Output

The core value-add. Better charts, better tables, better reports.

#### 3.1 Add Plotly charts for each analysis

- [x] Create `ils_kickoff/charts/` subpackage (single `charts.py` would exceed 300 lines with 15+ chart functions)
  - `__init__.py`: `CHART_REGISTRY` dict mapping analysis names to builder functions; `create_charts(results, config) -> dict[str, go.Figure]` dispatcher
  - One chart module per analysis group, matching the notebook's visual intent but written from scratch in Plotly (notebook uses matplotlib, NOT Plotly):
    - Analysis 1: Dual-axis bar (accounts) + line (pay ratio) by Account Status
    - Analysis 1 alt: Stacked bar (Paid vs Returned items) by Account Status
    - Analysis 2: Grouped bar (Personal vs Business by Open/Closed)
    - Analyses 3-4: Horizontal bar (deposit distribution)
    - Analyses 5-6: Horizontal bar (NSF volume stratification)
    - Analyses 7-8: Bar + line (NSF pay ratio by bracket)
    - Analyses 9-10: Multi-metric bar (NSF + deposits + swipes)
    - Analyses 11-12: Bar (OD status stratification)
    - Analysis 13: Simple bar (Reg E opt-in summary)
    - Analysis 14: Bar (OD limit stratification)
    - Analysis 15: Grouped bar (Reg E by year opened)
  - Set global Plotly defaults: `plotly_white` theme, brand colors from config
  - Each function takes `(df: pd.DataFrame, config: ChartConfig) -> go.Figure` -- no I/O
  - Separate export module: `render_chart_png(fig, config) -> bytes` using `fig.to_image(format="png", width=900, height=500, scale=3)` for ~216 DPI
  - Use `BytesIO` in-memory pipeline for PPTX/Excel embedding -- no temp files

### Research Insights: Plotly to PPTX Pipeline

**From best-practices-researcher (Plotly PPTX):**

- **In-memory pipeline:** `fig.to_image(format="png", scale=3)` returns bytes -> wrap in `BytesIO` -> pass to `slide.shapes.add_picture()`. No temp files, no cleanup.
- **Optimal dimensions:** Layout 900x500px with `scale=3` gives 2700x1500px PNG at ~216 DPI -- sharp for print/PDF export. Each chart PNG ~200-400KB, 15 charts adds ~3-6MB to PPTX.
- **Chart position on slide:** Upper portion at `top=Inches(1.0)`, width `Inches(5.5)`, maintaining 9:5 aspect ratio. Table below at `top=Inches(4.3)`.
- **Table overflow:** If >15 rows, reduce font to Pt(7). If >25 rows, split to chart-only + paginated table slides.
- **Professional styling:** Navy headers (#1B365D), Calibri font, zebra striping (alternating FAFAFA/white), bold Grand Total with gray background.
- **Kaleido v1 at ~2s per chart = ~30s for 15 charts.** With pinned v0.2.1 at ~250ms per chart = ~4s total.
- **Context manager pattern for temp images:**
  ```python
  @contextmanager
  def temp_chart_images(charts, config):
      tmpdir = Path(tempfile.mkdtemp())
      try:
          paths = {name: export_chart_png(fig, tmpdir / f"{name}.png")
                   for name, fig in charts.items()}
          yield paths
      finally:
          shutil.rmtree(tmpdir)
  ```

**Files:** `ils_kickoff/charts/*.py`

#### 3.2 Improve Excel report

- [x] Create `ils_kickoff/excel_report.py`
  - **Single workbook open/close** -- NOT 15 sequential open/close cycles (current script at lines 785-788 is O(n^2) I/O, wastes 3-7 seconds)
  - Register `NamedStyle` objects once for batch styling (faster than per-cell formatting):
    - `rpt_header`: bold, navy-on-light-blue, centered, wrapped
    - `rpt_data_even`: normal font, white background
    - `rpt_data_odd`: normal font, zebra gray (#FAFAFA) background
    - `rpt_total`: bold, light gray (#F0F0F0) background
  - Frozen headers (`freeze_panes = "A2"`), autofilter, zebra striping
  - Smart column width calculation: sample-based with format-aware display width, autofilter padding
  - Number formatting using column-name heuristic from `formatting.py`:
    - `#,##0` for integer counts
    - `$#,##0.00` for currency
    - `0.0%` for percentages (store as decimal, Excel multiplies by 100)
    - `0.00` for ratios
  - Report Info cover sheet with disabled gridlines, merged title cells, report metadata
  - Table of Contents with internal hyperlinks (`#SheetName!A1`) and "Back to Contents" links
  - Chart PNG embedding: `fig.to_image()` -> `BytesIO` -> `openpyxl.drawing.image.Image()` (no temp files)
  - Output filename: `{client_id}_ILS_Kickoff_Report_{YYYYMMDD}.xlsx`

### Research Insights: Excel Formatting

**From best-practices-researcher (Excel formatting):**

- **NamedStyle is faster than per-cell styling** and produces cleaner code. Register styles once, apply by name string.
- **Number format is separate from NamedStyle.** Set visual style (fill, font, border) via `cell.style = "name"`, then override `cell.number_format` per column.
- **Percentages:** DataFrame stores `45.2`; before writing to cell, divide by 100 to get `0.452`; use format `0.0%` so Excel displays `45.2%`.
- **openpyxl vs xlsxwriter:** Stick with openpyxl. You need read/modify capability for the cover sheet + formatting after data write. xlsxwriter is write-only.
- **Install `lxml`** as optional dependency for ~2x faster XML processing in openpyxl.
- **Column width:** Account for number format display width (1234567 formatted as "$1,234,567" is 10 chars, not 7), autofilter dropdown padding, and font proportionality.

**Files:** `ils_kickoff/excel_report.py`

#### 3.3 Improve PowerPoint report

- [x] Create `ils_kickoff/pptx_report.py`
  - Title slide with navy background, client name and date
  - For each analysis: slide with chart image (upper portion) + data table (below)
    - Chart: `Inches(0.5)` left, `Inches(1.0)` top, `Inches(5.5)` wide, maintaining 9:5 ratio
    - Table: `Inches(0.3)` left, `Inches(4.3)` top, full slide width
    - Embed chart via `fig.to_image()` -> `BytesIO` -> `add_picture()` (no temp files)
  - Grand Total rows styled bold with gray background (from `formatting.py`)
  - Navy headers (#1B365D), white header text, Calibri font, zebra striping
  - Adaptive table sizing: if >15 rows reduce font, if >25 rows split to separate slides
  - Output filename: `{client_id}_ILS_Kickoff_Presentation_{YYYYMMDD}.pptx`
  - Future hook: if `config.pptx_template` is set, load template and use its slide layouts

**Files:** `ils_kickoff/pptx_report.py`

### Phase 4: Streamlit Web UI (Optional)

A simple web interface for non-technical users who cannot use the CLI.

#### 4.1 Build single-page Streamlit app

- [ ] Create `app.py` (at project root, not inside package)
  - **Sidebar:** File upload + form with config options + "Run Analysis" button
  - **Main area:** Grouped tabs (Account Status, Deposits, NSF, OD, Reg E) with expanders inside each
  - **Progress:** Use `st.status` with `st.progress` for step-by-step feedback during analysis
  - **Downloads:** `BytesIO` buffers with correct MIME types in sidebar (always visible, not buried in scroll)
  - **Caching:** `@st.cache_data` on data loading (keyed on file content hash, not UploadedFile object); `st.session_state` for analysis results
  - **Error handling:** Catch custom exceptions and display via `st.error()` -- no raw tracebacks

### Research Insights: Streamlit UX

**From best-practices-researcher (Streamlit):**

- **`st.tabs` does NOT lazy-load** -- all tab content is computed on render. With 15 Plotly charts, this causes 8+ second initial render and UI lag. Use `st.expander(expanded=False)` inside each tab so only expanded charts render, or `st.segmented_control` + `st.fragment` (Streamlit 1.37.0+) for true conditional rendering.
- **Plotly chart limit:** Streamlit has known issues rendering >8 Plotly charts simultaneously (charts may fail to display, form inputs become unresponsive). The tab + expander pattern mitigates this.
- **Gate analysis behind button click.** Without an explicit `if run_btn:` guard, analysis re-runs on every widget interaction. Store results in `st.session_state`.
- **Use `st.form` for multi-input submission.** Without forms, each text input change triggers a full script rerun.
- **Bind to localhost** by default for security: `[server] address = "127.0.0.1"` in `.streamlit/config.toml`.
- **Marimo was evaluated as an alternative** but Streamlit is a better fit: larger ecosystem, better documentation, and the batch-pipeline-with-view pattern matches Streamlit's architecture. Consider Marimo if interactive parameter exploration becomes a requirement.
- **Download button MIME types matter:** `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` for xlsx, `application/vnd.openxmlformats-officedocument.presentationml.presentation` for pptx.

**Files:** `app.py`

### Phase 5: PowerPoint Template Support (Deferred)

Marked as secondary priority. Requires a `.pptx` template from the design team.

#### 5.1 Template-based PPTX generation

- [ ] Add template inspection utility: `python -m ils_kickoff inspect-template templates/company.pptx` -- prints layout names, placeholder indices, and sizes
- [ ] Extend `ils_kickoff/pptx_report.py` to load a template and populate placeholders by index
- [ ] Document template requirements in README (expected layouts, placeholder naming conventions)
- [ ] Add `templates/` directory with a sample template

**Blocked by:** User providing the branded `.pptx` template

**Files:** `ils_kickoff/pptx_report.py` (extend), `templates/`

## Acceptance Criteria

### Functional Requirements

- [x] Non-technical user can run the analysis by: (1) placing CSV in `data/`, (2) editing `config.yaml` to set `data_file`, (3) running `make run` or double-clicking `run.bat`
- [x] Changing client data requires editing only `data_file` in `config.yaml` (and optionally `client_id`/`client_name`)
- [x] All 15 analyses produce both a data table and a Plotly chart
- [x] Excel report has formatted headers, frozen panes, autofilter, zebra striping, number formatting, cover sheet, table of contents with hyperlinks
- [x] PowerPoint has chart images alongside tables on each analysis slide
- [x] CLI supports `--help` with clear descriptions of all options
- [ ] Streamlit UI allows file upload, displays all results, and offers download buttons
- [x] Tool handles edge cases: no Business accounts (skips relevant analyses), zero totals (no crashes), missing optional columns (warns and skips)
- [x] Clear error messages when CSV is missing, has wrong columns, or is unreadable (no raw tracebacks)

### Non-Functional Requirements

- [x] Python 3.10+ compatibility
- [x] All dependencies declared in `requirements.txt` with compatible-release pins
- [x] No client data committed to git (`.gitignore` enforced)
- [x] No hardcoded file paths, client IDs, or year values in source code
- [x] Structured logging (`logging` module, not `print()`; INFO default, DEBUG via `--verbose`)
- [x] `yaml.safe_load()` only (never `yaml.load()`) for all YAML parsing
- [x] Output directory created with restrictive permissions (`mode=0o700`)
- [x] Path validation on data_file to prevent path traversal
- [x] No global `warnings.filterwarnings("ignore")` -- targeted suppression only

### Quality Gates

- [x] Tests pass for data loading, column mapping, settings, CLI, and all 15 analyses using synthetic fixture data
- [x] 80% code coverage minimum on business logic (analyses, data loading, config) -- achieved 91%
- [x] Script runs end-to-end on the provided CSV producing all outputs without errors
- [ ] Streamlit app starts and displays results for sample data
- [ ] `ruff check` and `ruff format --check` pass with no errors
- [x] Total pipeline runtime <10 seconds for 17,706-row dataset (with kaleido 0.2.1)

### Research Insights: Performance Targets

**From performance-oracle:**

| Component | 17K Rows | 100K Rows |
|-----------|----------|-----------|
| CSV Load | 100ms | 500ms |
| 15 Analyses | 50ms | 200ms |
| 15 Plotly Charts (build) | 200ms | 200ms |
| 15 PNG Exports (kaleido 0.2.1) | 3s | 3s |
| 15 PNG Exports (kaleido 1.0+) | **30s** | **30s** |
| Excel Report (single writer) | 500ms | 500ms |
| Excel Report (current 15x open) | **5s** | **5s** |
| PowerPoint | 1s | 1s |
| **Total (optimized)** | **~5s** | **~5.5s** |
| **Total (worst case, kaleido v1+)** | **~37s** | **~37s** |

Analyses use `groupby().agg()` which is O(n) -- even at 1M rows, each takes <100ms. No parallelization needed. The bottleneck is Kaleido image export.

## Dependencies and Prerequisites

| Dependency | Purpose | Required By |
|---|---|---|
| Python 3.10+ | Runtime | All phases |
| pandas~=2.2, numpy~=1.26 | Data manipulation | Phase 1 |
| pyyaml~=6.0, pydantic~=2.6 | Configuration | Phase 1 |
| typer[all]~=0.12 | CLI interface (includes rich) | Phase 1 |
| plotly~=6.0, kaleido==0.2.1 | Charts (interactive + static export) | Phase 3 |
| python-pptx~=1.0 | PowerPoint generation | Phase 3 |
| openpyxl~=3.1 | Excel formatting | Phase 3 |
| lxml (optional) | Faster openpyxl XML processing | Phase 3 |
| streamlit~=1.37 | Web UI (optional) | Phase 4 |
| Branded .pptx template | Template-based slides | Phase 5 (blocked) |

**Dev dependencies:** pytest~=8.0, pytest-cov~=5.0, ruff~=0.4

## Risk Analysis and Mitigation

| Risk | Impact | Mitigation |
|---|---|---|
| Client banking data already exposed in public repo | **CRITICAL** -- potential GLBA violation | Phase 0: Make repo private, purge git history, contact GitHub Support, assess notification obligations |
| New client CSVs have different column names | Analysis crashes | Column alias system in `column_map.py` with configurable aliases; clear error messages listing expected vs actual columns |
| DepositCount tab-embedded format varies | Wrong deposit counts | Configurable cleaner that activates only when tabs detected; log warnings |
| Kaleido v1.0+ performance regression (50x slower) | 30s+ for chart export | Pin `kaleido==0.2.1`; document Chrome requirement; skip chart images gracefully if Kaleido fails |
| Streamlit with 15+ Plotly charts causes UI freeze | Unresponsive app | Use `st.expander(expanded=False)` inside grouped tabs; `@st.cache_data` on data loading; `st.fragment` for partial reruns |
| Streamlit adds scope creep | Delays core improvements | Phase 4 is optional; Phase 1-3 deliver a complete CLI-first tool |
| Template PPTX design never provided | Phase 5 blocked | Phase 3 delivers a good programmatic PPTX; template support is additive |
| openpyxl cell-by-cell formatting slow for large datasets | Long report generation | Use `NamedStyle` for batch styling; for >10K rows consider `xlsxwriter` write-only mode |

## Open Questions

These assumptions were made during planning. Override them if needed before implementation:

1. **Source of truth:** Using script analysis logic + new Plotly charts from scratch (notebook uses matplotlib). Override if the notebook's analysis logic is preferred.
2. **DepositCount extra values:** Treating the 4 tab-separated trailing values in DepositCount as irrelevant. Override if they represent meaningful data (the values are all `0.00` in the current dataset).
3. **Batch processing:** Not in scope. Users can loop CLI invocations for multiple clients. Override if batch is a priority.
4. **Notebook fate:** Archived to `legacy/`. Override if the notebook should be maintained alongside the new tool.
5. **Client ID format:** Regex `r'^(\d+)'` from filename. Override if client IDs can be non-numeric.
6. **Kaleido version:** Pinned to 0.2.1 for performance. Override if Plotly 6.x drops v0 support.
7. **Streamlit authentication:** Not included in Phase 4 (localhost-only). Override if the app will be network-accessible.

## References

### Internal References
- Script analysis logic: `ils_kickoff.py:229-666` (15 analyses)
- Script PPTX generation: `ils_kickoff.py:669-728`
- Script Excel generation: `ils_kickoff.py:731-790`
- Script column rename: `ils_kickoff.py:54-70`
- Script DepositCount parsing: `ils_kickoff.py:74-82`
- Script `add_grand_total` (defined but never called): `ils_kickoff.py:104-115`
- Script `format_ppt_table`: `ils_kickoff.py:118-168`
- Notebook chart code: Cells 22, 25, 27, 33, 39, 45, 52, 59, 66, 72, 78 (matplotlib, NOT Plotly)
- Notebook Excel formatting: Cells 13-14 (`format_excel_table`)
- Notebook column aliases: Cell 5 (`to_canonical`)
- Hardcoded year bug: `ils_kickoff.py:624-625`
- Hardcoded client ID: `ils_kickoff.py:47, 678, 726, 739, 751`
- Excel 15x open/close anti-pattern: `ils_kickoff.py:785-788`
- Global warning suppression: `ils_kickoff.py:34`

### External References
- [python-pptx documentation](https://python-pptx.readthedocs.io/)
- [Plotly Python documentation](https://plotly.com/python/)
- [Plotly Static Image Export](https://plotly.com/python/static-image-export/)
- [Kaleido v1 Performance Issue #400](https://github.com/plotly/Kaleido/issues/400)
- [Streamlit documentation](https://docs.streamlit.io/)
- [Streamlit st.fragment](https://docs.streamlit.io/develop/api-reference/execution-flow/st.fragment)
- [Typer documentation](https://typer.tiangolo.com/)
- [Pydantic BaseModel](https://docs.pydantic.dev/latest/concepts/models/)
- [openpyxl Styles Documentation](https://openpyxl.readthedocs.io/en/stable/styles.html)
- [openpyxl Performance Guide](https://openpyxl.readthedocs.io/en/stable/performance.html)
- [Cookiecutter Data Science](https://cookiecutter-data-science.drivendata.org/)
- [pydantic-settings YAML Issues](https://github.com/pydantic/pydantic-settings/issues/366)
