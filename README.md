# doc2pdf

[Tiếng Việt](docs/README.vi.md)

A Windows CLI that uses Microsoft Office to convert Word, Excel, and PowerPoint documents to PDF. It supports single files, recursive directory processing, path-pattern configuration, error reports, and PDF post-processing.

In addition to PDF conversion, the `convert-macros` command creates macro-free copies:

- `.docm` → `.docx`
- `.pptm` → `.pptx`
- `.xlsm` → `.xlsx`

> Converting to a macro-free format removes the VBA project. The source file is not modified.

## Requirements

- Windows
- Python 3.12+
- Microsoft Word, Excel, and/or PowerPoint installed for the corresponding file types

## Installation

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install .
```

Development installation:

```powershell
python -m pip install -e ".[dev]"
```

## Commands and Run Modes

### Help and version

```powershell
doc2pdf --help
doc2pdf --version
doc2pdf convert --help
doc2pdf convert-macros --help
```

Run directly from source without installing the command:

```powershell
python -m src.cli --help
```

### Mode 1: Office to PDF

Convert one file. The default output directory is `output`:

```powershell
doc2pdf convert "input\report.docx"
```

Specify the output PDF file:

```powershell
doc2pdf convert "input\report.docx" --output "output\report.pdf"
```

Recursively convert a directory while preserving its subdirectory structure:

```powershell
doc2pdf convert "input" --output "output"
```

Use a different configuration file:

```powershell
doc2pdf convert "input" --output "output" --config "config.yml"
```

Enable verbose logs, enable or disable whitespace trimming, or override the trim margin:

```powershell
doc2pdf convert "input" --output "output" --verbose
doc2pdf convert "input" --output "output" --trim
doc2pdf convert "input" --output "output" --no-trim
doc2pdf convert "input" --output "output" --trim --trim-margin 10
```

Recognized inputs are `.doc`, `.docx`, `.xls`, `.xlsx`, `.xlsm`, `.xlsb`, `.ppt`, `.pptx`, and `.pdf`. The `pdf_handling` section in `config.yml` controls how existing PDF files in the input are handled.

### Mode 2: Remove macros without creating PDFs

Convert one file. The default output directory is `output`:

```powershell
doc2pdf convert-macros "input\report.docm"
doc2pdf convert-macros "input\slides.pptm"
doc2pdf convert-macros "input\workbook.xlsm"
```

Specify the exact output file:

```powershell
doc2pdf convert-macros "input\report.docm" --output "clean\report.docx"
```

Convert all three supported types in a directory while preserving its subdirectory structure:

```powershell
doc2pdf convert-macros "input" --output "clean"
```

This command processes only `.docm`, `.pptm`, and `.xlsm` files. Other files in the directory are ignored. Macros are disabled when Office opens the files through automation and are not saved in the output files.

## Configuration

`config.yml` is the default configuration file. Its main sections include:

- `timeout`: document conversion and Excel trimming timeouts.
- `parallel.excel_workers`: `auto` (default) or a fixed `1`-`8` isolated Excel
  conversions; `parallel.excel_worker_cap` bounds auto mode (default `4`).
- `logging`: log level, log file, rotation, and retention.
- `post_processing`: conservative PDFium pixel-based whitespace trimming after
  PDF generation. The default `physical` mode tightens both MediaBox and CropBox;
  `cropbox` keeps the original physical page for compatibility.
- `suffix`: PDF filename suffixes for Word, Excel, and PowerPoint.
- `reporting`: result and error reports, plus copying failed source files.
- `pdf_handling`: handling of existing PDF files in the input.
- PDF and layout settings for Word, Excel, and PowerPoint, including pattern-based rules.

### Parallel Excel batches

Folder conversions run separate Excel workbooks in isolated spawned processes.
Auto mode chooses the minimum allowed by file count, half the logical CPUs,
available memory (2 GB reserve plus 1.5 GB per worker), and the configured cap.
It falls back to at most two workers if memory cannot be measured. Set
`parallel.excel_workers: 1` for explicit serial conversion. Larger estimated
workbooks enter the queue first, and only transient COM/printer/resource errors
receive one isolated serial retry. Parallelism is applied between workbooks,
not between worksheets in one workbook. Word, PowerPoint, and PDF handling
remain serial.

### Reliable Excel output

Excel supports `strict`, `balanced`, and `legacy` quality profiles.
Configurations without `quality_profile` now use strict quality-first behavior;
set `quality_profile: legacy` explicitly for compatibility. Strict mode uses an
isolated `DispatchEx` process, preserves persisted
authored print layouts, plans one printer-verified paper/orientation per sheet,
and refuses unverifiable output. Balanced mode uses the same pipeline with
warning postflight and CropBox trimming defaults.

Strict conversion caches printer form/geometry probes within one workbook,
plans a layout once per sheet, and reuses that verified decision across chunks.
Each staged chunk still receives required PageSetup readback and actual page-break
verification. Manifest schema v2 records phase timings and the resolved batch
concurrency decision; batch summaries include per-file duration and scheduling
evidence.

Authored XLSX/XLSM layout is detected from stored workbook XML. XLS/XLSB uses
conservative positive COM signals such as PrintArea, print titles, manual page
breaks, and custom headers. Missing layouts enter smart planning; only
`layout_policy: force_optimize` replaces a valid authored layout.

The quality planner inventories cell content and printable objects, moves row
chunk boundaries around merged ranges, tables, title bands, charts, and shapes,
then verifies Excel's actual page breaks. `orientation: auto` evaluates both
orientations. Accepted smart layouts use numeric `PageSetup.Zoom` at or above
`min_shrink_factor: 0.90`, with fit-to-pages flags disabled.

If no candidate can satisfy the requested fit axes at the 90% quality floor,
the default `oversized_action: paginate` keeps numeric Zoom at 90% and lets Excel
paginate horizontally and vertically. Strict ranking minimizes horizontal
splits, total pages, preferred-paper rank, whitespace, and paper area in that
order. `error`, `skip`, and `warn` remain available; low-quality `warn` is a
legacy/balanced compatibility behavior.

Set `row_dimensions` to `null` for natural vertical pagination, `0` to try a
one-page-tall region before the quality-preserving fallback, or a positive row
count to cap the number of source rows in each logical chunk. A chunk may still
span physical pages rather than shrink below the floor. Existing multi-area
PrintAreas are retained with `print_area_policy: preserve`; strict mode uses
`preserve_strict` and fails closed on an invalid area. Use `auto` to derive
bounds from cells and printable objects. `print_title_rows` and
`print_title_columns` accept Excel A1 ranges such as `$1:$2` and `$A:$B`; `null`
preserves the workbook's existing settings.

`oversized_action: error` is intentionally stricter: it requires the selected
region to fit both width and height on one page even when `row_dimensions` is
`null`.

Strict/balanced sheets are exported separately and merged without rasterization.
Postflight checks page ranges and boxes, blank pages, sentinel/searchable text,
font sizes, image DPI, rasterization, clipping indicators, and page budgets.
Versioned decision manifests are written under the configured reports directory.
All export, merge, trim, and verification work uses unique staging paths; the
parent worker replaces the destination only after final validation. Timeout or
failure leaves an existing destination untouched.

Strict Excel trimming is disabled by default. `cropbox` and explicit `physical`
trim modes re-render at least 150 DPI and verify that text and source ink remain
inside the resulting page boxes. Generic low-image-quality settings cannot
downgrade strict Excel output.

`one_logical_page`, `vector_stitch`, and PDF/A conversion are reserved extension
points. They fail explicitly in M0-M5 because no stitch or standards-validation
provider is bundled; Excel output is never labeled PDF/A without validator
success. Set `compliance: standard` when opting into strict mode.

### Edge-case fixture pack

The reproducible stress fixtures live under [tests/fixtures](tests/fixtures/README.md).
Build Excel sources with the Spreadsheet skill's artifact-tool builder, then seed
real `PrintArea` and printer settings with `seed_excel_print_settings.ps1` on a
Windows host with Excel. The matrix includes multiple non-A1 areas, shapes,
merged/wrapped/Unicode cells, 2,048-column width stress, maximum-width
oversized failures, quality pagination, print titles, auto orientation, rejected
printer forms, and 90% scale boundaries. Word and PowerPoint edge fixtures have
a separate Office COM seeder because this repository's acceptance target is
desktop Office.

#### Full fixture generation commands

Run from the repository root in PowerShell:

```powershell
# Install runtime and development dependencies.
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e ".[dev]"

# Create the Excel source workbooks (requires the Codex/Spreadsheet
# artifact-tool runtime).
$repo = (Get-Location).Path
$fixtureDir = Join-Path $repo "tests\fixtures"
$artifactSkill = Join-Path $env:USERPROFILE ".codex\plugins\cache\openai-primary-runtime\presentations\26.715.12143\skills\presentations"
$artifactWorkspace = Join-Path $env:TEMP "doc2pdf-artifact-workspace"
New-Item -ItemType Directory -Force -Path $artifactWorkspace | Out-Null
node (Join-Path $artifactSkill "container_tools\setup_artifact_tool_workspace.mjs") `
  --workspace $artifactWorkspace
# Run the builder from the prepared workspace so Node resolves its local
# @oai/artifact-tool package. Keep the repository as the working directory so
# the builder writes to tests\fixtures\generated.
$artifactBuilder = Join-Path $artifactWorkspace "build_excel_fixtures.mjs"
Copy-Item (Join-Path $fixtureDir "build_excel_fixtures.mjs") $artifactBuilder -Force
node $artifactBuilder

# Apply real PrintArea, PaperSize, Orientation, hidden rows/columns, and
# one-page-wide settings through Excel COM.
$excelCases = @(
  "excel_printareas_shapes", "excel_chunks", "excel_max_width",
  "excel_oversized_error", "excel_oversized_skip", "excel_empty_shapes",
  "excel_empty_workbook", "excel_orientation_margin", "excel_quality_pagination",
  "excel_paper_forms"
)
foreach ($case in $excelCases) {
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File `
    (Join-Path $fixtureDir "seed_excel_print_settings.ps1") `
    -WorkbookPath (Join-Path $fixtureDir "generated\$case.xlsx") `
    -CaseName $case
}

# Create Word and PowerPoint fixtures (requires desktop Word and PowerPoint).
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File `
  (Join-Path $fixtureDir "build_word_ppt_fixtures.ps1") `
  -OutputDir (Join-Path $fixtureDir "generated")

# Run tests and convert the generated fixture directory.
python -m pytest tests/test_config.py tests/test_pdf_processor.py -q
python -m pytest -q
doc2pdf convert (Join-Path $fixtureDir "generated") `
  --output (Join-Path $repo "tests\fixtures\output") `
  --config (Join-Path $repo "config.yml") --trim --verbose
```

If step 2 reports that `@oai/artifact-tool` or its `package.json` is missing,
the machine does not have the Codex fixture-authoring runtime. The application
itself does not require OpenAI at runtime; only this source-workbook builder
does. Steps involving Excel/Word/PowerPoint COM require desktop Office.

PDF trimming defaults are `render_dpi: 72`, `max_render_pixels: 20000000`,
`background_tolerance: 8`, and `include_annotations: true`. Signed PDFs are
refused unless signature invalidation is explicitly allowed. Encrypted PDFs
require credentials.

`convert-macros` does not use the PDF settings in `config.yml`.

## Testing

```powershell
python -m pytest
```

Run the real two-workbook Excel acceptance test and record serial/parallel timing
properties in JUnit output on a Windows host with desktop Excel:

```powershell
$env:DOC2PDF_RUN_EXCEL_INTEGRATION = "1"
python -m pytest tests/test_excel_parallel_windows.py `
  --junitxml reports/excel_parallel_integration.xml
```

## Operational Notes

- Close open Office documents before running a batch to avoid file locks and COM dialogs.
- Do not open an output file while it is being converted.
- Logs, summaries, and error reports are written according to `config.yml`.
- If PowerShell cannot find `doc2pdf`, activate the virtual environment or replace `doc2pdf` in the examples with `python -m src.cli`.
