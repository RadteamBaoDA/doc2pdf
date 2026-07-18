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
- `logging`: log level, log file, rotation, and retention.
- `post_processing`: conservative PDFium pixel-based whitespace trimming after
  PDF generation. The default `physical` mode tightens both MediaBox and CropBox;
  `cropbox` keeps the original physical page for compatibility.

### Reliable Excel output

Excel conversion uses an isolated `DispatchEx` instance and printer-verified
paper/orientation pairs. Its quality-first planner measures content width and
height, margins, and repeating print titles. The default `orientation: auto`
evaluates both portrait and landscape; an explicit orientation restricts the
planner to that choice. A layout that satisfies the requested fit axes at
`min_shrink_factor: 0.90` uses the largest numeric `PageSetup.Zoom` up to 100%,
with both fit-to-pages flags disabled.

If no candidate can satisfy the requested fit axes at the 90% quality floor,
the default `oversized_action: paginate` keeps numeric Zoom at 90% and lets Excel paginate
horizontally and vertically. The planner chooses the accepted layout with the
fewest estimated pages, then the smaller paper. `error`, `skip`, and `warn`
remain available for compatibility; `warn` may fit below the configured quality
floor.

Set `row_dimensions` to `null` for natural vertical pagination, `0` to try a
one-page-tall region before the quality-preserving fallback, or a positive row
count to cap the number of source rows in each logical chunk. A chunk may still
span physical pages rather than shrink below the floor. Existing multi-area
PrintAreas are retained with `print_area_policy: preserve`; use `auto` to derive
bounds from formula/value cells and visible shapes. `print_title_rows` and
`print_title_columns` accept Excel A1 ranges such as `$1:$2` and `$A:$B`; `null`
preserves the workbook's existing settings.

`oversized_action: error` is intentionally stricter: it requires the selected
region to fit both width and height on one page even when `row_dimensions` is
`null`.

The PDF is exported to a unique staging file, validated, optionally trimmed,
and only then atomically replaces the destination. Excel jobs run in spawned
workers; timeout or trim failure leaves an existing destination untouched and
never terminates unrelated Excel sessions.

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
- `suffix`: PDF filename suffixes for Word, Excel, and PowerPoint.
- `reporting`: result and error reports, plus copying failed source files.
- `pdf_handling`: handling of existing PDF files in the input.
- PDF and layout settings for Word, Excel, and PowerPoint, including pattern-based rules.

`convert-macros` does not use the PDF settings in `config.yml`.

## Testing

```powershell
python -m pytest
```

## Operational Notes

- Close open Office documents before running a batch to avoid file locks and COM dialogs.
- Do not open an output file while it is being converted.
- Logs, summaries, and error reports are written according to `config.yml`.
- If PowerShell cannot find `doc2pdf`, activate the virtual environment or replace `doc2pdf` in the examples with `python -m src.cli`.
