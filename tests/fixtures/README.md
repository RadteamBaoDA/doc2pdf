# Edge-case fixture pack

This pack is the source-of-truth matrix for the Excel/PDF conversion tests. The
generated workbooks are intentionally separate from the unit tests because some
cases require desktop Excel and a real printer driver.

Run `build_excel_fixtures.mjs` on a machine where the Spreadsheet skill's
`@oai/artifact-tool` workspace is available. It writes `.xlsx` files to
`tests/fixtures/generated/`. Then run `seed_excel_print_settings.ps1` on a
Windows machine with Excel installed. That second step applies real `PrintArea`,
paper-form, hidden-row/column, and chart settings through Excel COM so the
fixtures exercise the same PageSetup surface as production.

From the repository root, the complete PowerShell sequence is:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e ".[dev]"

$repo = (Get-Location).Path
$fixtureDir = Join-Path $repo "tests\fixtures"
$artifactSkill = Join-Path $env:USERPROFILE ".codex\plugins\cache\openai-primary-runtime\presentations\26.715.12143\skills\presentations"
$artifactWorkspace = Join-Path $env:TEMP "doc2pdf-artifact-workspace"
New-Item -ItemType Directory -Force -Path $artifactWorkspace | Out-Null
node (Join-Path $artifactSkill "container_tools\setup_artifact_tool_workspace.mjs") `
  --workspace $artifactWorkspace
$artifactBuilder = Join-Path $artifactWorkspace "build_excel_fixtures.mjs"
Copy-Item (Join-Path $fixtureDir "build_excel_fixtures.mjs") $artifactBuilder -Force
node $artifactBuilder

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

& powershell.exe -NoProfile -ExecutionPolicy Bypass -File `
  (Join-Path $fixtureDir "build_word_ppt_fixtures.ps1") `
  -OutputDir (Join-Path $fixtureDir "generated")

python -m pytest -q
doc2pdf convert (Join-Path $fixtureDir "generated") `
  --output (Join-Path $repo "tests\fixtures\output") `
  --config (Join-Path $repo "config.yml") --trim --verbose
```

The artifact-tool setup/build step is only for authoring source workbooks and
requires the Codex Spreadsheet runtime. The converter itself does not require
OpenAI. Excel, Word, and PowerPoint seeding/conversion require desktop Office.

## Fixture matrix

| File | Main cases | Expected use |
|---|---|---|
| `excel_printareas_shapes.xlsx` | Two non-A1 areas (`B3:F6`, `H20:J21`), merged cells, Unicode, wrapped text, hidden row/column, chart shape, left/right sentinels | `print_area_policy: preserve`; verify both areas and shapes survive; `row_dimensions: 0` must paginate rather than shrink below the configured floor when one-page output is not feasible |
| `excel_chunks.xlsx` | 120 populated rows, 40 columns, repeated headers, formulas, merged title | `row_dimensions: 10`; expect 12 logical chunks with at most 10 source rows each; a chunk may span physical pages; configure print-title overrides and verify they repeat |
| `excel_max_width.xlsx` | 2,048 columns with narrow and wide columns, long Unicode headers, formulas and left/right sentinels | Run quality pagination at the 90% floor; verify both sentinels remain somewhere in the complete PDF and record the horizontal page count |
| `excel_quality_pagination.xlsx` | `WideOnly`, `TallOnly`, and `WideTall` sheets with sentinels, merged cells, a chart, and repeating titles | Use `orientation: auto`, `min_shrink_factor: 0.90`, `oversized_action: paginate`, plus print-title overrides; verify multi-page output, all sentinels, merged/chart content, repeated titles, and effective PDF font scale at or above 90% |
| `excel_oversized_error.xlsx` | 512 deliberately wide columns plus a 16,384-column maximum-width sheet | `oversized_action: error`; a best 2D scale below the configured floor must fail conversion and preserve the destination |
| `excel_oversized_skip.xlsx` | One normal sheet plus one oversized sheet | `oversized_action: skip`; normal sheet is exported and oversized sheet is reported as partial |
| `excel_empty_shapes.xlsx` | Empty visible sheet, hidden sheet, and shapes-only sheet | Shapes-only sheet must be exportable with `print_area_policy: auto` |
| `excel_empty_workbook.xlsx` | Only an empty visible sheet and a hidden sheet | Conversion must fail explicitly with no output |
| `excel_orientation_margin.xlsx` | Portrait and landscape sentinel sheets with merged cells and wrapped labels | Verify usable width and height after all margins; `orientation: auto` must choose the better 2D result, while forced orientation remains fixed; quality layouts use numeric Zoom with both fit flags disabled |

## Paper-form coverage

Paper selection is not based on fixed content-width bands. The converter probes
the forms and dimensions advertised by the active printer, falling back to the
internal catalog only when driver enumeration is unavailable. Every candidate
paper/orientation pair must survive COM set/readback before it can be selected.
Because margins, content height, print titles, and driver dimensions affect the
result, a band that is correct for one printer is not a portable expectation.

Use strict COM fakes for deterministic boundaries and live Excel/printer runs
for integration evidence:

| Scenario | Expected result |
|---|---|
| `orientation: auto` | Evaluate both portrait and landscape; choose the highest eligible numeric Zoom, then smaller paper/whitespace |
| Forced portrait or landscape | Probe only the configured orientation; never silently switch orientation |
| Effective 2D scale exactly 90% | Eligible for one-page quality fit at numeric `Zoom: 90`; both fit-to-pages flags are disabled |
| Effective 2D scale immediately below 90% with `paginate` | Keep numeric `Zoom: 90`, disable both fit flags, and select the candidate with the fewest estimated pages, then smaller paper area |
| Rejected form or failed paper/orientation readback | Exclude the candidate; fail explicitly if none remain |
| A2 and other optional forms | Consider only when advertised by the active printer and COM readback succeeds |
| `print_title_rows` / `print_title_columns` are `null` | Preserve authored workbook titles; configured A1 ranges override them and repeat on paginated pages |

For every run, record content and usable dimensions, width/height/effective
scale, limiting axis, `PageSetup.PaperSize`, `Orientation`, numeric Zoom,
estimated and actual page counts, and the final PDF MediaBox.

## PDF trim cases

The generated source workbooks are converted to PDFs first. Keep the original
PDF beside the trimmed PDF for raster comparison. Include fixtures for blank,
full-bleed, rotated, existing CropBox, light gridline, thick vector, annotation,
scanned-border, non-zero-origin, and large-page cases in the PDF fixture folder.
