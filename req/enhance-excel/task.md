# Excel Quality-First Enhancement Task Register

**Specification:** [spec.md](spec.md)  
**Plan:** [implementation_plan.md](implementation_plan.md)  
**Status legend:** unchecked items are not implemented; check an item only when its evidence item is also complete.

## 1. Scope-lock protocol

- Work is limited to the requirements in `spec.md`.
- Preserve existing atomic output, timeout cleanup, and compatibility behavior unless a listed task changes it explicitly.
- Do not execute macros, refresh external connections, add a raster renderer, or broaden changes to Word/PowerPoint.
- Stop for approval if implementation requires a materially different file format, external service, printer installation strategy, or destructive migration.
- New discoveries must be added to this register with requirement mapping before implementation.

## 2. Delivery milestones

| Milestone | Scope | Exit condition |
|---|---|---|
| M0 | Baseline, schema, decision model | New config and manifests tested; legacy unchanged |
| M1 | Authored layout and immutable staging | Valid authored layouts are not mutated |
| M2 | Content resolver and safe chunking | No printable fixture content is unaccounted for or split unsafely |
| M3 | Printer-aware sheet planner | One deterministic quality decision per sheet |
| M4 | Actual pagination and PDF postflight | Final PDFs prove layout/content quality before commit |
| M5 | Calculation, fonts, color, trim | OCR and visual quality policies enforced end-to-end |
| M6 | Overflow/compliance extensions | Strategies isolated and validated |
| M7 | Windows CI and strict rollout | Automated acceptance passes and strict becomes deployable |

## 3. M0 — Schema, profiles, and evidence model

### Work checklist

- [ ] `EXQ-001` Add every field in specification section 6 to `ExcelSettings` with validation and documentation.
- [ ] `EXQ-002` Implement `strict`, `balanced`, and `legacy` profile expansion.
- [ ] `EXQ-003` Define explicit-setting precedence over profile defaults and sheet/folder rules.
- [ ] `EXQ-004` Preserve deprecated `page_shrink_threshold` warning-only migration behavior.
- [ ] `EXQ-005` Implement versioned `LayoutDecision` and stable decision IDs.
- [ ] `EXQ-006` Add structured serialization for rejected candidates and policy exceptions.
- [ ] `EXQ-007` Capture baseline artifacts for the current fixture pack.
- [ ] `EXQ-008` Update sample config comments without presenting unimplemented fields as active before their feature lands.

### Acceptance criteria

- [ ] Existing configs load into effective legacy behavior without output-setting drift.
- [ ] Invalid enum, paper, dimension, font, DPI, and policy combinations fail during config load.
- [ ] Sheet-specific settings do not leak into unrelated sheets.
- [ ] Equivalent inputs produce byte-stable decision JSON apart from documented timestamps/paths.

### Evidence checklist

- [ ] Unit tests in `tests/test_config.py` and `tests/test_excel_config_folder.py`.
- [ ] Before/after effective-config snapshots.
- [ ] Example `LayoutDecision` JSON committed under test fixtures.

## 4. M1 — Preserve authored layout and source structure

### Work checklist

- [ ] `EXQ-101` Implement authored PageSetup inspection and validity classification. `[XL-QF-001]`
- [ ] `EXQ-102` Snapshot paper, orientation, margins, Zoom/Fit flags, PrintArea, print titles, page order, manual breaks, headers/footers, color, and draft state. `[XL-QF-001]`
- [ ] `EXQ-103` Add no-optimization fast path for `preserve_authored`. `[XL-QF-001]`
- [ ] `EXQ-104` Fail closed for invalid authored PrintArea under `preserve_strict`. `[XL-QF-001, XL-QF-003]`
- [ ] `EXQ-105` Remove source-sheet row insertion from strict staging. `[XL-QF-008]`
- [ ] `EXQ-106` Implement metadata `preserve`, `append`, and `replace` without clearing unspecified sections. `[XL-QF-008]`
- [ ] `EXQ-107` Escape Excel header control characters and validate header/footer limits. `[XL-QF-008]`
- [ ] `EXQ-108` Apply permitted metadata before final PageSetup readback. `[XL-QF-008]`
- [ ] `EXQ-109` Fingerprint staged authored layout and report canonicalization. `[XL-QF-010, XL-QF-014]`

### Acceptance criteria

- [ ] A valid authored layout exports with the same PageSetup fingerprint.
- [ ] Invalid authored PrintArea never falls back to automatic discovery in strict mode.
- [ ] Source rows, formulas, merges, tables, titles, and object anchors are unchanged.
- [ ] Existing headers and footers survive the default strict policy.

### Evidence checklist

- [ ] Unit/mock tests for all three layout policies and metadata policies.
- [ ] Windows fixture comparison for authored A4 portrait and A3 landscape sheets.
- [ ] Structural before/after workbook manifest.

## 5. M2 — Printable content and safe chunking

### Work checklist

- [ ] `EXQ-201` Inventory visible worksheet and chart-sheet types through `workbook.Sheets`. `[XL-QF-003]`
- [ ] `EXQ-202` Preserve authored multi-area PrintArea ordering. `[XL-QF-003]`
- [ ] `EXQ-203` Resolve value/formula, merge, table, pivot, styled-cell, conditional-format, and text-overflow bounds. `[XL-QF-003]`
- [ ] `EXQ-204` Resolve printable shapes/charts/images/camera/OLE rectangles and honor `PrintObject`. `[XL-QF-003]`
- [ ] `EXQ-205` Isolate per-object COM failures and expose region confidence. `[XL-QF-003]`
- [ ] `EXQ-206` Build forbidden row/column boundaries from atomic content. `[XL-QF-004]`
- [ ] `EXQ-207` Treat positive `row_dimensions` as a soft maximum and shift to the nearest safe boundary. `[XL-QF-004]`
- [ ] `EXQ-208` Detect impossible single elements and title bands before candidate scoring. `[XL-QF-004, XL-QF-006]`
- [ ] `EXQ-209` Keep `null`, `0`, and positive row semantics backward-compatible outside strict mode. `[XL-QF-004]`

### Acceptance criteria

- [ ] Every printable fixture cell/object is represented by a region manifest.
- [ ] One object read failure does not hide later objects; strict mode reports the uncertainty.
- [ ] No safe chunk boundary intersects a merge, table keep-together group, or printable object.
- [ ] Shapes-only and chart sheets remain exportable and ordered.

### Evidence checklist

- [ ] Pure-data safe-boundary tests.
- [ ] Mock COM tests for per-object failures.
- [ ] Render comparisons for merge/chart/image boundary fixtures.

## 6. M3 — Printer-aware sheet-level layout

### Work checklist

- [ ] `EXQ-301` Enforce configured printer according to `printer_policy`. `[XL-QF-005]`
- [ ] `EXQ-302` Record printer, driver/version, port, forms, and fallback decisions. `[XL-QF-005, XL-QF-014]`
- [ ] `EXQ-303` Map full supported paper/custom-form catalog. `[XL-QF-005]`
- [ ] `EXQ-304` Read imageable area/hard margins per form and orientation. `[XL-QF-005]`
- [ ] `EXQ-305` Reject invalid usable geometry without positive-value clamping. `[XL-QF-005]`
- [ ] `EXQ-306` Add page dimension/area, font, image-DPI, and preferred-paper constraints. `[XL-QF-006]`
- [ ] `EXQ-307` Implement horizontal-split-first scoring and deterministic tie-break. `[XL-QF-006]`
- [ ] `EXQ-308` Implement `page_size_scope: sheet` using the worst-case region/chunk. `[XL-QF-002]`
- [ ] `EXQ-309` Apply one paper/orientation to every chunk of the sheet. `[XL-QF-002]`
- [ ] `EXQ-310` Set/readback PageOrder and reject impossible repeated titles. `[XL-QF-006]`

### Acceptance criteria

- [ ] Different sheets may select different page sizes, but one sheet never mixes sizes in sheet scope.
- [ ] Strict mode never silently uses an unconfigured system-default printer.
- [ ] Candidate selection meets 90% scale, page, font, and DPI budgets.
- [ ] Ties always resolve identically and favor fewer horizontal splits.

### Evidence checklist

- [ ] Boundary unit tests for every eligible paper/orientation.
- [ ] Tests for A2 advertised/not advertised and custom forms.
- [ ] Windows printer imageable-area/readback report.
- [ ] Mixed-sheet PDF MediaBox report.

## 7. M4 — Actual pagination, export, and postflight

### Work checklist

- [ ] `EXQ-401` Force Excel print-layout recomputation after final PageSetup. `[XL-QF-007]`
- [ ] `EXQ-402` Read actual H/V page breaks and distinguish manual versus automatic breaks where available. `[XL-QF-007]`
- [ ] `EXQ-403` Implement `manual_page_break_policy: preserve|reset`. `[XL-QF-007]`
- [ ] `EXQ-404` Retry/re-score bounded candidate mismatches without crossing the quality floor. `[XL-QF-007]`
- [ ] `EXQ-405` Export verified staged sheets individually. `[XL-QF-010]`
- [ ] `EXQ-406` Merge sheet PDFs without rasterization and preserve workbook order. `[XL-QF-010]`
- [ ] `EXQ-407` Reverify PageSetup after any retained cross-workbook copy. `[XL-QF-010]`
- [ ] `EXQ-408` Track expected and actual page ranges per sheet/region. `[XL-QF-010]`
- [ ] `EXQ-409` Build source sentinel/object manifest. `[XL-QF-011]`
- [ ] `EXQ-410` Validate PDF page boxes, rotation, blanks, searchable text, sentinels, fonts, clipping, edge ink, rasterization, image DPI, and page-size budgets. `[XL-QF-011]`
- [ ] `EXQ-411` Connect strict/warn/disabled postflight policies to atomic commit. `[XL-QF-011]`

### Acceptance criteria

- [ ] Recorded actual page grid agrees with the exported PDF.
- [ ] Pagination in one sheet does not disable exact validation for another sheet.
- [ ] Sheet PDF merge preserves text searchability, vectors, Unicode, boxes, and order.
- [ ] Every strict postflight failure preserves the previous destination byte-for-byte.

### Evidence checklist

- [ ] Mock pagination mismatch/retry tests.
- [ ] Per-sheet PDF postflight JSON.
- [ ] Atomic failure tests for export, merge, postflight, and skip/error actions.
- [ ] Windows PDF page/sentinel/font validation report.

## 8. M5 — Calculation, fonts, print quality, and trim

### Work checklist

- [ ] `EXQ-501` Implement `saved_cache`, `calculate`, and `full_rebuild` policies with completion waits. `[XL-QF-009]`
- [ ] `EXQ-502` Inventory external links/connections without refreshing under the default policy. `[XL-QF-009]`
- [ ] `EXQ-503` Record calculation mode/state, formula errors, macro/UDF limits, and freshness evidence. `[XL-QF-009, XL-QF-014]`
- [ ] `EXQ-504` Add font availability/substitution preflight. `[XL-QF-009]`
- [ ] `EXQ-505` Enforce standard Excel export quality and `Draft=False` in strict mode. `[XL-QF-009]`
- [ ] `EXQ-506` Implement preserve/force-color/black-and-white policies. `[XL-QF-009]`
- [ ] `EXQ-507` Prevent generic low-image-quality settings from downgrading strict Excel jobs. `[XL-QF-009]`
- [ ] `EXQ-508` Add Excel-specific disabled/cropbox/physical trim policy. `[XL-QF-012]`
- [ ] `EXQ-509` Add higher-resolution protected-edge and before/after render verification. `[XL-QF-012]`
- [ ] `EXQ-510` Run final postflight after trim. `[XL-QF-011, XL-QF-012]`

### Acceptance criteria

- [ ] Repeated runs under one calculation policy are deterministic.
- [ ] Missing/substituted fonts cannot silently violate strict readability limits.
- [ ] Draft or minimum-quality export cannot occur in strict mode.
- [ ] Light gridlines, thin borders, rotations, annotations, existing CropBoxes, and edge text survive trim.

### Evidence checklist

- [ ] Calculation-policy tests with stale formulas and links.
- [ ] Missing-font/substitution fixture results.
- [ ] PDF font/image-DPI report.
- [ ] Before/after trim render-diff report.

## 9. M6 — Horizontal strategies and PDF/A

### Work checklist

- [ ] `EXQ-601` Implement and test `paginate`, `error`, and bounded `one_logical_page`. `[XL-QF-013]`
- [ ] `EXQ-602` Define experimental vector-stitch page order and repeated-title de-duplication. `[XL-QF-013]`
- [ ] `EXQ-603` Implement vector-only stitch without rasterizing text or objects. `[XL-QF-013]`
- [ ] `EXQ-604` Enforce PDF user-space, OCR dimension, and area limits on stitched pages. `[XL-QF-013]`
- [ ] `EXQ-605` Add an explicit fallback/error policy for stitch failure. `[XL-QF-013]`
- [ ] `EXQ-606` Add optional PDF/A postprocessing and standards validation. `[XL-QF-015]`
- [ ] `EXQ-607` Emit compliance claims only after validator success. `[XL-QF-015]`

### Acceptance criteria

- [ ] Every overflow strategy has deterministic output and failure behavior.
- [ ] Vector stitch retains searchable Unicode text, lines, images, and object positions.
- [ ] No page exceeds configured page/OCR budgets.
- [ ] A nonvalidated Excel PDF is never labeled PDF/A.

### Evidence checklist

- [ ] Strategy matrix tests.
- [ ] Vector/text comparison for stitched output.
- [ ] PDF/A validator reports for pass and fail fixtures.

## 10. M7 — Integration, operations, and rollout

### Work checklist

- [ ] `EXQ-701` Expand the fixture pack for all specification edge cases.
- [ ] `EXQ-702` Provision a pinned Windows + desktop Excel + printer-driver CI runner.
- [ ] `EXQ-703` Automate page count, MediaBox, sentinel, font, DPI, object, and render-diff checks.
- [ ] `EXQ-704` Run strict and legacy shadow comparisons.
- [ ] `EXQ-705` Add performance timings and enforce bounded probe/postflight budgets.
- [ ] `EXQ-706` Document printer provisioning, diagnostics, staging retention, and rollback.
- [ ] `EXQ-707` Update README and `config.yml` only after implemented behavior is verified.
- [ ] `EXQ-708` Promote strict sample defaults after two clean full-matrix runs.

### Acceptance criteria

- [ ] Wide, tall, wide+tall, authored, mixed-size, merge/object, title, manual-break, font, formula, printer, and trim fixtures all pass.
- [ ] No unexplained content delta exists between source manifest and strict PDF.
- [ ] Timeout and every failure stage preserve atomic destination semantics.
- [ ] Legacy rollback remains documented and tested.

### Evidence checklist

- [ ] Two clean CI run links/artifacts.
- [ ] Strict-versus-legacy comparison report.
- [ ] Performance report by conversion stage.
- [ ] Operations and rollback runbook.

## 11. Final acceptance checklist

- [ ] All `XL-QF-001` through `XL-QF-015` requirements have implementation evidence.
- [ ] Every completed work item has its acceptance and evidence items completed.
- [ ] Unit, mock COM, and Windows integration suites pass.
- [ ] `git diff --check` passes.
- [ ] Sample config, README, specification, plan, and task register agree on defaults and semantics.
- [ ] Strict mode commits only postflight-verified output.
- [ ] No visible sheet, region, object, or authored layout can be silently lost or replaced.

