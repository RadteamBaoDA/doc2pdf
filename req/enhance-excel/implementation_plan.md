# Excel Quality-First Enhancement Implementation Plan

**Specification:** [spec.md](spec.md)  
**Execution checklist:** [task.md](task.md)  
**Delivery strategy:** incremental, feature-gated, atomic at every phase

## 1. Implementation principles

- Preserve and extend the existing 2D planner instead of replacing it wholesale.
- Keep COM interaction behind narrow adapters; keep scoring and policy logic pure and unit-testable.
- Separate discovery, planning, application, export, and verification.
- Make strict behavior opt-in until the Windows/Excel integration matrix passes.
- Do not weaken existing atomic destination, `error`, `skip`, or timeout behavior.
- Every phase must be independently testable and reversible through configuration.

## 2. Target architecture

```text
Workbook inventory
  -> AuthoredLayoutInspector
  -> PrintableContentResolver
  -> SafeChunkPlanner
  -> PrinterCapabilityProvider
  -> SheetLayoutPlanner
  -> ExcelPaginationProbe
  -> ImmutableSheetStager
  -> VerifiedSheetExporter
  -> PdfQualityPostflight
  -> VectorPdfMerger
  -> SafeTrim/Compliance (optional)
  -> Atomic commit
```

Suggested pure-data models:

- `WorkbookSheetInfo`
- `AuthoredLayoutSnapshot`
- `PrintableObject`
- `ResolvedRegion`
- `SafeChunk`
- `PrinterFormCapability`
- `LayoutConstraints`
- `LayoutCandidate`
- `LayoutDecision`
- `ExportedSheetArtifact`
- `PdfPostflightResult`

## 3. Phase 0 — Baseline capture and schema

**Requirements:** XL-QF-001, 002, 006, 012, 013, 014

### Work

1. Capture current output for every Excel fixture: page count, page boxes, text, and rendered images.
2. Extend `ExcelSettings` and the YAML loader with the new profile and policy fields.
3. Add strict validation for enum values, dimensions, paper names, font/DPI limits, and incompatible settings.
4. Define profile expansion and explicit-field precedence.
5. Add versioned `LayoutDecision` serialization and decision IDs.
6. Preserve existing configurations as `legacy` during the compatibility release.

### Primary files

- `src/config.py`
- `config.yml`
- `tests/test_config.py`
- `tests/test_excel_config_folder.py`
- new `src/core/excel_layout_models.py`

### Exit gate

- All old config tests pass.
- New settings round-trip deterministically.
- Loading an old config produces the same effective legacy settings plus migration warnings only.

## 4. Phase 1 — Authored-layout fast path and immutable staging

**Requirements:** XL-QF-001, 008, 010

### Work

1. Implement `AuthoredLayoutInspector` and snapshot all relevant PageSetup values.
2. Classify each sheet as valid authored, missing layout, or invalid authored.
3. Route valid authored sheets through a no-optimization fast path under `preserve_authored`.
4. Fail closed for invalid authored PrintArea in strict mode.
5. Move any permitted metadata changes to disposable staging or PDF overlay.
6. Disable row insertion features in strict profile.
7. Apply metadata before final readback and preserve unspecified headers/footers.
8. Compare authored fingerprints before export and record any canonicalization.

### Primary files

- `src/core/excel_converter.py`
- new `src/core/excel_layout.py`
- `tests/test_excel_converter.py`

### Exit gate

- Authored PageSetup fixtures export without planner mutation.
- Source workbook and in-memory source-sheet structure remain unchanged.
- Invalid authored layout never falls through to auto discovery in strict mode.

## 5. Phase 2 — Complete content discovery and safe chunks

**Requirements:** XL-QF-003, 004

### Work

1. Replace worksheet-only enumeration with typed `workbook.Sheets` inventory.
2. Implement strict multi-area PrintArea parsing and ordering.
3. Add content providers for cells/formulas, merged ranges, tables, pivots, formatting, text overflow, and printable objects.
4. Read each shape/object defensively so one COM failure does not terminate discovery.
5. Add region confidence and strict uncertainty errors.
6. Build forbidden chunk boundaries from merges, tables, title bands, and object rectangles.
7. Convert positive `row_dimensions` to a soft maximum and choose the closest safe boundary.
8. Detect atomic elements that cannot fit any eligible page at the quality floor.

### Primary files

- `src/core/excel_converter.py`
- new `src/core/excel_content.py`
- new `src/core/excel_chunking.py`
- `tests/test_excel_converter.py`
- `tests/fixtures/*`

### Exit gate

- Region manifests account for every printable fixture object.
- No generated chunk boundary intersects an atomic object.
- Chart sheets, shapes-only sheets, and multi-area PrintAreas export in deterministic order.

## 6. Phase 3 — Printer geometry and sheet-level planner

**Requirements:** XL-QF-002, 005, 006

### Work

1. Implement `PrinterCapabilityProvider` with configured-printer enforcement.
2. Collect full form mappings, custom form dimensions, imageable areas, hard margins, and driver metadata.
3. Probe and read back paper/orientation pairs; reject invalid geometry rather than clamping.
4. Extend pure candidate data with page caps, font/DPI budgets, page grid, preferred-paper rank, and rejection reasons.
5. Plan one decision for the worst-case chunk of each sheet when `page_size_scope: sheet`.
6. Update pagination ranking to minimize horizontal splits before total pages.
7. Validate print-title feasibility and set deterministic `PageOrder`.

### Primary files

- `src/core/excel_converter.py`
- new `src/core/excel_printer.py`
- new `src/core/excel_layout_planner.py`
- `tests/test_excel_converter.py`

### Exit gate

- Pure planner tests cover every paper/orientation boundary and deterministic tie-break.
- All chunks of a sheet receive the same paper/orientation in sheet scope.
- A4/A3 are preferred over oversized pages when OCR budgets favor pagination.

## 7. Phase 4 — Actual pagination and verified export

**Requirements:** XL-QF-007, 010, 011

### Work

1. Implement a COM pagination probe after final PageSetup application.
2. Preserve or reset manual page breaks according to policy.
3. Re-score candidates when actual and predicted page grids differ.
4. Replace the second cross-workbook copy path with per-sheet PDF export plus vector merge where feasible.
5. If any cross-workbook copy remains, perform full PageSetup and title readback after that copy.
6. Track expected page ranges and validation independently per sheet/region.
7. Build source manifests with boundary values/sentinels and printable-object counts.
8. Implement `PdfQualityPostflight` and execute it before atomic commit.

### Primary files

- `src/core/excel_converter.py`
- new `src/core/excel_pagination.py`
- new `src/core/pdf_quality.py`
- `src/core/job_runner.py`
- `tests/test_excel_converter.py`
- `tests/test_job_runner.py`

### Exit gate

- Exported page grids match COM pagination evidence.
- Per-sheet postflight verifies page boxes, text/sentinels, fonts, clipping, rasterization, and image DPI.
- Any strict postflight failure leaves the destination unchanged.

## 8. Phase 5 — Calculation, font, color, and trim quality

**Requirements:** XL-QF-009, 012, 014

### Work

1. Add calculation policy orchestration and wait for calculation completion.
2. Inventory external links/connections and record refresh restrictions.
3. Add formula-error and stale/freshness evidence to the manifest.
4. Implement font availability/substitution preflight.
5. Enforce standard export quality, `Draft=False`, and color policy in strict mode.
6. Separate Excel trim policy from generic PDF defaults.
7. Raise strict trim render quality and add before/after visual and protected-edge comparison.
8. Feed trim results into final postflight rather than treating trim as an unverified last step.

### Primary files

- `src/core/excel_converter.py`
- `src/core/pdf_processor.py`
- `src/core/job_runner.py`
- `src/config.py`
- `tests/test_pdf_processor.py`

### Exit gate

- Calculation/font policy is deterministic and visible in output evidence.
- Light gridlines, thin borders, annotations, and edge content survive enabled trim.
- Low global image quality cannot downgrade strict Excel export.

## 9. Phase 6 — Horizontal overflow and compliance extensions

**Requirements:** XL-QF-013, 015

### Work

1. Stabilize `paginate`, `error`, and bounded `one_logical_page` behaviors.
2. Prototype vector-only tile stitching behind an experimental flag.
3. Define page-order and repeated-title de-duplication for stitch mode.
4. Reject stitch output that exceeds PDF or OCR page budgets.
5. Add optional PDF/A postprocessor and standards validator; do not infer compliance from Excel export.

### Exit gate

- Every horizontal strategy has isolated tests and deterministic fallback semantics.
- Vector stitch preserves searchable text, vectors, images, Unicode, and object positions.
- Compliance labels are emitted only after validator success.

## 10. Phase 7 — Windows/Excel CI and rollout

**Requirements:** all

### Work

1. Expand fixtures for authored layouts, chart sheets, merge/object boundaries, manual breaks, missing fonts, stale formulas, light gridlines, printer rejection, and mixed page sizes.
2. Provision a Windows runner with desktop Excel and a pinned PDF printer/driver.
3. Automate page count, MediaBox, sentinel, font, image-DPI, and render-diff assertions.
4. Run strict and legacy modes against the same corpus and publish decision manifests.
5. Execute shadow conversions on representative production workbooks.
6. Promote strict defaults only after acceptance and performance budgets pass.

### Exit gate

- The complete automated matrix passes twice on a clean Windows runner.
- No unexplained strict-versus-legacy content delta remains.
- Operations documentation covers printer provisioning, diagnostics, rollback, and evidence retention.

## 11. Test strategy

### Unit tests

- Profile/default expansion and config validation.
- Authored-layout classification and fingerprint comparison.
- Region and safe-boundary pure-data logic.
- Candidate feasibility, OCR constraints, scoring, and deterministic tie-break.
- Equality at 90%, just below 90%, invalid title capacity, and indivisible objects.
- All `oversized_action`, layout, page-break, metadata, calculation, and overflow policies.
- PDF page-box, font, sentinel, clipping, and trim validators.

### COM contract tests with mocks

- Required property set/readback failures are fatal in strict mode.
- Printer rejection and configured fallback behavior.
- Page-break recomputation and mismatch retry.
- Chart-sheet enumeration and per-object COM failures.
- Post-copy verification when the compatibility export path is used.

### Windows integration tests

- Wide, tall, and wide+tall sheets.
- Authored A4/A3/mixed-orientation workbook.
- Multi-area PrintArea with excluded/private cells.
- Merged cells, tables, pivots, conditional formats, shapes, charts, images, camera pictures, and OLE objects.
- Repeated title rows/columns and impossible title bands.
- Manual horizontal/vertical page breaks.
- Missing fonts, Unicode, rich text, stale formulas, links, and calculation modes.
- Printer form availability/rejection and hard-margin variation.
- Light-gridline and edge-content trim regression.
- Multi-sheet per-sheet export and atomic failure semantics.

## 12. Performance and operational budgets

- Cache printer capabilities by printer/driver identity for one worker lifetime.
- Avoid repeated full-sheet COM scans; collect content inventory once per sheet.
- Bound candidate probes and pagination retries.
- Keep postflight rendering under a configurable megapixel limit while never silently reducing strict verification below its minimum DPI.
- Record timings for inventory, planning, COM probes, export, postflight, trim, and merge.
- A timeout at any stage follows existing process cleanup and atomic destination rules.

## 13. Rollback

- `quality_profile: legacy` restores the existing planner/export behavior.
- New modules are invoked through policy interfaces so strict paths can be disabled without reverting schema changes.
- Do not remove old config fields during the initial rollout.
- Preserve staging artifacts and decision manifests for failed strict jobs when diagnostic retention is enabled.

## 14. Definition of done

- Every requirement in `spec.md` maps to completed tasks and automated evidence in `task.md`.
- Unit, mocked COM, and Windows/Excel integration suites pass.
- Sample config and README describe strict, balanced, and legacy behavior accurately.
- Strict output is committed only after successful per-sheet and final-document postflight.
- No known silent content-loss path remains.

