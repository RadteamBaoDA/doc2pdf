# Excel Quality-First Enhancement Specification

**Status:** Proposed implementation baseline  
**Scope:** Excel-to-PDF conversion only  
**Primary implementation:** `src/core/excel_converter.py`, `src/config.py`, `src/core/pdf_processor.py`  
**Related documents:** [Implementation plan](implementation_plan.md), [Task checklist](task.md)

## 1. Purpose

This specification upgrades Excel conversion from a geometry-first page planner into an end-to-end quality-first pipeline. The result must preserve authored print layouts, avoid silent content loss, choose a consistent page size for each sheet, retain readable text and images, and prove that the final PDF matches the selected layout before it replaces the destination.

The existing 2D paper/orientation planner is the baseline, not a component to rewrite. It already measures width and height, probes printer-supported paper/orientation pairs, applies numeric Zoom, enforces `min_shrink_factor`, supports pagination, and performs COM readback. This enhancement surrounds that planner with stronger content discovery, sheet-level planning, printer geometry, staging, postflight, and evidence.

## 2. Product modes

Two behaviors that were previously mixed together are separated explicitly:

1. **Authored layout:** the workbook has a valid print setup. The converter preserves it exactly unless the user explicitly requests optimization.
2. **Smart layout:** the workbook has no usable print setup. The converter chooses paper, orientation, zoom, titles, and pagination using the quality constraints in this specification.

The converter must never silently move a workbook from authored-layout behavior to smart-layout behavior.

## 3. Goals

- Preserve valid authored `PrintArea`, paper, orientation, margins, zoom, print titles, manual page breaks, headers, footers, colors, objects, and page order.
- Automatically choose one consistent paper/orientation policy per sheet when smart layout is required.
- Keep effective 2D scale at or above 90% by default; paginate instead of shrinking further.
- Prefer OCR-readable A4/A3 pages over very large pages that downstream OCR would downsample.
- Detect all printable content, including content not represented by ordinary non-empty cells.
- Keep merged ranges, tables, shapes, charts, and repeated title bands intact across chunk boundaries.
- Verify the final exported PDF, not only the pre-export COM state.
- Preserve atomic destination semantics for every failure policy.
- Produce structured evidence explaining every layout and fallback decision.

## 4. Non-goals

- Reimplementing Excel rendering outside desktop Excel.
- Executing workbook macros or untrusted UDF code.
- Refreshing external data connections without an explicit policy.
- Guaranteeing PDF/A compliance using Excel export alone.
- Raster-stitching pages; any stitch mode must preserve vector/text content.
- Changing Word or PowerPoint conversion behavior.

## 5. Terminology

- **Authored PageSetup:** a valid print configuration already stored in the workbook.
- **Region:** one printable rectangular area derived from an authored PrintArea or strict content discovery.
- **Chunk:** a logical row group created by `row_dimensions`; it may occupy multiple physical PDF pages.
- **Candidate:** one printer-accepted paper/orientation combination plus its usable geometry and predicted pagination.
- **LayoutDecision:** immutable record of the chosen layout and the evidence used to choose it.
- **Quality floor:** minimum permitted effective scale, default 90%.
- **Postflight:** validation performed against the final PDF produced by Excel.

## 6. Configuration contract

Existing fields remain supported. New fields are additive unless a migration is explicitly stated.

```yaml
excel:
  quality_profile: strict                 # strict | balanced | legacy
  layout_policy: preserve_authored        # preserve_authored | optimize_missing | force_optimize
  page_size_scope: sheet                  # sheet | chunk
  orientation: auto                       # auto | portrait | landscape
  min_shrink_factor: 0.90
  oversized_action: paginate              # paginate | error | skip | warn
  horizontal_overflow_strategy: paginate  # paginate | error | one_logical_page | vector_stitch

  row_dimensions: null                    # null natural, 0 try whole region, N soft maximum
  print_area_policy: preserve_strict       # preserve_strict | expand_visible_objects | auto
  manual_page_break_policy: preserve       # preserve | reset
  print_title_rows: null
  print_title_columns: null

  preferred_papers: [A4, A3]
  allowed_papers: null                     # null means all mapped printer forms
  max_page_dimension_in: 24
  max_page_area_in2: 300
  avoid_horizontal_pagination: true

  min_effective_font_pt: 10
  min_effective_image_dpi: 150
  print_quality: standard
  draft_mode: false
  color_policy: preserve                   # preserve | force_color | black_and_white

  metadata_header_policy: preserve         # preserve | append | replace
  metadata_header: true
  ocr_sheet_name_label: false
  is_write_file_path: false

  calculation_policy: saved_cache          # saved_cache | calculate | full_rebuild
  external_link_policy: never_refresh      # never_refresh | refresh_allowed
  printer_policy: required                 # required | configured_fallback | system_default
  printer_name: "Microsoft Print to PDF"

  postflight_policy: strict                # strict | warn | disabled
  trim_policy: disabled                    # disabled | cropbox | physical
```

### 6.1 Profile behavior

- `strict` applies the defaults shown above and rejects uncertain or unverifiable output.
- `balanced` may use `print_area_policy: expand_visible_objects`, `postflight_policy: warn`, and `trim_policy: cropbox`.
- `legacy` preserves current compatibility behavior, including chunk-level selection and warning fallbacks where applicable.
- Explicit fields override profile defaults.

### 6.2 Compatibility and migration

- Keep `orientation: auto`, `min_shrink_factor: 0.90`, and `oversized_action: paginate` as quality-first defaults.
- Change the strict-profile effective defaults to `row_dimensions: null`, `ocr_sheet_name_label: false`, and `is_write_file_path: false`.
- Keep positive `row_dimensions=N`, but redefine the boundary as a soft maximum that may move to preserve atomic content.
- Continue accepting deprecated `page_shrink_threshold` only to emit a migration warning; it must not influence layout.
- Existing configurations without `quality_profile` initially load as `legacy` during one compatibility release. A later major release may default them to `strict` after an explicit migration notice.

## 7. Functional requirements

### XL-QF-001 — Authored layout preservation

- Detect whether each sheet has a valid authored PageSetup.
- Under `preserve_authored`, snapshot and fingerprint all relevant PageSetup properties before staging.
- Do not replace paper, orientation, margins, zoom, Fit flags, print titles, manual page breaks, page order, header, footer, color, or draft settings.
- An invalid authored PrintArea is a conversion error in strict mode; it must not fall back to automatic bounds.
- `force_optimize` is the only mode allowed to replace a valid authored layout.

### XL-QF-002 — Sheet-level page-size planning

- Under `page_size_scope: sheet`, select one paper and orientation for all regions/chunks belonging to a sheet.
- Evaluate the worst-case region/chunk and apply the winning decision consistently to the complete sheet output.
- Different sheets in the same workbook may use different page sizes.
- `page_size_scope: chunk` remains a legacy opt-in.
- Store the reason for the selected paper and any rejected alternatives in `LayoutDecision`.

### XL-QF-003 — Strict printable-content discovery

- Enumerate visible worksheet and chart-sheet types from `workbook.Sheets`.
- Preserve all authored multi-area PrintAreas in their authored order.
- Automatic discovery must account for values, formulas, merged areas, tables, pivot output, styled printable cells, conditional-format output, wrapped/rich/rotated text overflow, and printable objects.
- Object discovery must consider actual object rectangles and `PrintObject`; a failure reading one object must not abort scanning later objects.
- Region discovery returns a confidence/result object. Strict mode rejects uncertainty rather than silently excluding content.
- Hidden sheets remain excluded unless explicitly selected by configuration.

### XL-QF-004 — Object-aware chunk planning

- Treat merged areas, table row groups, shapes, charts, images, camera pictures, OLE objects, and title bands as atomic where splitting would clip or duplicate content.
- Build forbidden row/column boundaries and move a requested `row_dimensions` boundary to the nearest safe location.
- Reject a candidate if a single atomic element cannot fit at the quality floor.
- `row_dimensions=0` tries one logical region first, then permits quality-preserving pagination.
- `row_dimensions=null` uses natural Excel pagination.
- Positive `N` is a soft maximum, not permission to cut content.

### XL-QF-005 — Printer capability and printable geometry

- Use a configured printer in strict mode; do not silently switch to the system default.
- Record printer name, driver/version, port, advertised forms, and failure reason.
- Map the full supported `XlPaperSize`/`DMPAPER` catalog, including smaller forms and printer custom forms where Excel can address them.
- Set and read back both `PaperSize` and `Orientation` for every candidate.
- Obtain the hard printable/imageable area for the selected form and orientation; physical paper size minus configured margins is not sufficient.
- Retry margins using `max(configured margin, hard margin + safety padding)`.
- Reject non-finite or non-positive usable geometry; never clamp an impossible candidate to a small positive value.

### XL-QF-006 — OCR-aware candidate scoring

- Continue calculating `width_scale`, `height_scale`, and `effective_scale` in two dimensions.
- Enforce `quality_zoom = ceil(min_shrink_factor * 100)` and a maximum Zoom of 100%.
- Apply page dimension, page area, effective font, and effective image-DPI constraints before ranking candidates.
- Prefer candidates in this order:
  1. all hard quality constraints satisfied;
  2. fewer horizontal splits;
  3. fewer total pages;
  4. repeated key rows/columns available;
  5. preferred paper order;
  6. less whitespace;
  7. smaller paper area;
  8. deterministic paper/orientation key.
- Set and read back `PageOrder` for layouts that paginate both horizontally and vertically.
- A repeated title band that consumes the usable page at the quality floor makes the candidate invalid.

### XL-QF-007 — Actual Excel pagination probe

- Geometric page estimates are advisory only.
- After applying a candidate, force Excel to recompute print layout and inspect actual horizontal and vertical page breaks.
- Honor or reset manual page breaks according to `manual_page_break_policy`.
- If actual pagination differs from the prediction, re-score or retry with another candidate; a bounded 1–2% Zoom adjustment is allowed but must never cross the quality floor.
- Store predicted and actual page grids separately.

### XL-QF-008 — Immutable source and metadata fidelity

- Do not insert rows or cells into the source sheet before staging.
- Strict mode disables the OCR sheet-name row and file-path row.
- Metadata uses existing headers/footers or a PDF overlay; it must not change cell layout.
- `preserve` leaves authored headers and footers unchanged.
- `append` retains authored content, escapes Excel header control characters, observes Excel length limits, and verifies header/footer margins.
- `replace` is explicit and must not clear unspecified footer/header sections.
- Apply metadata before final PageSetup verification.

### XL-QF-009 — Calculation, fonts, color, and print-quality preflight

- `saved_cache` exports the workbook's stored results without recalculation and records that choice.
- `calculate` calculates open workbooks without refreshing external links.
- `full_rebuild` rebuilds dependencies and waits for calculation completion.
- Record external links, connections, calculation mode/state, formula errors, macro/UDF restrictions, and freshness policy.
- Detect unavailable fonts or probable font substitution before export; strict mode fails when substitution violates the font-size/readability constraint.
- Strict mode requires standard export quality, `Draft=False`, and policy-driven color handling.
- A global low-image-quality setting must not silently downgrade a strict Excel conversion.

### XL-QF-010 — Verified export pipeline

- Prefer exporting each already-verified staged sheet/region to its own PDF and merging those PDFs without rasterization.
- If a second cross-workbook copy is retained, repeat complete PageSetup readback after the copy and before export.
- Track expected page counts per sheet/region; one paginated sheet must not disable exact checks for unrelated strict sheets.
- Preserve sheet order and authored multi-area order.
- Export and merge failures must leave the existing destination unchanged.

### XL-QF-011 — PDF quality postflight

- Create a source manifest for every sheet/region containing boundary cells, sentinel text where available, printable object count, expected page grid, paper, orientation, and Zoom.
- Validate the final PDF for:
  - non-empty document and expected per-sheet page ranges;
  - MediaBox/CropBox and rotation matching the decision;
  - unexpected blank pages;
  - searchable text and boundary/sentinel coverage;
  - minimum and percentile effective font size;
  - clipping and protected edge ink;
  - image DPI and excessive full-page rasterization;
  - maximum page dimensions and page area.
- Strict postflight failure prevents atomic commit. `warn` records all failures but permits commit. `disabled` is legacy-only.

### XL-QF-012 — Safe trimming

- Strict Excel conversion defaults to no trim.
- `cropbox` may be enabled with a higher-resolution render and protected-edge verification.
- `physical` trim requires explicit opt-in and before/after visual equivalence checks.
- Light gridlines, thin borders, annotations, existing CropBoxes, rotations, signatures, and blank pages must remain protected.
- Trim failure must preserve both the staged original and the existing destination behavior defined by the job runner.

### XL-QF-013 — Horizontal overflow strategies

- `paginate` is the quality-first default and repeats configured title columns when available.
- `error` fails when one-page-wide output cannot meet all quality constraints.
- `one_logical_page` permits a single wide page only within the configured page dimension/area and OCR budgets.
- `vector_stitch` exports vector page tiles and combines them into one logical wide page without rasterizing text, objects, or images.
- Stitch mode must define title de-duplication, page order, maximum PDF user-space size, and failure fallback before it can graduate from experimental status.

### XL-QF-014 — Structured diagnostics

Every layout attempt must emit a machine-readable `LayoutDecision` containing at least:

- workbook/sheet/region/chunk identifiers;
- authored versus smart-layout mode;
- content and object bounds;
- printer, driver, form, imageable area, orientation, and margins;
- width/height/effective scale and limiting axis;
- quality floor and final Zoom;
- predicted and actual horizontal/vertical page counts;
- print titles, manual-break policy, metadata policy, and PageOrder;
- rejected candidates and reasons;
- calculation/font/export/trim/postflight outcomes;
- all fallbacks, warnings, and quality-policy exceptions.

Human-readable logs must reference the same decision ID.

### XL-QF-015 — PDF/A and compliance

- Treat Excel `ExportAsFixedFormat` output as an ordinary PDF unless validated otherwise.
- PDF/A requests require a dedicated conversion/postprocessing component and a standards validator.
- Compliance failure follows the same atomic and strict/warn policies as quality postflight.

## 8. Failure and atomicity semantics

- `error`: fail the complete conversion and leave the destination untouched.
- `skip`: omit only the failed sheet, retain verified sheets, and record the omission in the manifest.
- `paginate`: retain the quality floor and permit additional horizontal/vertical pages.
- `warn`: allow output below a quality requirement only with a structured warning; it must never conceal missing content or failed COM readback.
- Any required PageSetup readback failure is fatal in strict mode.
- A missing/invalid visible sheet or uncertain region is never silently skipped in strict mode.
- All export, merge, trim, postflight, and compliance work occurs on unique staging paths before one final atomic destination replacement.

## 9. Acceptance criteria

- A workbook containing A4 portrait, A3 landscape, and an unconfigured wide sheet preserves the first two authored layouts and smart-selects only the third.
- Every chunk belonging to one sheet uses the same paper/orientation when `page_size_scope: sheet`.
- No chunk boundary crosses a merged range or printable object.
- Invalid authored PrintArea fails closed under `preserve_strict`.
- Chart sheets and worksheet charts survive export in workbook order.
- Actual Excel page breaks agree with the recorded `LayoutDecision`.
- No selected output is below 90% effective scale, 10pt effective font, or 150 DPI effective image quality under strict defaults.
- Multi-sheet export receives per-sheet postflight; pagination in one sheet does not weaken validation of another.
- Light gridlines and edge sentinels survive any enabled trim.
- Calculation policy and font substitution state are visible in the manifest.
- Every strict failure leaves an existing destination byte-for-byte unchanged.
- Windows integration fixtures exercise wide, tall, wide+tall, authored layouts, merged cells, charts, shapes, titles, manual breaks, missing fonts, stale formulas, and printer variations.

## 10. Rollout and compatibility gates

1. Land schema and decision logging without changing legacy output.
2. Implement strict behavior behind `quality_profile: strict`.
3. Run legacy and strict modes in shadow comparison on the fixture corpus.
4. Make strict the sample-config default only after Windows/Excel acceptance passes.
5. Keep a documented `legacy` rollback profile for at least one major release.
6. Graduate `vector_stitch` only after vector/text preservation and OCR-budget tests pass.

