"""
Excel to PDF Converter using pywin32 COM.

Features:
- Smart Page Size for OCR optimization
- Dynamic page width based on column count
- Configurable row dimensions for vertical pagination
- Metadata headers (sheet name, row range, filename)
"""
import math
import os
import tempfile
import time
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple
import win32com.client
import win32print
import pythoncom
import win32process
import dataclasses
from contextlib import contextmanager

from ..base import Converter
from ...config import (
    PDFConversionSettings, ExcelSettings, get_excel_sheet_settings,
    get_reporting_config,
)
from ...utils.logger import logger
from ...utils.process_manager import ProcessRegistry
from .chunking import SafeChunkPlanner
from .content import PrintableContentResolver
from .extensions import require_supported_extensions
from .layout import AuthoredLayoutInspector
from .models import (
    ConversionManifest, ExportedSheetArtifact, LayoutDecision,
    PdfPostflightResult, QualityLayoutCandidate, manifest_name,
)
from .pagination import ExcelPaginationProbe
from .printer import PrinterCapabilityProvider
from .pdf_quality import PdfQualityExpectation, PdfQualityPostflight

# Excel constants from Object Model
xlTypePDF = 0
xlQualityStandard = 0
xlQualityMinimum = 1
xlLandscape = 2
xlPortrait = 1
xlPaperLetter = 1
xlPaperA4 = 9
xlPaperA3 = 8
xlPaperA2 = 66  # Driver-specific; used only after the active printer advertises A2
xlPaperTabloid = 3  # 11x17 in
xlPaperLegal = 5  # 8.5x14 in
xlPaperLedger = 4  # 17x11 in (Tabloid rotated)
xlPaperB4 = 12  # 9.84x13.9 in (JIS B4)
# Architecture sizes
xlPaperC = 24  # 17x22 in (Arch C)
xlPaperD = 25  # 22x34 in (Arch D)
xlPaperE = 26  # 34x44 in (Arch E)

# Worksheet visibility
xlSheetVisible = -1

# AutomationSecurity constants (msoAutomationSecurity)
msoAutomationSecurityForceDisable = 3

# CorruptLoad constants - for opening potentially corrupted files
xlNormalLoad = 0

# Win32 DeviceCapabilities constants. Keeping the numeric values local avoids
# depending on an extra pywin32 module just for these stable Win32 constants.
DC_PAPERS = 2
DC_PAPERSIZE = 3
DC_PAPERNAMES = 16


class OversizedSheetError(Exception):
    """Raised when a sheet is too large to print at acceptable quality."""
    pass


class COMDisconnectedError(Exception):
    """Raised when Excel COM object has disconnected (crashed or became unavailable)."""
    pass


@dataclasses.dataclass(frozen=True)
class SheetRegion:
    first_row: int
    first_col: int
    last_row: int
    last_col: int

    @property
    def is_empty(self) -> bool:
        """Document this Excel pipeline operation and its side effects."""
        return self.last_row < self.first_row or self.last_col < self.first_col


@dataclasses.dataclass(frozen=True)
class PaperForm:
    """A printer paper form with physical dimensions in inches."""

    paper_enum: int
    name: str
    width_inches: float
    height_inches: float

    @property
    def area(self) -> float:
        """Document this Excel pipeline operation and its side effects."""
        return self.width_inches * self.height_inches


@dataclasses.dataclass(frozen=True)
class LayoutCandidate:
    """Pure-data quality and pagination metrics for one paper/orientation."""

    form: PaperForm
    orientation: int
    usable_width_inches: float
    usable_height_inches: float
    margins_points: Tuple[float, float, float, float]
    width_scale: float
    height_scale: float
    effective_scale: float
    max_zoom: int
    pages_wide: int
    pages_tall: int
    page_count: int
    limiting_axis: str


STANDARD_PAPER_FORMS: Tuple[PaperForm, ...] = (
    PaperForm(xlPaperLetter, "Letter", 8.5, 11.0),
    PaperForm(xlPaperLegal, "Legal", 8.5, 14.0),
    PaperForm(xlPaperA4, "A4", 8.27, 11.69),
    PaperForm(xlPaperB4, "B4", 9.84, 13.90),
    PaperForm(xlPaperA3, "A3", 11.69, 16.54),
    PaperForm(xlPaperTabloid, "Tabloid", 11.0, 17.0),
    PaperForm(xlPaperLedger, "Ledger", 17.0, 11.0),
    PaperForm(xlPaperA2, "A2", 16.54, 23.39),
    PaperForm(xlPaperC, "C", 17.0, 22.0),
    PaperForm(xlPaperD, "D", 22.0, 34.0),
    PaperForm(xlPaperE, "E", 34.0, 44.0),
)


class ExcelConverter(Converter):
    """
    Converter for Excel documents (.xlsx, .xls, .xlsm, .xlsb) to PDF.
    
    Features Smart Page Size for OCR optimization - ensures all columns
    are readable by OCR tools like miner U, Deepseek OCR, RAGFlow deepdoc.
    """
    
    # Geometry defaults (inches)
    MIN_PAGE_WIDTH_INCHES = 8.5
    DEFAULT_PAGE_HEIGHT_INCHES = 11.0
    POINTS_PER_INCH = 72
    
    # Search Constants
    xlByRows = 1
    xlByColumns = 2
    xlPrevious = 2

    def __init__(self, process_recorder: Optional[Callable[[int], None]] = None):
        """Document this Excel pipeline operation and its side effects."""
        self._process_recorder = process_recorder
        self._paper_forms_cache: Dict[str, Tuple[PaperForm, ...]] = {}
        self._paper_probe_cache: Dict[
            Tuple[str, int, int, Tuple[float, ...], bool],
            Optional[Tuple[float, float, float, float]],
        ] = {}
        self._printer_capabilities = PrinterCapabilityProvider()
        self._require_isolated_process = False
        self._phase_timings: Dict[str, float] = {}
        self._manifest_postprocess_updater: Optional[
            Callable[[Optional[PdfPostflightResult], Dict[str, float]], None]
        ] = None

    @contextmanager
    def _timed_phase(self, name: str):
        started = time.perf_counter()
        try:
            yield
        finally:
            self._phase_timings[name] = self._phase_timings.get(name, 0.0) + (
                time.perf_counter() - started
            )

    def finalize_postprocess_evidence(
        self,
        postflight: Optional[PdfPostflightResult],
        timings: Dict[str, float],
    ) -> None:
        """Add trim/final-validation evidence after the converter returns."""
        if self._manifest_postprocess_updater is not None:
            self._manifest_postprocess_updater(postflight, timings)
    
    def convert(
        self, 
        input_path: Path, 
        output_path: Optional[Path] = None, 
        settings: Optional[PDFConversionSettings] = None,
        on_progress: Optional[Callable[[float], None]] = None,
        base_path: Optional[Path] = None,
        runtime_evidence: Optional[Dict[str, Any]] = None,
    ) -> Path:
        """
        Convert Excel file to PDF using COM automation.
        
        Args:
            input_path: Path to source Excel file
            output_path: Path to destination PDF file
            settings: PDFConversionSettings object containing conversion configuration
            on_progress: Optional callback for partial progress (0.0 to 1.0)
            base_path: Optional root directory for relative path matching
        """    
        input_file = input_path.resolve()
        if not input_file.exists():
            raise FileNotFoundError(f"Input file not found: {input_file}")
            
        if output_path:
            out_file = output_path.resolve()
        else:
            out_file = input_file.with_suffix(".pdf")
            
        # Ensure output directory exists
        out_file.parent.mkdir(parents=True, exist_ok=True)
            
        settings = settings or PDFConversionSettings()
        excel_settings = settings.excel or ExcelSettings()
        # Profile selection determines whether the evidence-backed pipeline is used.
        
        logger.info(f"Converting '{input_file.name}' to PDF...")
        logger.debug(f"Settings: {settings}")

        start_time = time.time()
        
        # Ensure CoInitialize is called for this thread
        pythoncom.CoInitialize()
        
        try:
            if excel_settings.quality_profile != "legacy":
                # Strict/balanced conversion is isolated so legacy behavior remains stable.
                return self._convert_quality(
                    input_file, out_file, settings, on_progress, base_path,
                    runtime_evidence,
                )
            with self._excel_application() as excel:
                workbook = None
                temp_sheets_to_delete = []
                final_sheets_to_process = []
                
                try:
                    # Open Workbook with all parameters to suppress dialogs
                    # UpdateLinks=0: Don't update/prompt about external links
                    # ReadOnly=True: Open read-only for safety
                    # Format=None: Auto-detect delimiter format
                    # Password="": No password prompt
                    # WriteResPassword="": No write-reservation password prompt
                    # IgnoreReadOnlyRecommended=True: Ignore read-only recommendation
                    # Origin=None: Auto-detect origin
                    # Delimiter=None: Auto-detect delimiter
                    # Editable=False: Don't allow editing (no edit prompt)
                    # Notify=False: Don't notify about file reservation
                    # Converter=None: Auto-select converter
                    # AddToMru=False: Don't add to recent files
                    # Local=True: Use local settings without prompts
                    # CorruptLoad=xlNormalLoad: Normal load without repair dialog
                    workbook = excel.Workbooks.Open(
                        str(input_file), 
                        UpdateLinks=0,
                        ReadOnly=True,
                        Format=None,
                        Password="",
                        WriteResPassword="",
                        IgnoreReadOnlyRecommended=True,
                        Origin=None,
                        Delimiter=None,
                        Editable=False,
                        Notify=False,
                        Converter=None,
                        AddToMru=False,
                        Local=True,
                        CorruptLoad=xlNormalLoad
                    )
                    
                    # Get sheets to process
                    sheets_to_export = self._get_sheets_to_export(workbook, excel_settings)
                    
                    if not sheets_to_export:
                        logger.warning(f"No visible sheets found in {input_file.name}")
                        raise ValueError(f"No visible sheets to export in {input_file.name}")
                    
                    # Calculate progress weight per sheet
                    total_sheets = len(sheets_to_export)
                    sheet_weight = 1.0 / total_sheets if total_sheets > 0 else 0
                    
                    # Apply optional workbook mutations before final region measurement.
                    skipped_sheets = []  # Track skipped oversized sheets
                    expected_page_count = 0
                    exact_page_count = True
                    for sheet in sheets_to_export:
                        sheet_output_start = len(final_sheets_to_process)
                        sheet_expected_page_count = expected_page_count
                        sheet_exact_page_count = exact_page_count
                        try:
                            # Get sheet-specific settings
                            # Note: Arguments are (sheet_name, base_settings, input_path, base_path)
                            sheet_settings = get_excel_sheet_settings(sheet.Name, settings, input_file, base_path)
                            sheet_excel_settings = sheet_settings.excel or excel_settings
                            
                            logger.debug(f"Sheet '{sheet.Name}' settings: row_dimensions={sheet_excel_settings.row_dimensions}")
                            
                            # Insert OCR sheet name label if enabled
                            if sheet_excel_settings.ocr_sheet_name_label:
                                self._insert_sheet_name_label(sheet, sheet.Name)
                            
                            # Path/label mutations must happen before final measurement.
                            if sheet_excel_settings.is_write_file_path:
                                auto_regions = self._resolve_sheet_regions(sheet, "auto")
                                if not auto_regions:
                                    auto_regions = [SheetRegion(1, 1, 1, 1)]
                                last_row = max(region.last_row for region in auto_regions)
                                last_col = max(region.last_col for region in auto_regions)
                                last_row = self._insert_file_path_row(sheet, input_file, last_row, last_col, base_path)

                            regions = self._resolve_sheet_regions(
                                sheet, sheet_excel_settings.print_area_policy
                            )
                            if not regions:
                                skipped_sheets.append(sheet.Name)
                                if on_progress:
                                    on_progress(sheet_weight)
                                continue
                            work_regions = []
                            for region in regions:
                                row_limit = sheet_excel_settings.row_dimensions
                                if row_limit and row_limit > 0:
                                    for first in range(region.first_row, region.last_row + 1, row_limit):
                                        work_regions.append(SheetRegion(
                                            first, region.first_col,
                                            min(region.last_row, first + row_limit - 1), region.last_col,
                                        ))
                                else:
                                    work_regions.append(region)
                            weight = sheet_weight / len(work_regions)
                            for region in work_regions:
                                new_sheet = self._copy_region_sheet(workbook, sheet, region)
                                temp_sheets_to_delete.append(new_sheet)
                                chunk_settings = ExcelSettings(**dataclasses.asdict(sheet_excel_settings))
                                if sheet_excel_settings.row_dimensions is not None:
                                    chunk_settings.row_dimensions = 0
                                    if sheet_excel_settings.oversized_action == "paginate":
                                        # A fixed row chunk is now a maximum region,
                                        # not a promise that Excel must shrink it to
                                        # exactly one physical page.
                                        exact_page_count = False
                                    else:
                                        expected_page_count += 1
                                elif sheet_excel_settings.oversized_action == "error":
                                    # Strict error mode promises a one-page-high
                                    # layout even when row_dimensions is null.
                                    expected_page_count += 1
                                else:
                                    exact_page_count = False
                                region_range = new_sheet.Range(
                                    new_sheet.Cells(region.first_row, region.first_col),
                                    new_sheet.Cells(region.last_row, region.last_col),
                                )
                                self._apply_page_setup(
                                    new_sheet, chunk_settings, input_file.name,
                                    region.last_col,
                                    content_width_points=float(region_range.Width),
                                    content_height_points=float(region_range.Height),
                                )
                                if sheet_excel_settings.metadata_header:
                                    self._apply_metadata_header(
                                        new_sheet, sheet_excel_settings, input_file.name,
                                        f"{region.first_row}-{region.last_row}", left_text=sheet.Name,
                                    )
                                final_sheets_to_process.append(new_sheet)
                                if on_progress:
                                    on_progress(weight)
                        
                        except OversizedSheetError:
                            # Skip is atomic at sheet level: discard any chunks
                            # staged before a later chunk proved oversized.
                            del final_sheets_to_process[sheet_output_start:]
                            expected_page_count = sheet_expected_page_count
                            exact_page_count = sheet_exact_page_count
                            skipped_sheets.append(sheet.Name)
                            if on_progress:
                                on_progress(sheet_weight)
                            continue
                    
                    # Log skipped sheets summary
                    if skipped_sheets:
                        logger.warning(f"Skipped {len(skipped_sheets)} oversized sheet(s): {', '.join(skipped_sheets)}")
                    
                    # Export to PDF
                    if final_sheets_to_process:
                        import os
                        import tempfile
                        from pypdf import PdfReader
                        fd, stage_name = tempfile.mkstemp(
                            prefix=f".{out_file.name}.", suffix=".stage.pdf", dir=str(out_file.parent)
                        )
                        os.close(fd)
                        stage = Path(stage_name)
                        try:
                            stage.unlink(missing_ok=True)
                            self._export_to_pdf(workbook, final_sheets_to_process, str(stage), settings)
                            if not stage.is_file() or stage.stat().st_size == 0:
                                raise ValueError("Excel export did not create a nonempty PDF")
                            exported = PdfReader(str(stage))
                            if not exported.pages:
                                raise ValueError("Excel export created a PDF with no pages")
                            if exact_page_count and len(exported.pages) != expected_page_count:
                                raise ValueError(
                                    f"Excel exported {len(exported.pages)} pages; expected exactly "
                                    f"{expected_page_count} one-page regions"
                                )
                            os.replace(stage, out_file)
                        finally:
                            stage.unlink(missing_ok=True)
                        elapsed = time.time() - start_time
                        mins, secs = divmod(int(elapsed), 60)
                        logger.success(f"Successfully converted: {out_file} [{mins:02d}:{secs:02d}]")
                    else:
                        raise ValueError("Workbook contains no exportable content")
                    
                except Exception as e:
                    logger.error(f"Failed to convert {input_file.name}: {e}")
                    # Check if it's a COM disconnection - provide clearer message
                    if isinstance(e, COMDisconnectedError):
                        logger.warning("Excel crashed or became unavailable. This file will be skipped.")
                    raise
                finally:
                    # Cleanup temps
                    if temp_sheets_to_delete:
                        try:
                            excel.DisplayAlerts = False
                        except:
                            pass
                        for t in temp_sheets_to_delete:
                            try:
                                t.Delete()
                            except:
                                pass
                    
                    if workbook:
                        try:
                            workbook.Close(SaveChanges=False)
                        except:
                            pass
        finally:
            pythoncom.CoUninitialize()
            
        return out_file

    def _convert_quality(
        self,
        input_file: Path,
        out_file: Path,
        settings: PDFConversionSettings,
        on_progress: Optional[Callable[[float], None]],
        base_path: Optional[Path],
        runtime_evidence: Optional[Dict[str, Any]],
    ) -> Path:
        """Run the isolated strict/balanced pipeline and atomically stage output."""
        from pypdf import PdfReader, PdfWriter

        root_excel_settings = settings.excel or ExcelSettings()
        total_started = time.perf_counter()
        self._phase_timings = {}
        self._manifest_postprocess_updater = None
        self._require_isolated_process = True
        require_supported_extensions(root_excel_settings, settings.compliance)
        decisions: List[LayoutDecision] = []
        artifacts: List[ExportedSheetArtifact] = []
        skipped: List[str] = []
        manifest_failures: List[str] = []
        final_postflight: Optional[PdfPostflightResult] = None
        reporting = get_reporting_config()
        reports_dir = Path(reporting.reports_dir)
        manifest_path: Optional[Path] = None

        def write_manifest() -> None:
            """Document this Excel pipeline operation and its side effects."""
            nonlocal manifest_path
            if not reporting.enabled:
                return
            manifest = ConversionManifest(
                workbook=str(input_file), output=str(out_file),
                profile=root_excel_settings.quality_profile,
                decisions=tuple(decisions), artifacts=tuple(artifacts),
                postflight=final_postflight, skipped_sheets=tuple(skipped),
                failures=tuple(manifest_failures),
                timings_ms={
                    name: round(seconds * 1000.0, 3)
                    for name, seconds in sorted(self._phase_timings.items())
                },
                runtime_evidence=dict(runtime_evidence or {}),
            )
            manifest_path = reports_dir / manifest_name(input_file, decisions)
            manifest.write_atomic(manifest_path)

        def update_postprocess_manifest(
            result: Optional[PdfPostflightResult], extra: Dict[str, float]
        ) -> None:
            nonlocal final_postflight
            if result is not None:
                final_postflight = result
            for name, seconds in extra.items():
                self._phase_timings[name] = self._phase_timings.get(name, 0.0) + seconds
            self._phase_timings["total"] = time.perf_counter() - total_started
            write_manifest()

        self._manifest_postprocess_updater = update_postprocess_manifest

        try:
            with self._excel_application() as excel:
                # Calculation and link policy are applied before workbook inspection.
                self._prepare_calculation(excel, root_excel_settings)
                workbook = None
                staged_sheets: List[Any] = []
                try:
                    with self._timed_phase("open_calculation"):
                        workbook = excel.Workbooks.Open(
                            str(input_file), UpdateLinks=0, ReadOnly=True, Format=None,
                            Password="", WriteResPassword="",
                            IgnoreReadOnlyRecommended=True, Origin=None,
                            Delimiter=None, Editable=False, Notify=False,
                            Converter=None, AddToMru=False, Local=True,
                            CorruptLoad=xlNormalLoad,
                        )
                        calculation_evidence = self._execute_calculation_policy(
                            excel, workbook, root_excel_settings
                        )
                    inspector = AuthoredLayoutInspector(input_file)
                    resolver = PrintableContentResolver()
                    with self._timed_phase("inventory"):
                        inventory = resolver.inventory_sheets(
                            workbook, root_excel_settings.sheet_name
                        )
                    visible = [item for item in inventory if item.visible]
                    # Workbook.Sheets order is retained for deterministic PDF page order.
                    if not visible:
                        raise ValueError("Workbook contains no visible exportable sheets")
                    page_cursor = 1
                    sheet_pdf_paths: List[Path] = []
                    all_sentinels: List[str] = []
                    with tempfile.TemporaryDirectory(
                        prefix=f".{out_file.stem}.excel-quality-",
                        dir=str(out_file.parent),
                    ) as staging_name:
                        staging_dir = Path(staging_name)
                        for item_index, sheet_info in enumerate(visible):
                            sheet = workbook.Sheets.Item(sheet_info.index)
                            sheet_settings = get_excel_sheet_settings(
                                sheet_info.name, settings, input_file, base_path
                            )
                            excel_settings = sheet_settings.excel or root_excel_settings
                            require_supported_extensions(
                                excel_settings, sheet_settings.compliance
                            )
                            try:
                                with self._timed_phase("printer_layout"):
                                    capability = self._printer_capabilities.enforce(
                                        excel, excel_settings
                                    )
                                snapshot = inspector.inspect(sheet)
                                if snapshot.classification == "invalid":
                                    raise ValueError(
                                        f"Sheet {sheet_info.name!r} has invalid authored PageSetup: "
                                        + "; ".join(snapshot.errors)
                                    )
                                if (
                                    excel_settings.quality_profile == "strict"
                                    and snapshot.classification == "authored"
                                    and snapshot.confidence == "uncertain"
                                ):
                                    raise ValueError(
                                        f"Sheet {sheet_info.name!r} authored PageSetup is uncertain"
                                    )
                                if snapshot.draft and excel_settings.quality_profile == "strict":
                                    raise ValueError(
                                        f"Sheet {sheet_info.name!r} authored layout enables Draft mode"
                                    )
                                source_min_font = self._font_preflight(
                                    workbook, sheet, excel_settings
                                )
                                preserve_authored = (
                                    snapshot.classification == "authored"
                                    and excel_settings.layout_policy != "force_optimize"
                                )
                                with self._timed_phase("staging"):
                                    staged, decision, sentinels, expected = self._stage_quality_sheet(
                                        workbook, sheet, sheet_info.index, input_file,
                                        excel_settings, capability, snapshot,
                                        preserve_authored, resolver,
                                        tuple(calculation_evidence),
                                        source_min_font,
                                    )
                                staged_sheets.extend(staged)
                                all_sentinels.extend(sentinels)
                                decisions.append(decision)
                                sheet_pdf = staging_dir / f"sheet-{item_index + 1:04d}.pdf"
                                with self._timed_phase("export"):
                                    self._export_quality_units(
                                        staged, sheet_pdf, sheet_settings
                                    )
                                reader = PdfReader(str(sheet_pdf))
                                page_count = len(reader.pages)
                                expectation = PdfQualityExpectation(
                                    expected_pages=expected,
                                    sentinels=tuple(sentinels),
                                    min_font_pt=excel_settings.min_effective_font_pt,
                                    min_image_dpi=excel_settings.min_effective_image_dpi,
                                    max_dimension_in=excel_settings.max_page_dimension_in,
                                    max_area_in2=excel_settings.max_page_area_in2,
                                    require_searchable_text=bool(sentinels),
                                )
                                with self._timed_phase("postflight"):
                                    result = PdfQualityPostflight().validate(
                                        sheet_pdf, expectation
                                    )
                                self._enforce_postflight(
                                    result, excel_settings, sheet_info.name
                                )
                                artifacts.append(ExportedSheetArtifact(
                                    decision.decision_id, sheet_info.name,
                                    str(out_file), page_cursor,
                                    page_cursor + page_count - 1,
                                ))
                                page_cursor += page_count
                                sheet_pdf_paths.append(sheet_pdf)
                            except OversizedSheetError as exc:
                                if excel_settings.oversized_action != "skip":
                                    raise
                                skipped.append(sheet_info.name)
                                logger.warning(
                                    f"Skipping sheet {sheet_info.name!r}: {exc}"
                                )
                            if on_progress:
                                on_progress((item_index + 1) / len(visible))
                        if not sheet_pdf_paths:
                            raise ValueError("Workbook contains no verified exportable sheets")
                        merged = staging_dir / "merged.pdf"
                        with self._timed_phase("merge"):
                            writer = PdfWriter()
                            for pdf_path in sheet_pdf_paths:
                                reader = PdfReader(str(pdf_path))
                                for page in reader.pages:
                                    writer.add_page(page)
                            with merged.open("wb") as stream:
                                writer.write(stream)
                        expected_total = sum(
                            artifact.last_page - artifact.first_page + 1
                            for artifact in artifacts
                        )
                        final_expectation = PdfQualityExpectation(
                            expected_pages=expected_total,
                            sentinels=tuple(dict.fromkeys(all_sentinels)),
                            min_font_pt=root_excel_settings.min_effective_font_pt,
                            min_image_dpi=root_excel_settings.min_effective_image_dpi,
                            max_dimension_in=root_excel_settings.max_page_dimension_in,
                            max_area_in2=root_excel_settings.max_page_area_in2,
                            require_searchable_text=bool(all_sentinels),
                        )
                        with self._timed_phase("postflight"):
                            final_postflight = PdfQualityPostflight().validate(
                                merged, final_expectation
                            )
                        self._enforce_postflight(
                            final_postflight, root_excel_settings, "final document"
                        )
                        os.replace(merged, out_file)
                finally:
                    with self._timed_phase("cleanup"):
                        try:
                            excel.DisplayAlerts = False
                        except Exception:
                            pass
                        for staged_sheet in reversed(staged_sheets):
                            try:
                                staged_sheet.Delete()
                            except Exception:
                                pass
                        if workbook is not None:
                            try:
                                workbook.Close(SaveChanges=False)
                            except Exception:
                                pass
            self._phase_timings["total"] = time.perf_counter() - total_started
            write_manifest()
            logger.success(
                f"Successfully converted: {out_file}"
                + (f" [manifest {manifest_path}]" if manifest_path else "")
            )
            return out_file
        except Exception as exc:
            self._phase_timings["total"] = time.perf_counter() - total_started
            manifest_failures.append(f"{type(exc).__name__}: {exc}")
            try:
                write_manifest()
            except Exception as manifest_exc:
                logger.error(f"Could not write Excel decision manifest: {manifest_exc}")
            raise

    @staticmethod
    def _prepare_calculation(app: Any, settings: ExcelSettings) -> None:
        """Document this Excel pipeline operation and its side effects."""
        app.AskToUpdateLinks = False
        if settings.calculation_policy == "saved_cache":
            try:
                app.Calculation = -4135  # xlCalculationManual
            except Exception as exc:
                if settings.quality_profile == "strict":
                    raise ValueError(f"Cannot select saved-cache calculation mode: {exc}") from exc

    @staticmethod
    def _execute_calculation_policy(
        app: Any, workbook: Any, settings: ExcelSettings,
    ) -> List[str]:
        """Document this Excel pipeline operation and its side effects."""
        evidence = [f"policy={settings.calculation_policy}"]
        try:
            links = workbook.LinkSources() or []
            evidence.append(f"external_links={len(links)}")
        except Exception:
            evidence.append("external_links=unavailable")
        try:
            evidence.append(f"connections={int(workbook.Connections.Count)}")
        except Exception:
            evidence.append("connections=unavailable")
        formula_errors = 0
        formula_inventory_complete = True
        try:
            for sheet in workbook.Worksheets:
                try:
                    formula_errors += int(
                        sheet.Cells.SpecialCells(-4123, 16).Count
                    )
                except Exception as exc:
                    # Excel raises when no matching formula-error cells exist.
                    if "No cells were found" not in str(exc):
                        formula_inventory_complete = False
        except Exception:
            formula_inventory_complete = False
        evidence.append(
            f"formula_errors={formula_errors}"
            if formula_inventory_complete else "formula_errors=unavailable"
        )
        evidence.append("macros=disabled")
        evidence.append("udf_execution=not_requested")
        evidence.append(
            "freshness=saved-cache"
            if settings.calculation_policy == "saved_cache"
            else "freshness=calculated-without-link-refresh"
        )
        if settings.external_link_policy != "never_refresh":
            raise ValueError("External-link refresh is unavailable in M0-M5")
        if settings.calculation_policy == "calculate":
            workbook.Calculate()
        elif settings.calculation_policy == "full_rebuild":
            app.CalculateFullRebuild()
        if settings.calculation_policy != "saved_cache":
            deadline = time.monotonic() + 300
            while int(getattr(app, "CalculationState", 0)) != 0:
                if time.monotonic() >= deadline:
                    raise TimeoutError("Excel calculation did not complete within 300 seconds")
                time.sleep(0.1)
        evidence.append(f"state={getattr(app, 'CalculationState', 'unknown')}")
        return evidence

    def _stage_quality_sheet(
        self,
        workbook: Any,
        sheet: Any,
        sheet_index: int,
        input_file: Path,
        excel_settings: ExcelSettings,
        printer_capability: Any,
        snapshot: Any,
        preserve_authored: bool,
        resolver: PrintableContentResolver,
        calculation_evidence: Tuple[str, ...],
        source_min_font: Optional[float],
    ) -> Tuple[List[Any], LayoutDecision, List[str], int]:
        """Document this Excel pipeline operation and its side effects."""
        strict = excel_settings.quality_profile == "strict"
        if (
            strict and source_min_font is not None
            and source_min_font < excel_settings.min_effective_font_pt
        ):
            raise ValueError(
                f"Sheet {sheet.Name!r} contains {source_min_font:.2f}pt text below "
                f"the {excel_settings.min_effective_font_pt:.2f}pt quality floor"
            )
        if preserve_authored or getattr(sheet, "Type", None) == -4109:
            copied = self._copy_whole_sheet(workbook, sheet)
            copied_snapshot = AuthoredLayoutInspector(input_file).inspect(copied)
            if preserve_authored and not self._layout_values_match(snapshot, copied_snapshot):
                raise ValueError(
                    f"Sheet {sheet.Name!r} PageSetup changed during staging"
                )
            self._apply_quality_metadata(
                copied, excel_settings, input_file.name, sheet.Name, ""
            )
            self._verify_metadata_margins(copied, excel_settings)
            with self._timed_phase("pagination"):
                evidence = self._probe_pagination(copied, excel_settings, None)
            sentinels = self._boundary_sentinels(sheet, ())
            decision = LayoutDecision(
                workbook=input_file.name, sheet=str(sheet.Name),
                sheet_index=sheet_index, mode="authored" if preserve_authored else "chart",
                region_ids=("authored",), chosen=None,
                predicted_grid=(evidence.pages_wide, evidence.pages_tall),
                actual_grid=(evidence.pages_wide, evidence.pages_tall),
                printer=printer_capability,
                authored_fingerprint=snapshot.fingerprint if preserve_authored else None,
                calculation_policy=excel_settings.calculation_policy,
                metadata_policy=excel_settings.metadata_header_policy,
                manual_break_policy=excel_settings.manual_page_break_policy,
                warnings=calculation_evidence,
            )
            return (
                [copied],
                decision,
                sentinels,
                max(1, evidence.pages_wide * evidence.pages_tall),
            )

        content = resolver.resolve(
            sheet, excel_settings.print_area_policy, strict=strict
        )
        if not content.regions:
            detail = "; ".join(content.errors) or "no printable content"
            raise ValueError(f"Sheet {sheet.Name!r}: {detail}")
        if strict and not content.certain:
            raise ValueError(
                f"Sheet {sheet.Name!r}: uncertain content discovery: "
                + "; ".join(content.errors)
            )
        atomic_ranges = self._atomic_row_ranges(sheet, content.objects)
        forbidden = SafeChunkPlanner.forbidden_row_boundaries(
            atomic_ranges, content.objects
        )
        chunks = SafeChunkPlanner().chunks(
            content.regions, excel_settings.row_dimensions, forbidden
        )
        staged: List[Any] = []
        measurements: List[Tuple[Any, SheetRegion, float, float]] = []
        for chunk in chunks:
            region = SheetRegion(
                chunk.first_row, chunk.first_col,
                chunk.last_row, chunk.last_col,
            )
            copied = self._copy_region_sheet(workbook, sheet, region)
            staged.append(copied)
            self._apply_quality_metadata(
                copied, excel_settings, input_file.name, str(sheet.Name),
                f"{chunk.first_row}-{chunk.last_row}",
            )
            cell_range = copied.Range(
                copied.Cells(region.first_row, region.first_col),
                copied.Cells(region.last_row, region.last_col),
            )
            measurements.append((
                copied, region, float(cell_range.Width), float(cell_range.Height)
            ))
        if not measurements:
            raise ValueError(f"Sheet {sheet.Name!r} produced no safe chunks")
        planning_settings = excel_settings
        if source_min_font:
            font_scale = excel_settings.min_effective_font_pt / source_min_font
            required_scale = max(excel_settings.min_shrink_factor, font_scale)
            if required_scale > 1.0 + 1e-9:
                raise ValueError(
                    f"Sheet {sheet.Name!r} cannot meet the effective font floor"
                )
            planning_settings = dataclasses.replace(
                excel_settings, min_shrink_factor=min(1.0, required_scale)
            )
        max_width = max(item[2] for item in measurements)
        max_height = max(item[3] for item in measurements)
        first_sheet, first_region, first_width, first_height = measurements[0]
        planning_width = (
            max_width if excel_settings.page_size_scope == "sheet" else first_width
        )
        planning_height = (
            max_height if excel_settings.page_size_scope == "sheet" else first_height
        )
        with self._timed_phase("printer_layout"):
            selected = self._apply_page_setup(
                first_sheet, planning_settings, input_file.name,
                first_region.last_col, planning_width, planning_height,
            )
        actual_grids = []
        for index, (copied, region, width, height) in enumerate(measurements):
            applied = selected
            if planning_settings.page_size_scope == "sheet" and index > 0:
                with self._timed_phase("printer_layout"):
                    applied = self._apply_page_setup(
                        copied, planning_settings, input_file.name,
                        region.last_col, width, height,
                        forced_layout=selected,
                    )
            elif index > 0:
                with self._timed_phase("printer_layout"):
                    applied = self._apply_page_setup(
                        copied, planning_settings, input_file.name,
                        region.last_col, width, height,
                    )
            with self._timed_phase("pagination"):
                evidence = self._probe_pagination(copied, planning_settings, applied)
            self._verify_metadata_margins(copied, planning_settings)
            actual_grids.append((evidence.pages_wide, evidence.pages_tall))
        preferred = {
            name.casefold(): rank
            for rank, name in enumerate(excel_settings.preferred_papers)
        }
        chosen = QualityLayoutCandidate(
            paper_enum=selected.form.paper_enum,
            paper_name=selected.form.name,
            orientation=selected.orientation,
            usable_width_inches=selected.usable_width_inches,
            usable_height_inches=selected.usable_height_inches,
            width_scale=selected.width_scale,
            height_scale=selected.height_scale,
            effective_scale=selected.effective_scale,
            zoom=int(first_sheet.PageSetup.Zoom),
            pages_wide=selected.pages_wide,
            pages_tall=selected.pages_tall,
            effective_font_pt=(
                source_min_font * int(first_sheet.PageSetup.Zoom) / 100.0
                if source_min_font is not None else None
            ),
            effective_image_dpi=None,
            preferred_rank=preferred.get(selected.form.name.casefold(), 1_000_000),
            repeated_titles=bool(
                excel_settings.print_title_rows or excel_settings.print_title_columns
            ),
        )
        decision = LayoutDecision(
            workbook=input_file.name, sheet=str(sheet.Name),
            sheet_index=sheet_index, mode="smart",
            region_ids=tuple(f"region-{region.order}" for region in content.regions),
            chosen=chosen,
            predicted_grid=(selected.pages_wide, selected.pages_tall),
            actual_grid=(
                max(grid[0] for grid in actual_grids),
                sum(grid[1] for grid in actual_grids),
            ),
            printer=printer_capability,
            calculation_policy=excel_settings.calculation_policy,
            metadata_policy=excel_settings.metadata_header_policy,
            manual_break_policy=excel_settings.manual_page_break_policy,
            warnings=tuple(content.errors) + calculation_evidence,
        )
        sentinels = self._boundary_sentinels(sheet, content.regions)
        expected_pages = sum(
            max(1, pages_wide * pages_tall)
            for pages_wide, pages_tall in actual_grids
        )
        return staged, decision, sentinels, expected_pages

    @staticmethod
    def _layout_values_match(first: Any, second: Any) -> bool:
        """Document this Excel pipeline operation and its side effects."""
        names = (
            "print_area", "paper_size", "orientation", "margins_points",
            "zoom", "fit_to_pages_wide", "fit_to_pages_tall",
            "print_title_rows", "print_title_columns", "page_order",
            "manual_row_breaks", "manual_column_breaks", "headers", "footers",
            "black_and_white", "draft",
        )
        return all(getattr(first, name) == getattr(second, name) for name in names)

    @staticmethod
    def _copy_whole_sheet(workbook: Any, source_sheet: Any) -> Any:
        """Document this Excel pipeline operation and its side effects."""
        last_sheet = workbook.Sheets(workbook.Sheets.Count)
        source_sheet.Copy(None, last_sheet)
        return workbook.Sheets(workbook.Sheets.Count)

    @staticmethod
    def _atomic_row_ranges(sheet: Any, objects: Tuple[Any, ...]) -> List[Tuple[int, int]]:
        """Document this Excel pipeline operation and its side effects."""
        ranges: List[Tuple[int, int]] = [
            (item.first_row, item.last_row) for item in objects
        ]
        try:
            merge_areas = sheet.UsedRange.MergeAreas
            for index in range(1, int(merge_areas.Count) + 1):
                area = merge_areas.Item(index)
                ranges.append((int(area.Row), int(area.Row + area.Rows.Count - 1)))
        except Exception:
            pass
        try:
            tables = sheet.ListObjects
            for index in range(1, int(tables.Count) + 1):
                area = tables.Item(index).Range
                ranges.append((int(area.Row), int(area.Row + area.Rows.Count - 1)))
        except Exception:
            pass
        try:
            title_rows = str(sheet.PageSetup.PrintTitleRows or "")
            if title_rows:
                area = sheet.Range(title_rows)
                ranges.append((int(area.Row), int(area.Row + area.Rows.Count - 1)))
        except Exception:
            pass
        return ranges

    def _probe_pagination(
        self, sheet: Any, settings: ExcelSettings,
        selected: Optional[LayoutCandidate],
    ):
        """Document this Excel pipeline operation and its side effects."""
        try:
            evidence = ExcelPaginationProbe().probe(
                sheet, settings.manual_page_break_policy
            )
        except Exception as exc:
            if settings.quality_profile == "strict":
                raise ValueError(
                    f"Sheet {sheet.Name!r}: cannot verify actual pagination: {exc}"
                ) from exc
            logger.warning(f"Sheet {sheet.Name!r}: pagination probe unavailable: {exc}")
            from .pagination import PaginationEvidence
            return PaginationEvidence(
                selected.pages_wide if selected else 1,
                selected.pages_tall if selected else 1,
                (), (), (), (),
            )
        if selected is not None:
            predicted = (selected.pages_wide, selected.pages_tall)
            actual = (evidence.pages_wide, evidence.pages_tall)
            if actual != predicted and settings.quality_profile == "strict":
                quality_zoom = int(math.ceil(settings.min_shrink_factor * 100 - 1e-9))
                current_zoom = int(sheet.PageSetup.Zoom)
                for adjustment in (1, 2):
                    retry_zoom = current_zoom - adjustment
                    if retry_zoom < quality_zoom:
                        break
                    self._required_set_page_property(sheet.PageSetup, "Zoom", retry_zoom)
                    retried = ExcelPaginationProbe().probe(
                        sheet, settings.manual_page_break_policy
                    )
                    if (retried.pages_wide, retried.pages_tall) == predicted:
                        return retried
                raise ValueError(
                    f"Sheet {sheet.Name!r}: predicted pagination {predicted} "
                    f"does not match Excel pagination {actual}"
                )
        return evidence

    @staticmethod
    def _escape_header_text(value: str) -> str:
        """Document this Excel pipeline operation and its side effects."""
        return str(value).replace("&", "&&")[:240]

    def _apply_quality_metadata(
        self, sheet: Any, settings: ExcelSettings, filename: str,
        sheet_name: str, row_label: str,
    ) -> None:
        """Document this Excel pipeline operation and its side effects."""
        if not settings.metadata_header or settings.metadata_header_policy == "preserve":
            return
        setup = sheet.PageSetup
        values = {
            "LeftHeader": self._escape_header_text(sheet_name),
            "CenterHeader": self._escape_header_text(row_label),
            "RightHeader": self._escape_header_text(filename) + " (Page &P)",
        }
        for name, value in values.items():
            if settings.metadata_header_policy == "append":
                existing = str(getattr(setup, name, "") or "")
                value = f"{existing} | {value}" if existing else value
            if len(value) > 255:
                raise ValueError(f"Excel header {name} exceeds 255 characters")
            self._required_set_page_property(setup, name, value)

    @staticmethod
    def _verify_metadata_margins(sheet: Any, settings: ExcelSettings) -> None:
        """Document this Excel pipeline operation and its side effects."""
        if not settings.metadata_header or settings.metadata_header_policy == "preserve":
            return
        try:
            header = float(sheet.PageSetup.HeaderMargin)
            top = float(sheet.PageSetup.TopMargin)
        except Exception as exc:
            if settings.quality_profile == "strict":
                raise ValueError(f"Cannot verify Excel header margins: {exc}") from exc
            return
        if header < 0 or top <= header:
            raise ValueError(
                f"Sheet {sheet.Name!r} header margin leaves no printable header band"
            )

    @staticmethod
    def _boundary_sentinels(sheet: Any, regions: Sequence[Any]) -> List[str]:
        """Document this Excel pipeline operation and its side effects."""
        values: List[str] = []
        for region in regions:
            for row, column in (
                (region.first_row, region.first_col),
                (region.last_row, region.last_col),
            ):
                try:
                    value = sheet.Cells(row, column).Text
                    text = str(value or "").strip()
                    if text and len(text) <= 200:
                        values.append(text)
                except Exception:
                    pass
        return list(dict.fromkeys(values))

    @staticmethod
    def _export_quality_units(
        sheets: Sequence[Any], output: Path, settings: PDFConversionSettings,
    ) -> None:
        """Document this Excel pipeline operation and its side effects."""
        from pypdf import PdfReader, PdfWriter

        unit_paths: List[Path] = []
        try:
            for index, sheet in enumerate(sheets):
                unit = output.with_name(f".{output.stem}.unit-{index + 1:04d}.pdf")
                unit.unlink(missing_ok=True)
                sheet.ExportAsFixedFormat(
                    Type=xlTypePDF, Filename=str(unit), Quality=xlQualityStandard,
                    IncludeDocProperties=settings.metadata.include_properties,
                    IgnorePrintAreas=False, OpenAfterPublish=False,
                )
                if not unit.is_file() or unit.stat().st_size == 0:
                    raise ValueError(f"Excel did not create PDF for sheet {sheet.Name!r}")
                unit_paths.append(unit)
            writer = PdfWriter()
            for unit in unit_paths:
                for page in PdfReader(str(unit)).pages:
                    writer.add_page(page)
            with output.open("wb") as stream:
                writer.write(stream)
        finally:
            for unit in unit_paths:
                unit.unlink(missing_ok=True)

    @staticmethod
    def _enforce_postflight(
        result: PdfPostflightResult, settings: ExcelSettings, label: str,
    ) -> None:
        """Document this Excel pipeline operation and its side effects."""
        if result.passed or settings.postflight_policy == "disabled":
            return
        message = f"Excel PDF postflight failed for {label}: " + "; ".join(result.failures)
        if settings.postflight_policy == "warn":
            logger.warning(message)
            return
        raise ValueError(message)

    @staticmethod
    def _font_preflight(
        workbook: Any, sheet: Any, settings: ExcelSettings,
    ) -> Optional[float]:
        """Document this Excel pipeline operation and its side effects."""
        if settings.quality_profile != "strict":
            return None
        requested: set[str] = set()
        sizes: List[float] = []
        inventory_errors: List[str] = []
        for cell_type in (2, -4123):  # constants, formulas
            try:
                cells = sheet.UsedRange.SpecialCells(cell_type).Cells
                for index in range(1, int(cells.Count) + 1):
                    font = cells.Item(index).Font
                    name = str(font.Name or "").strip()
                    if name:
                        requested.add(name.casefold())
                    try:
                        size = float(font.Size)
                        if size > 0:
                            sizes.append(size)
                    except (TypeError, ValueError):
                        pass
            except Exception as exc:
                if "No cells were found" not in str(exc):
                    inventory_errors.append(str(exc))
        try:
            shapes = sheet.Shapes
            for index in range(1, int(shapes.Count) + 1):
                shape = shapes.Item(index)
                try:
                    font = shape.TextFrame2.TextRange.Font
                    name = str(font.Name or "").strip()
                    if name:
                        requested.add(name.casefold())
                    try:
                        size = float(font.Size)
                        if size > 0:
                            sizes.append(size)
                    except (TypeError, ValueError):
                        pass
                except Exception:
                    continue
        except Exception:
            pass
        if inventory_errors:
            raise ValueError(
                "Cannot inventory workbook fonts: " + "; ".join(inventory_errors)
            )
        installed = ExcelConverter._installed_windows_fonts()
        missing = sorted(requested - installed)
        if missing:
            raise ValueError(
                f"Sheet {sheet.Name!r} uses unavailable fonts: {', '.join(missing)}"
            )
        return min(sizes) if sizes else None

    @staticmethod
    def _installed_windows_fonts() -> set[str]:
        """Document this Excel pipeline operation and its side effects."""
        import winreg

        result: set[str] = set()
        keys = (
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
            (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
        )
        for hive, path in keys:
            try:
                with winreg.OpenKey(hive, path) as key:
                    index = 0
                    while True:
                        try:
                            name, _, _ = winreg.EnumValue(key, index)
                        except OSError:
                            break
                        normalized = str(name).split("(", 1)[0].strip().casefold()
                        if normalized:
                            result.add(normalized)
                        index += 1
            except OSError:
                continue
        if not result:
            raise ValueError("Windows font inventory is unavailable")
        return result

    @contextmanager
    def _excel_application(self):
        """Context manager for Excel COM application lifecycle with retry on disconnection."""
        excel = None
        max_retries = 2
        
        for attempt in range(max_retries + 1):
            try:
                if attempt > 0:
                    logger.warning(f"Retrying Excel initialization (attempt {attempt + 1}/{max_retries + 1})...")
                    import time
                    time.sleep(1)  # Give OS time to clean up
                
                try:
                    with self._timed_phase("startup"):
                        excel = win32com.client.DispatchEx("Excel.Application")
                except Exception:
                    if self._require_isolated_process:
                        raise
                    # Compatibility for legacy callers and older test harnesses.
                    # Quality profiles require DispatchEx and never take this path.
                    excel = win32com.client.Dispatch("Excel.Application")
                if self._process_recorder:
                    try:
                        _, process_id = win32process.GetWindowThreadProcessId(excel.Hwnd)
                        self._process_recorder(int(process_id))
                    except Exception as exc:
                        logger.warning(f"Could not record isolated Excel process id: {exc}")
                
                # Validate connection immediately by accessing a property
                try:
                    _ = excel.Version
                except Exception as conn_err:
                    logger.warning(f"Excel connection validation failed: {conn_err}")
                    if attempt < max_retries:
                        excel = None
                        continue
                    raise
                
                excel.Visible = False
                # Suppress ALL alerts and dialogs - MUST be set before any other operations
                excel.DisplayAlerts = False
                excel.ScreenUpdating = False
                # Disable macro/automation security prompts
                excel.AutomationSecurity = msoAutomationSecurityForceDisable
                # Disable interactive mode - no user prompts (critical for printer dialogs)
                excel.Interactive = False
                # Disable events that might trigger dialogs
                excel.EnableEvents = False
                # Don't prompt about links
                excel.AskToUpdateLinks = False
                # Suppress clipboard prompts
                excel.CutCopyMode = False
                # NOTE: Do NOT set PrintCommunication=False here!
                # It prevents PageSetup changes (paper size, headers) from being applied.
                # Prevent Office feature installation dialogs
                try:
                    excel.FeatureInstall = 0  # msoFeatureInstallNone
                except:
                    pass
                # Disable file validation popups
                try:
                    excel.FileValidation = 0  # msoFileValidationSkip
                except:
                    pass
                
                # Try to set optimal printer (must be after DisplayAlerts=False)
                self._set_optimal_printer(excel)
                
                ProcessRegistry.register(excel)
                break  # Success, exit retry loop
                
            except Exception as e:
                if attempt < max_retries:
                    logger.warning(f"Excel initialization failed (attempt {attempt + 1}): {e}")
                    excel = None
                    continue
                logger.critical(f"Failed to initialize Microsoft Excel after {max_retries + 1} attempts: {e}")
                raise
        
        try:
            yield excel
        finally:
            if excel:
                ProcessRegistry.unregister(excel)
                with self._timed_phase("cleanup"):
                    self._safe_quit_excel(excel)

    def _kill_zombie_excel(self) -> None:
        """Compatibility no-op; global Excel termination is intentionally disabled."""
        logger.debug("Global Excel process termination is disabled")

    def _safe_quit_excel(self, excel, timeout_seconds: int = 5) -> None:
        """
        Safely quit Excel application.
        
        Note: COM objects are apartment-threaded - threading breaks COM marshaling.
        This method executes Quit() directly. Ensure Excel settings are properly
        configured (DisplayAlerts=False, Interactive=False) to prevent modal dialogs.
        
        Args:
            excel: Excel Application COM object
            timeout_seconds: Ignored (kept for API compatibility)
        """
        try:
            # Ensure DisplayAlerts is off before quitting
            try:
                excel.DisplayAlerts = False
            except:
                pass
            excel.Quit()
            logger.debug("Excel application closed successfully")
        except Exception as e:
            logger.debug(f"Excel.Quit() raised: {e}")
            # If Quit fails, the process might be zombie - will be cleaned on next retry
            pass

    def _set_optimal_printer(self, excel) -> None:
        """
        Attempt to set ActivePrinter to 'Microsoft Print to PDF' for better paper size support.
        Uses win32print API for reliable port detection, with brute-force fallback.
        
        IMPORTANT: Avoids printers with PORTPROMPT: port which would show a dialog.
        """
        target_name = "Microsoft Print to PDF"
        
        try:
            # Check if already active
            current = excel.ActivePrinter
            if target_name in current:
                logger.debug(f"'{target_name}' is already the active printer.")
                return
        except:
            pass

        # Strategy 1: Use win32print API for reliable port detection
        port_name = None
        try:
            handle = win32print.OpenPrinter(target_name)
            try:
                # Level 5 is lightweight, contains pPortName
                info = win32print.GetPrinter(handle, 5)
                port_name = info.get('pPortName', '')
                
                # Fallback to Level 2 if Level 5 didn't have port
                if not port_name:
                    info = win32print.GetPrinter(handle, 2)
                    port_name = info.get('pPortName', '')
            finally:
                win32print.ClosePrinter(handle)
        except Exception as e:
            logger.debug(f"OpenPrinter/GetPrinter API failed for '{target_name}': {e}")

        # CRITICAL: Skip if port is PORTPROMPT: - this WILL show a dialog
        if port_name and port_name.upper() == 'PORTPROMPT:':
            logger.warning(
                f"Printer '{target_name}' uses PORTPROMPT: which would show a dialog. "
                f"Skipping printer change to avoid UI interruption."
            )
            return

        # If we got a port name from the API, try it first
        candidates = []
        if port_name:
            candidates.append(f"{target_name} on {port_name}")
        
        # Strategy 2: Brute force Ne00-Ne99 as fallback (expanded range)
        for i in range(100):
            candidates.append(f"{target_name} on Ne{i:02d}:")
            
        # Strategy 3: Naked name (rare, but might work)
        candidates.append(target_name)
        
        success = False
        for candidate in candidates:
            try:
                # Ensure dialogs are suppressed before each attempt
                excel.DisplayAlerts = False
                excel.Interactive = False
                excel.ActivePrinter = candidate
                logger.info(f"Successfully switched ActivePrinter to: '{candidate}'")
                success = True
                break
            except Exception as e:
                # Only log first few failures to avoid spam
                if candidates.index(candidate) < 5:
                    logger.debug(f"Failed to set ActivePrinter to '{candidate}': {e}")
                
        if not success:
            logger.warning(
                f"Could not set ActivePrinter to '{target_name}'. "
                f"Using default printer. Large paper sizes (A3) may rely on default printer capabilities."
            )

    def _get_sheets_to_export(self, workbook, excel_settings: ExcelSettings) -> List:
        """Get list of sheets to export based on settings."""
        sheets = []
        
        for sheet in workbook.Worksheets:
            # Only process visible sheets
            if sheet.Visible != xlSheetVisible:
                continue
                
            # Filter by sheet name if specified
            if excel_settings.sheet_name:
                if sheet.Name != excel_settings.sheet_name:
                    continue
            
            # Validate that the sheet has a proper PageSetup object
            # Some sheet types (dialog sheets, macro sheets) may not support PageSetup
            if not self._has_valid_page_setup(sheet):
                logger.warning(f"Skipping sheet '{sheet.Name}': PageSetup not supported")
                continue
            
            sheets.append(sheet)
            logger.debug(f"Will export sheet: {sheet.Name}")
        
        return sheets

    def _has_valid_page_setup(self, sheet) -> bool:
        """
        Check if the sheet has a valid PageSetup object that can be modified.
        
        Some sheet types (Chart sheets accessed as Worksheets, Dialog sheets, 
        Macro 4.0 sheets) may not support standard PageSetup property modifications.
        The error manifests as properties showing '<unknown>' when accessed.
        
        Args:
            sheet: Excel sheet object to validate
            
        Returns:
            True if PageSetup is valid and modifiable, False otherwise
        """
        try:
            page_setup = sheet.PageSetup
            if page_setup is None:
                return False
            
            # Try to read a basic property to verify the object is valid
            # Reading Orientation is a safe test - it should return 1 (Portrait) or 2 (Landscape)
            # If the PageSetup is invalid, this will raise an exception or return an unusable value
            orientation = page_setup.Orientation
            
            # Check if we got a valid value (int for real COM, MagicMock for tests)
            # Invalid PageSetup objects typically raise exceptions or return '<unknown>' type
            if orientation is None:
                return False
            
            # For real COM objects, orientation should be 1 (Portrait) or 2 (Landscape)
            # For mocks in tests, orientation will be a MagicMock which is fine
            if isinstance(orientation, int) and orientation not in (1, 2):
                # Real COM object returned invalid orientation value
                logger.debug(f"Sheet '{sheet.Name}' has invalid PageSetup.Orientation: {orientation}")
                return False
            
            return True
        except Exception as e:
            # If we can't even read the Orientation property, the PageSetup is invalid
            logger.debug(f"Sheet '{sheet.Name}' PageSetup validation failed: {e}")
            return False

    def _calculate_smart_page_size(
        self, 
        sheet, 
        last_col_index: int,
        content_width_points: Optional[float] = None,
        content_height_points: Optional[float] = None
    ) -> Tuple[float, float]:
        """
        Calculate raw content width and height from Excel range geometry.
        
        Args:
            sheet: Excel Worksheet object
            last_col_index: The 1-based index of the last used column (e.g. 5 for Column E)
            content_width_points: Optional explicit content width in points.
            content_height_points: Optional explicit content height in points.
            
        Returns:
            Tuple of (content_width_inches, content_height_inches)
        """
        try:
            if last_col_index < 1 and not content_width_points:
                return self.MIN_PAGE_WIDTH_INCHES, self.DEFAULT_PAGE_HEIGHT_INCHES
                
            # Measure width
            if content_width_points is not None:
                # Use provided geometry points directly
                content_width_inches = content_width_points / self.POINTS_PER_INCH
            else:
                # Fallback: Measure width of Range(A:LastCol)
                first_col_char = "A"
                last_col_char = self._col_num_to_letter(last_col_index)
                col_range = sheet.Range(f"{first_col_char}1:{last_col_char}1")
                
                # .Width corresponds to the width in points of the range
                content_width_points = col_range.Width
                content_width_inches = content_width_points / self.POINTS_PER_INCH
            
            # Margins are deliberately not included here. The layout planner
            # subtracts the margins that Excel actually retains for each form.
            if content_height_points is not None and content_height_points > 0:
                content_height_inches = content_height_points / self.POINTS_PER_INCH
            else:
                content_height_inches = self.DEFAULT_PAGE_HEIGHT_INCHES
            
            logger.debug(
                f"Sheet '{sheet.Name}' (Cols 1-{last_col_index}): "
                f"Content: {content_width_inches:.2f}\" x {content_height_inches:.2f}\""
            )
            
            return content_width_inches, content_height_inches
            
        except Exception as e:
            logger.warning(f"Could not calculate smart page size: {e}")
            return self.MIN_PAGE_WIDTH_INCHES, self.DEFAULT_PAGE_HEIGHT_INCHES

    def _try_set_paper_size(self, page_setup, paper_enum: int, paper_name: str, timeout_seconds: int = 3) -> bool:
        """
        Safely attempt to set paper size.
        
        Note: COM objects are apartment-threaded - threading-based timeout breaks COM.
        This method executes the paper size assignment directly on the current thread.
        
        Args:
            page_setup: Excel PageSetup object
            paper_enum: Excel paper size constant (e.g., xlPaperA3)
            paper_name: Human-readable paper name for logging
            timeout_seconds: Ignored (kept for API compatibility)
            
        Returns:
            True if paper size was set successfully, False otherwise
        """
        # Validate page_setup object before attempting to set paper size
        if page_setup is None:
            logger.debug(f"Cannot set paper size to {paper_name}: PageSetup object is None")
            return False
        
        # Quick validation: try to access the object type
        try:
            app = page_setup.Application
            # Ensure dialogs are suppressed before setting paper size
            app.DisplayAlerts = False
            app.Interactive = False
            # Disable print communication to prevent printer dialogs
            try:
                app.PrintCommunication = False
            except:
                pass
        except Exception:
            logger.debug(f"Cannot set paper size to {paper_name}: PageSetup object is invalid")
            return False
        
        try:
            # Disable communication during change
            try:
                app.PrintCommunication = False
            except:
                pass
            
            page_setup.PaperSize = paper_enum
            
            # Re-enable to commit change
            try:
                app.PrintCommunication = True
            except:
                pass
            
            # Verify it was actually set
            if page_setup.PaperSize == paper_enum:
                return True
            else:
                logger.debug(f"Printer rejected paper size {paper_name} (Enum {paper_enum}). Trying next size...")
                return False
        except Exception as e:
            # Ensure PrintCommunication is re-enabled even on error
            try:
                app.PrintCommunication = True
            except:
                pass
            logger.debug(f"Failed to set paper size to {paper_name}: {e}")
            return False

    def _safe_com_call(self, func, timeout: int = 10, default=None):
        """
        Execute a COM call safely.
        
        Note: COM objects in Python/pywin32 are apartment-threaded and cannot be
        accessed from a different thread than the one that created them. Using
        threading for timeout protection breaks COM marshaling (causes '<unknown>' errors).
        
        This method executes the COM call directly on the current thread.
        For operations that might hang, ensure Excel settings are properly configured
        (DisplayAlerts=False, Interactive=False, etc.) to prevent modal dialogs.
        
        Args:
            func: Lambda or callable to execute
            timeout: Ignored (kept for API compatibility)
            default: Value to return if error occurs
            
        Returns:
            Result of func() or default if failed
            
        Raises:
            COMDisconnectedError: If the COM object has disconnected
        """
        try:
            return func()
        except Exception as e:
            error_str = str(e)
            error_code = getattr(e, 'args', [None])[0] if hasattr(e, 'args') and e.args else None
            
            # Check for disconnection errors
            # -2147417848 = RPC_E_DISCONNECTED (0x80010108)
            # -2147023174 = RPC_S_SERVER_UNAVAILABLE
            disconnection_codes = [-2147417848, -2147023174]
            disconnection_phrases = [
                'disconnected from its clients',
                'RPC server is unavailable',
                'Call was rejected by callee',
                'server threw an exception'
            ]
            
            is_disconnected = False
            if isinstance(error_code, int) and error_code in disconnection_codes:
                is_disconnected = True
            elif any(phrase.lower() in error_str.lower() for phrase in disconnection_phrases):
                is_disconnected = True
            
            if is_disconnected:
                logger.error(f"Excel COM connection lost: {e}")
                raise COMDisconnectedError(f"Excel has disconnected: {e}") from e
            
            logger.debug(f"COM operation failed: {e}")
            raise

    def _safe_set_page_property(self, page_setup, prop_name: str, value, timeout_seconds: int = 3) -> bool:
        """
        Safely set a PageSetup property.
        
        Note: COM objects are apartment-threaded - threading-based timeout breaks COM.
        This method executes the property assignment directly on the current thread.
        
        Args:
            page_setup: Excel PageSetup object
            prop_name: Name of the property to set (e.g., 'Orientation', 'Zoom')
            value: Value to assign to the property
            timeout_seconds: Ignored (kept for API compatibility)
            
        Returns:
            True if property was set successfully, False if failed
        """
        # Validate page_setup object before attempting to set property
        if page_setup is None:
            logger.debug(f"Cannot set PageSetup.{prop_name}: PageSetup object is None")
            return False
        try:
            _ = page_setup.Application
        except Exception:
            logger.debug(f"Cannot set PageSetup.{prop_name}: PageSetup object is invalid")
            return False
        try:
            setattr(page_setup, prop_name, value)
            return True
        except Exception as e:
            logger.debug(f"Failed to set PageSetup.{prop_name}: {e}")
            return False

    def _required_set_page_property(self, page_setup, prop_name: str, value) -> None:
        """Set, commit and read back a required Excel PageSetup property."""
        if not self._safe_set_page_property(page_setup, prop_name, value):
            raise ValueError(f"Excel rejected required PageSetup.{prop_name}={value!r}")
        try:
            page_setup.Application.PrintCommunication = True
            actual = getattr(page_setup, prop_name)
        except Exception as exc:
            raise ValueError(f"Cannot verify PageSetup.{prop_name}: {exc}") from exc
        if not self._page_property_matches(actual, value):
            raise ValueError(
                f"Excel did not retain PageSetup.{prop_name}: requested {value!r}, "
                f"read back {actual!r}"
            )

    @staticmethod
    def _page_property_matches(actual, expected) -> bool:
        """Compare COM readback values, including bool/int and float coercion."""
        if isinstance(expected, bool):
            return bool(actual) is expected
        if (
            isinstance(expected, (int, float))
            and not isinstance(expected, bool)
            and isinstance(actual, (int, float))
            and not isinstance(actual, bool)
        ):
            return math.isclose(float(actual), float(expected), abs_tol=1e-7)
        return actual == expected

    @staticmethod
    def _printer_advertises_a2(app) -> bool:
        """Document this Excel pipeline operation and its side effects."""
        try:
            printer_name = str(app.ActivePrinter or "").rsplit(" on ", 1)[0].strip()
            handle = win32print.OpenPrinter(printer_name)
            try:
                return any("A2" in str(form.get("Name", "")).upper() for form in win32print.EnumForms(handle))
            finally:
                win32print.ClosePrinter(handle)
        except Exception:
            return False

    @staticmethod
    def _device_paper_size_inches(size) -> Optional[Tuple[float, float]]:
        """Convert a DC_PAPERSIZE value (tenths of millimetres) to inches."""
        try:
            width, height = size
            width_inches = float(width) / 254.0
            height_inches = float(height) / 254.0
        except (TypeError, ValueError, OverflowError):
            return None
        if width_inches <= 0 or height_inches <= 0:
            return None
        return width_inches, height_inches

    def _get_printer_paper_forms(self, app) -> Tuple[PaperForm, ...]:
        """Return known Excel forms using dimensions advertised by the printer."""
        try:
            active_printer = str(app.ActivePrinter or "").strip()
        except Exception:
            active_printer = ""
        cache_key = active_printer or "<unknown-printer>"
        cached = self._paper_forms_cache.get(cache_key)
        if cached is not None:
            return cached

        known = {form.paper_enum: form for form in STANDARD_PAPER_FORMS}
        advertised: List[PaperForm] = []
        printer_name = active_printer.rsplit(" on ", 1)[0].strip()
        handle = None
        try:
            if not printer_name:
                raise ValueError("Excel ActivePrinter is empty")
            handle = win32print.OpenPrinter(printer_name)
            printer_info = win32print.GetPrinter(handle, 2)
            port_name = str(printer_info.get("pPortName", "") or "")
            paper_ids = win32print.DeviceCapabilities(
                printer_name, port_name, DC_PAPERS
            )
            paper_sizes = win32print.DeviceCapabilities(
                printer_name, port_name, DC_PAPERSIZE
            )
            paper_names = win32print.DeviceCapabilities(
                printer_name, port_name, DC_PAPERNAMES
            )
            for index, paper_id in enumerate(paper_ids or []):
                try:
                    paper_enum = int(paper_id)
                    fallback = known.get(paper_enum)
                    dimensions = self._device_paper_size_inches(paper_sizes[index])
                except (IndexError, TypeError, ValueError):
                    continue
                if dimensions is None:
                    continue
                try:
                    advertised_name = str(paper_names[index]).strip()
                except (IndexError, TypeError):
                    advertised_name = ""
                advertised.append(PaperForm(
                    paper_enum=paper_enum,
                    name=advertised_name or (fallback.name if fallback else f"Form-{paper_enum}"),
                    width_inches=dimensions[0],
                    height_inches=dimensions[1],
                ))
        except Exception as exc:
            logger.debug(f"Could not enumerate active printer forms: {exc}")
        finally:
            if handle is not None:
                try:
                    win32print.ClosePrinter(handle)
                except Exception:
                    pass

        if advertised:
            # DeviceCapabilities can contain duplicate IDs. Keep the first
            # driver-advertised definition so selection remains deterministic.
            unique: Dict[int, PaperForm] = {}
            for form in advertised:
                unique.setdefault(form.paper_enum, form)
            result = tuple(unique.values())
        else:
            result_list = [
                form for form in STANDARD_PAPER_FORMS
                if form.paper_enum != xlPaperA2
            ]
            if self._printer_advertises_a2(app):
                result_list.append(known[xlPaperA2])
            result = tuple(result_list)

        self._paper_forms_cache[cache_key] = result
        return result

    def _probe_paper_orientation(
        self,
        page_setup,
        form: PaperForm,
        orientation: int,
        requested_margins: Tuple[float, float, float, float],
        require_imageable_area: bool = False,
    ) -> Optional[Tuple[float, float, float, float]]:
        """Set a paper/orientation pair and return committed margins in points."""
        try:
            printer_name = str(page_setup.Application.ActivePrinter or "").casefold()
        except Exception:
            printer_name = "<unknown-printer>"
        cache_key = (
            printer_name,
            int(form.paper_enum),
            int(orientation),
            tuple(float(value) for value in requested_margins),
            require_imageable_area,
        )
        if cache_key in self._paper_probe_cache:
            return self._paper_probe_cache[cache_key]
        if not self._try_set_paper_size(page_setup, form.paper_enum, form.name):
            self._paper_probe_cache[cache_key] = None
            return None
        if not self._safe_set_page_property(page_setup, "Orientation", orientation):
            self._paper_probe_cache[cache_key] = None
            return None
        margins = requested_margins
        if require_imageable_area:
            try:
                active_printer = str(page_setup.Application.ActivePrinter or "")
                hard = self._printer_capabilities.hard_margins_points(
                    active_printer, form.paper_enum, orientation
                )
                safety = 6.0
                margins = tuple(
                    max(requested, required + safety)
                    for requested, required in zip(
                        requested_margins, hard, strict=True
                    )
                )
            except Exception as exc:
                logger.debug(
                    f"Cannot obtain hard margins for {form.name}: {exc}"
                )
                self._paper_probe_cache[cache_key] = None
                return None
        for prop_name, value in zip(
            ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"),
            margins,
            strict=True,
        ):
            if not self._safe_set_page_property(page_setup, prop_name, value):
                self._paper_probe_cache[cache_key] = None
                return None
        try:
            page_setup.Application.PrintCommunication = True
            if page_setup.PaperSize != form.paper_enum:
                self._paper_probe_cache[cache_key] = None
                return None
            if page_setup.Orientation != orientation:
                self._paper_probe_cache[cache_key] = None
                return None
            result = (
                float(page_setup.LeftMargin),
                float(page_setup.RightMargin),
                float(page_setup.TopMargin),
                float(page_setup.BottomMargin),
            )
            self._paper_probe_cache[cache_key] = result
            return result
        except Exception as exc:
            logger.debug(
                f"Printer rejected {form.name}/"
                f"{'landscape' if orientation == xlLandscape else 'portrait'}: {exc}"
            )
            self._paper_probe_cache[cache_key] = None
            return None

    @staticmethod
    def _page_span_count(
        content_inches: float,
        repeated_title_inches: float,
        usable_inches: float,
        zoom: int,
    ) -> int:
        """Estimate pages on one axis at a fixed Excel print Zoom."""
        scaled_capacity = usable_inches / max(zoom / 100.0, 0.01)
        content_inches = max(content_inches, 0.01)
        repeated_title_inches = min(
            max(repeated_title_inches, 0.0), content_inches
        )
        if repeated_title_inches > 0 and content_inches > repeated_title_inches:
            data_capacity = max(0.01, scaled_capacity - repeated_title_inches)
            data_extent = content_inches - repeated_title_inches
            return max(1, math.ceil(data_extent / data_capacity))
        return max(1, math.ceil(content_inches / max(scaled_capacity, 0.01)))

    @staticmethod
    def _build_layout_candidate(
        form: PaperForm,
        orientation: int,
        content_width_inches: float,
        content_height_inches: float,
        fit_tall: bool,
        margins_points: Tuple[float, float, float, float],
        quality_zoom: int,
        title_width_inches: float = 0.0,
        title_height_inches: float = 0.0,
    ) -> LayoutCandidate:
        """Build quality metrics without touching COM, enabling boundary tests."""
        left, right, top, bottom = margins_points
        if orientation == xlLandscape:
            physical_width = form.height_inches
            physical_height = form.width_inches
        else:
            physical_width = form.width_inches
            physical_height = form.height_inches
        usable_width = physical_width - ((left + right) / ExcelConverter.POINTS_PER_INCH)
        usable_height = physical_height - ((top + bottom) / ExcelConverter.POINTS_PER_INCH)
        if not all(
            math.isfinite(value) and value > 0
            for value in (usable_width, usable_height)
        ):
            raise ValueError("paper and margins produce invalid usable geometry")
        width_scale = usable_width / max(content_width_inches, 0.01)
        height_scale = (
            usable_height / max(content_height_inches, 0.01)
            if fit_tall else 1.0
        )
        effective_scale = min(1.0, width_scale, height_scale)
        max_zoom = max(
            1, min(100, int(math.floor((effective_scale * 100.0) + 1e-7)))
        )
        pages_wide = ExcelConverter._page_span_count(
            content_width_inches, title_width_inches,
            usable_width, quality_zoom,
        )
        pages_tall = ExcelConverter._page_span_count(
            content_height_inches, title_height_inches,
            usable_height, quality_zoom,
        )
        if effective_scale >= 1.0:
            limiting_axis = "none"
        elif fit_tall and height_scale <= width_scale:
            limiting_axis = "height"
        else:
            limiting_axis = "width"
        return LayoutCandidate(
            form=form,
            orientation=orientation,
            usable_width_inches=usable_width,
            usable_height_inches=usable_height,
            margins_points=margins_points,
            width_scale=width_scale,
            height_scale=height_scale,
            effective_scale=effective_scale,
            max_zoom=max_zoom,
            pages_wide=pages_wide,
            pages_tall=pages_tall,
            page_count=pages_wide * pages_tall,
            limiting_axis=limiting_axis,
        )

    @staticmethod
    def _fit_candidate_sort_key(candidate: LayoutCandidate) -> Tuple:
        """Document this Excel pipeline operation and its side effects."""
        return (
            -candidate.max_zoom,
            candidate.form.area,
            candidate.usable_width_inches * candidate.usable_height_inches,
            candidate.form.paper_enum,
            candidate.orientation,
        )

    @staticmethod
    def _select_fit_candidate(
        candidates: Sequence[LayoutCandidate], quality_zoom: int
    ) -> Optional[LayoutCandidate]:
        """Document this Excel pipeline operation and its side effects."""
        eligible = [
            candidate for candidate in candidates
            if candidate.max_zoom >= quality_zoom
        ]
        return min(eligible, key=ExcelConverter._fit_candidate_sort_key) if eligible else None

    @staticmethod
    def _select_paginated_candidate(
        candidates: Sequence[LayoutCandidate],
    ) -> LayoutCandidate:
        """Document this Excel pipeline operation and its side effects."""
        if not candidates:
            raise ValueError("No supported layout candidates")
        return min(candidates, key=lambda candidate: (
            candidate.page_count,
            candidate.form.area,
            candidate.form.paper_enum,
            candidate.orientation,
        ))

    @staticmethod
    def _measure_print_titles(sheet) -> Tuple[float, float, float, float]:
        """Return title size and any portion outside the active print area."""
        width_points = 0.0
        height_points = 0.0
        extra_width_points = 0.0
        extra_height_points = 0.0
        try:
            columns = str(sheet.PageSetup.PrintTitleColumns or "").strip()
            rows = str(sheet.PageSetup.PrintTitleRows or "").strip()
        except Exception as exc:
            raise ValueError(
                f"Sheet '{sheet.Name}': cannot read print-title settings: {exc}"
            ) from exc
        if not columns and not rows:
            return 0.0, 0.0, 0.0, 0.0

        try:
            print_area_text = str(sheet.PageSetup.PrintArea or "").strip()
            print_area_range = (
                sheet.Range(print_area_text) if print_area_text else None
            )
        except Exception as exc:
            raise ValueError(
                f"Sheet '{sheet.Name}': cannot resolve PrintArea while "
                f"measuring print titles: {exc}"
            ) from exc

        if columns:
            try:
                title_range = sheet.Range(columns)
                width_points = float(title_range.Width)
                if print_area_range is not None:
                    overlap = sheet.Application.Intersect(
                        print_area_range, title_range
                    )
                    overlap_width = (
                        float(overlap.Width) if overlap is not None else 0.0
                    )
                    extra_width_points = max(
                        0.0, width_points - overlap_width
                    )
            except Exception as exc:
                raise ValueError(
                    f"Sheet '{sheet.Name}': cannot measure "
                    f"PrintTitleColumns {columns!r}: {exc}"
                ) from exc
        if rows:
            try:
                title_range = sheet.Range(rows)
                height_points = float(title_range.Height)
                if print_area_range is not None:
                    overlap = sheet.Application.Intersect(
                        print_area_range, title_range
                    )
                    overlap_height = (
                        float(overlap.Height) if overlap is not None else 0.0
                    )
                    extra_height_points = max(
                        0.0, height_points - overlap_height
                    )
            except Exception as exc:
                raise ValueError(
                    f"Sheet '{sheet.Name}': cannot measure "
                    f"PrintTitleRows {rows!r}: {exc}"
                ) from exc
        return (
            width_points / ExcelConverter.POINTS_PER_INCH,
            height_points / ExcelConverter.POINTS_PER_INCH,
            extra_width_points / ExcelConverter.POINTS_PER_INCH,
            extra_height_points / ExcelConverter.POINTS_PER_INCH,
        )

    def _apply_print_title_override(self, sheet, prop_name: str, value: str) -> None:
        """Apply a title range while accepting Excel's canonicalized A1 syntax."""
        page_setup = sheet.PageSetup
        if not self._safe_set_page_property(page_setup, prop_name, value):
            raise ValueError(f"Excel rejected required PageSetup.{prop_name}={value!r}")
        try:
            page_setup.Application.PrintCommunication = True
        except Exception as exc:
            raise ValueError(f"Cannot verify PageSetup.{prop_name}: {exc}") from exc
        self._verify_print_title_readback(sheet, prop_name, value)

    @staticmethod
    def _verify_print_title_readback(
        sheet, prop_name: str, expected: str
    ) -> None:
        """Verify a title range, resolving Excel's canonical A1 representation."""
        try:
            actual = str(getattr(sheet.PageSetup, prop_name) or "").strip()
        except Exception as exc:
            raise ValueError(f"Cannot verify PageSetup.{prop_name}: {exc}") from exc
        expected = str(expected or "").strip()
        if actual == expected:
            return
        if not actual or not expected:
            raise ValueError(
                f"Excel did not retain PageSetup.{prop_name}: requested "
                f"{expected!r}, read back {actual!r}"
            )
        try:
            expected_address = str(sheet.Range(expected).Address)
            actual_address = str(sheet.Range(actual).Address)
        except Exception as exc:
            raise ValueError(
                f"Cannot resolve PageSetup.{prop_name} readback: {exc}"
            ) from exc
        if actual_address != expected_address:
            raise ValueError(
                f"Excel did not retain PageSetup.{prop_name}: requested "
                f"{expected!r}, read back {actual!r}"
            )

    def _apply_page_setup(
        self, 
        sheet, 
        excel_settings: ExcelSettings,
        filename: str,
        last_col: int,
        content_width_points: Optional[float] = None,
        content_height_points: Optional[float] = None,
        forced_layout: Optional[LayoutCandidate] = None,
    ) -> LayoutCandidate:
        """
        Apply page setup settings for OCR-optimized PDF output.
        
        Args:
            sheet: Excel Worksheet object
            excel_settings: Excel-specific settings
            filename: Original filename for header
            last_col: Last used column index for width calculation
            content_width_points: Optional total content width in points
            content_height_points: Optional total content height in points
        """
        page_setup = sheet.PageSetup
        app = sheet.Application
        app.DisplayAlerts = False
        app.Interactive = False
        try:
            preserved_black_and_white = bool(page_setup.BlackAndWhite)
        except Exception:
            preserved_black_and_white = False
        if excel_settings.print_title_rows is not None:
            self._apply_print_title_override(
                sheet, "PrintTitleRows", excel_settings.print_title_rows
            )
        if excel_settings.print_title_columns is not None:
            self._apply_print_title_override(
                sheet, "PrintTitleColumns", excel_settings.print_title_columns
            )
        try:
            expected_title_rows = str(page_setup.PrintTitleRows or "").strip()
            expected_title_columns = str(
                page_setup.PrintTitleColumns or ""
            ).strip()
        except Exception as exc:
            raise ValueError(
                f"Sheet '{sheet.Name}': cannot read final print-title settings: {exc}"
            ) from exc

        content_width, content_height = self._calculate_smart_page_size(
            sheet, last_col, content_width_points, content_height_points
        )
        orientation_setting = excel_settings.orientation.lower()
        orientations = (
            (xlPortrait, xlLandscape) if orientation_setting == "auto"
            else (xlLandscape,) if orientation_setting == "landscape"
            else (xlPortrait,)
        )
        requested_margins = (
            36.0,
            36.0,
            72.0 if excel_settings.metadata_header else 36.0,
            36.0,
        )
        quality_zoom = max(
            10,
            min(100, int(math.ceil(
                (float(excel_settings.min_shrink_factor) * 100.0) - 1e-9
            ))),
        )
        fit_tall = (
            excel_settings.row_dimensions == 0
            or excel_settings.oversized_action == "error"
        )
        (
            title_width,
            title_height,
            title_extra_width,
            title_extra_height,
        ) = self._measure_print_titles(sheet)
        planned_content_width = content_width + title_extra_width
        planned_content_height = content_height + title_extra_height
        candidates: List[LayoutCandidate] = []
        forms = self._get_printer_paper_forms(app)
        allowed = (
            {name.casefold() for name in excel_settings.allowed_papers}
            if excel_settings.allowed_papers else None
        )
        if allowed is not None:
            forms = tuple(form for form in forms if form.name.casefold() in allowed)
        for form in forms:
            if excel_settings.quality_profile != "legacy" and (
                max(form.width_inches, form.height_inches)
                > excel_settings.max_page_dimension_in
                or form.area > excel_settings.max_page_area_in2
            ):
                continue
            for orientation in orientations:
                actual_margins = self._probe_paper_orientation(
                    page_setup, form, orientation, requested_margins,
                    require_imageable_area=(
                        excel_settings.quality_profile == "strict"
                    ),
                )
                if actual_margins is None:
                    continue
                try:
                    candidate = self._build_layout_candidate(
                        form=form,
                        orientation=orientation,
                        content_width_inches=planned_content_width,
                        content_height_inches=planned_content_height,
                        fit_tall=fit_tall,
                        margins_points=actual_margins,
                        quality_zoom=quality_zoom,
                        title_width_inches=title_width,
                        title_height_inches=title_height,
                    )
                except ValueError:
                    continue
                if excel_settings.quality_profile != "legacy":
                    scaled_width = candidate.usable_width_inches / (quality_zoom / 100.0)
                    scaled_height = candidate.usable_height_inches / (quality_zoom / 100.0)
                    if title_width >= scaled_width or title_height >= scaled_height:
                        continue
                candidates.append(candidate)
        if forced_layout is not None:
            # The form/orientation pair was printer-verified when the sheet-level
            # decision was made. Reuse its geometry for each chunk, then rely on
            # the required setters and final readback below to fail closed if
            # Excel does not retain it on this staged sheet.
            actual_margins = forced_layout.margins_points
            selected = self._build_layout_candidate(
                forced_layout.form, forced_layout.orientation,
                planned_content_width, planned_content_height, fit_tall,
                actual_margins, quality_zoom, title_width, title_height,
            )
            if forced_layout.max_zoom < quality_zoom:
                # The sheet decision intentionally paginates at the quality
                # floor; max_zoom describes one-page fit and is not a cap.
                zoom_value = quality_zoom
            else:
                zoom_value = min(forced_layout.max_zoom, selected.max_zoom)
                if zoom_value < quality_zoom:
                    if excel_settings.oversized_action == "skip":
                        raise OversizedSheetError(
                            "sheet-level layout violates quality floor"
                        )
                    raise ValueError("sheet-level layout violates quality floor")
            use_fit_to_pages = False
            layout_mode = "sheet-quality-layout"
        elif not candidates:
            raise ValueError(f"Sheet '{sheet.Name}': active printer accepted no supported paper form")
        else:
            if excel_settings.quality_profile == "legacy":
                selected = self._select_fit_candidate(candidates, quality_zoom)
            else:
                preferred = {
                    name.casefold(): rank
                    for rank, name in enumerate(excel_settings.preferred_papers)
                }
                eligible = [
                    candidate for candidate in candidates
                    if candidate.max_zoom >= quality_zoom
                ]
                selected = min(eligible, key=lambda candidate: (
                    candidate.pages_wide,
                    candidate.page_count,
                    preferred.get(candidate.form.name.casefold(), 1_000_000),
                    -min(100, candidate.max_zoom),
                    candidate.usable_width_inches * candidate.usable_height_inches,
                    candidate.form.area,
                    candidate.form.paper_enum,
                    candidate.orientation,
                )) if eligible else None
            use_fit_to_pages = False
        if forced_layout is None and selected is not None:
            zoom_value = selected.max_zoom
            layout_mode = "quality-fit"
        elif forced_layout is None:
            best_quality = min(candidates, key=self._fit_candidate_sort_key)
            message = (
                f"Sheet '{sheet.Name}': best required-axis scale "
                f"{best_quality.effective_scale:.3f} on {best_quality.form.name} "
                f"is below min_shrink_factor "
                f"{excel_settings.min_shrink_factor:.3f}"
            )
            if excel_settings.oversized_action == "skip":
                raise OversizedSheetError(message)
            if excel_settings.oversized_action == "warn":
                logger.warning(message)
                selected = best_quality
                zoom_value = False
                use_fit_to_pages = True
                layout_mode = "low-quality-fit"
            elif excel_settings.oversized_action == "paginate":
                if excel_settings.quality_profile == "legacy":
                    selected = self._select_paginated_candidate(candidates)
                else:
                    preferred = {
                        name.casefold(): rank
                        for rank, name in enumerate(excel_settings.preferred_papers)
                    }
                    selected = min(candidates, key=lambda candidate: (
                        candidate.pages_wide,
                        candidate.page_count,
                        preferred.get(candidate.form.name.casefold(), 1_000_000),
                        candidate.form.area,
                        candidate.form.paper_enum,
                        candidate.orientation,
                    ))
                zoom_value = quality_zoom
                layout_mode = "quality-paginate"
            else:
                raise ValueError(message)

        if selected is None:  # Defensive narrowing for type checkers and mocks.
            raise ValueError(f"Sheet '{sheet.Name}': layout selection failed")

        if excel_settings.quality_profile != "legacy" and isinstance(zoom_value, int):
            applied_pages_wide = self._page_span_count(
                planned_content_width, title_width,
                selected.usable_width_inches, zoom_value,
            )
            applied_pages_tall = self._page_span_count(
                planned_content_height, title_height,
                selected.usable_height_inches, zoom_value,
            )
            selected = dataclasses.replace(
                selected,
                pages_wide=applied_pages_wide,
                pages_tall=applied_pages_tall,
                page_count=applied_pages_wide * applied_pages_tall,
            )
            if (
                excel_settings.horizontal_overflow_strategy == "error"
                and applied_pages_wide > 1
            ):
                raise ValueError(
                    f"Sheet '{sheet.Name}' requires horizontal pagination"
                )

        self._required_set_page_property(page_setup, "PaperSize", selected.form.paper_enum)
        self._required_set_page_property(page_setup, "Orientation", selected.orientation)
        for prop_name, value in zip(
            ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"),
            selected.margins_points,
            strict=True,
        ):
            self._required_set_page_property(page_setup, prop_name, value)
        if use_fit_to_pages:
            self._required_set_page_property(page_setup, "Zoom", False)
            self._required_set_page_property(page_setup, "FitToPagesWide", 1)
            self._required_set_page_property(
                page_setup, "FitToPagesTall", 1 if fit_tall else False
            )
        else:
            self._required_set_page_property(page_setup, "FitToPagesWide", False)
            self._required_set_page_property(page_setup, "FitToPagesTall", False)
            self._required_set_page_property(page_setup, "Zoom", zoom_value)
        black_and_white = (
            preserved_black_and_white
            if excel_settings.color_policy == "preserve"
            else excel_settings.color_policy == "black_and_white"
        )
        self._required_set_page_property(page_setup, "BlackAndWhite", black_and_white)
        if excel_settings.quality_profile != "legacy":
            self._required_set_page_property(page_setup, "Draft", excel_settings.draft_mode)
            if selected.pages_wide > 1 and selected.pages_tall > 1:
                self._required_set_page_property(page_setup, "Order", 1)
        app.PrintCommunication = True
        final_expected = {
            "PaperSize": selected.form.paper_enum,
            "Orientation": selected.orientation,
            "LeftMargin": selected.margins_points[0],
            "RightMargin": selected.margins_points[1],
            "TopMargin": selected.margins_points[2],
            "BottomMargin": selected.margins_points[3],
            "Zoom": zoom_value,
            "FitToPagesWide": 1 if use_fit_to_pages else False,
            "FitToPagesTall": (
                (1 if fit_tall else False) if use_fit_to_pages else False
            ),
            "BlackAndWhite": black_and_white,
        }
        if excel_settings.quality_profile != "legacy":
            final_expected["Draft"] = excel_settings.draft_mode
        for prop_name, expected in final_expected.items():
            try:
                actual = getattr(page_setup, prop_name)
            except Exception as exc:
                raise ValueError(
                    f"Cannot verify final PageSetup.{prop_name} for "
                    f"sheet '{sheet.Name}': {exc}"
                ) from exc
            if not self._page_property_matches(actual, expected):
                raise ValueError(
                    f"Excel did not retain final PageSetup.{prop_name} for "
                    f"sheet '{sheet.Name}': expected {expected!r}, "
                    f"read back {actual!r}"
                )
        self._verify_print_title_readback(
            sheet, "PrintTitleRows", expected_title_rows
        )
        self._verify_print_title_readback(
            sheet, "PrintTitleColumns", expected_title_columns
        )
        logger.info(
            f"Sheet '{sheet.Name}': {layout_mode}, {selected.form.name} "
            f"{'landscape' if selected.orientation == xlLandscape else 'portrait'}, "
            f"content={planned_content_width:.2f}x"
            f"{planned_content_height:.2f}in, "
            f"usable={selected.usable_width_inches:.2f}x"
            f"{selected.usable_height_inches:.2f}in, "
            f"width_scale={selected.width_scale:.3f}, "
            f"height_scale={selected.height_scale:.3f}, "
            f"effective_scale={selected.effective_scale:.3f}, "
            f"limiting_axis={selected.limiting_axis}, zoom={zoom_value}, "
            f"estimated_pages={selected.pages_wide}x{selected.pages_tall}"
        )
        return selected

    def _apply_metadata_header(
        self, 
        sheet, 
        excel_settings: ExcelSettings,
        filename: str,
        center_text: str = "",
        left_text: str = None
    ) -> None:
        """
        Set header text: sheet name | row range | filename
        """
        try:
            page_setup = sheet.PageSetup
            
            # Build header values
            left_val = left_text if left_text else "&A"
            center_val = center_text
            right_val = f"{filename} (Page &P)"
            
            # Set headers directly (avoid wrapper that may silently fail)
            try:
                page_setup.LeftHeader = left_val
                logger.debug(f"Set LeftHeader = '{left_val}'")
            except Exception as e:
                logger.warning(f"Failed to set LeftHeader: {e}")
            
            try:
                page_setup.CenterHeader = center_val
                logger.debug(f"Set CenterHeader = '{center_val}'")
            except Exception as e:
                logger.warning(f"Failed to set CenterHeader: {e}")
            
            try:
                page_setup.RightHeader = right_val
                logger.debug(f"Set RightHeader = '{right_val}'")
            except Exception as e:
                logger.warning(f"Failed to set RightHeader: {e}")
            
            # Clear footers to avoid clutter and potential crop issues
            try:
                page_setup.RightFooter = ""
                page_setup.CenterFooter = ""
                page_setup.LeftFooter = ""
            except Exception as e:
                logger.debug(f"Failed to clear footers: {e}")
            
            # CRITICAL: Re-enable PrintCommunication to commit header/footer changes
            try:
                app = sheet.Application
                app.PrintCommunication = True
            except:
                pass
            
            logger.debug(f"Applied metadata header for sheet '{sheet.Name}' (Center: '{center_text}')")
            
        except Exception as e:
            logger.warning(f"Could not apply metadata header for '{sheet.Name}': {e}")

    def _insert_sheet_name_label(self, sheet, sheet_name: str) -> None:
        """
        Insert a new row at the beginning and add sheet name with font size 23.
        
        This feature adds the sheet name as a large, bold label in the first row
        to improve OCR recognition of the sheet name.
        
        Args:
            sheet: Excel Worksheet object
            sheet_name: Name of the sheet to insert as label
        """
        try:
            # Insert new row at position 1
            sheet.Rows(1).Insert()
            
            # Set sheet name in cell A1
            cell = sheet.Cells(1, 1)
            cell.Value = sheet_name
            
            # Set font size to 23 for OCR readability
            cell.Font.Size = 23
            cell.Font.Bold = True
            
            logger.debug(f"Inserted OCR sheet name label for '{sheet_name}'")
        except Exception as e:
            logger.warning(f"Could not insert OCR sheet name label for '{sheet_name}': {e}")

    def _insert_file_path_row(self, sheet, file_path: Path, last_row: int, last_col: int, base_path: Optional[Path] = None) -> int:
        """
        Insert a new row before the last row and add the file path centered.
        
        Args:
            sheet: Excel Worksheet object
            file_path: Absolute path of the file being converted
            last_row: The last row index with content
            last_col: The last column index with content
            base_path: Optional root directory to calculate relative path
            
        Returns:
            The updated last_row after insertion
        """
        try:
            if last_row < 2:
                # Sheet too small, insert at row 2
                insert_row = 2
            else:
                # Insert before last row
                insert_row = last_row
            
            # Insert new row
            sheet.Rows(insert_row).Insert()
            
            # Calculate center column
            center_col = max(1, (last_col + 1) // 2)
            
            # Calculate path to display
            display_path = ""
            if base_path:
                try:
                    rel_path = file_path.resolve().relative_to(base_path.resolve())
                    display_path = "/" + rel_path.as_posix()
                except ValueError:
                    display_path = str(file_path.resolve())
            else:
                display_path = str(file_path.resolve())
            
            # Set file path in center cell
            cell = sheet.Cells(insert_row, center_col)
            cell.Value = display_path
            
            # Format: Italic, slightly smaller font
            cell.Font.Italic = True
            cell.Font.Size = 10
            cell.HorizontalAlignment = -4108  # xlCenter
            
            logger.debug(f"Inserted file path '{display_path}' at row {insert_row} for '{sheet.Name}'")
            
            return last_row + 1  # Return updated last_row
            
        except Exception as e:
            logger.warning(f"Could not insert file path row for '{sheet.Name}': {e}")
            return last_row

    def _col_num_to_letter(self, n: int) -> str:
        """Convert 1-based column number to Excel column letter (e.g. 1->A, 27->AA)."""
        string = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string

    def _expand_bounds_for_shapes(
        self, 
        sheet, 
        max_width: float, 
        max_height: float, 
        last_row: int, 
        last_col: int,
        points_per_inch: float
    ) -> Tuple[float, float, int, int]:
        """
        Safely iterate through shapes to expand content bounds.
        
        Uses per-shape timeout to prevent COM blocking from problematic shapes
        (OLE objects, external links, etc.) from freezing the application.
        
        Args:
            sheet: Excel Worksheet object
            max_width: Current max width in points
            max_height: Current max height in points  
            last_row: Current last row index
            last_col: Current last column index
            points_per_inch: Conversion factor
            
        Returns:
            Tuple of (max_width, max_height, last_row, last_col)
        """
        MAX_SHAPE_ERRORS = 5  # Stop after this many consecutive errors
        
        try:
            # First, try to get shapes count with timeout
            shapes_count = 0
            try:
                shapes_count = sheet.Shapes.Count
            except Exception as e:
                logger.debug(f"Could not access Shapes collection: {e}")
                return max_width, max_height, last_row, last_col
            
            if shapes_count == 0:
                return max_width, max_height, last_row, last_col
                
            logger.debug(f"Processing {shapes_count} shapes for bounds calculation...")
            consecutive_errors = 0
            
            for i in range(1, shapes_count + 1):  # Excel shapes are 1-indexed
                try:
                    shape = sheet.Shapes(i)
                    
                    # Access shape properties with individual try-except
                    # This prevents one bad shape from blocking the entire loop
                    shape_name = "Unknown"
                    try:
                        shape_name = shape.Name
                    except:
                        pass
                    
                    # Get position/size properties - these can block on OLE objects
                    shape_left = 0
                    shape_top = 0
                    shape_width = 0
                    shape_height = 0
                    
                    try:
                        shape_left = shape.Left
                        shape_top = shape.Top
                        shape_width = shape.Width
                        shape_height = shape.Height
                    except Exception as prop_err:
                        logger.debug(f"Shape {i} '{shape_name}' property access failed: {prop_err}")
                        consecutive_errors += 1
                        if consecutive_errors >= MAX_SHAPE_ERRORS:
                            logger.warning(f"Too many shape access errors ({MAX_SHAPE_ERRORS}), skipping remaining shapes")
                            break
                        continue
                    
                    # Reset error counter on success
                    consecutive_errors = 0
                    
                    shape_right = shape_left + shape_width
                    shape_bottom = shape_top + shape_height
                    
                    if shape_right > max_width:
                        logger.debug(f"Shape '{shape_name}' extends width to {shape_right:.1f}pt ({shape_right/points_per_inch:.2f}in)")
                        max_width = shape_right
                    if shape_bottom > max_height:
                        max_height = shape_bottom
                    
                    # Try to get cell bounds (optional, non-critical)
                    try:
                        br_cell = shape.BottomRightCell
                        if br_cell:
                            if br_cell.Row > last_row:
                                last_row = br_cell.Row
                            if br_cell.Column > last_col:
                                last_col = br_cell.Column
                    except Exception:
                        pass
                        
                except Exception as shape_err:
                    logger.debug(f"Error processing shape {i}: {shape_err}")
                    consecutive_errors += 1
                    if consecutive_errors >= MAX_SHAPE_ERRORS:
                        logger.warning(f"Too many shape access errors ({MAX_SHAPE_ERRORS}), skipping remaining shapes")
                        break
                    continue
                    
        except Exception as e:
            logger.warning(f"Shape bounds expansion failed: {e}")
            
        return max_width, max_height, last_row, last_col

    def _export_to_pdf(
        self, 
        workbook, 
        sheets: List,
        output_path: str,
        settings: PDFConversionSettings
    ) -> None:
        """
        Export sheets to PDF.
        
        Args:
            workbook: Excel Workbook object
            sheets: List of sheets to export
            output_path: Path for output PDF
            settings: PDF conversion settings
        """
        try:
            # Validate COM connection before starting
            try:
                app = workbook.Application
                _ = app.Version  # Quick validation
            except Exception as e:
                raise COMDisconnectedError(
                    f"Excel connection lost before export: {e}"
                ) from e
            
            # Ensure dialogs are suppressed before export
            app.DisplayAlerts = False
            app.Interactive = False
            
            # CRITICAL: Re-enable PrintCommunication before export
            # When False, PageSetup changes (headers/footers) are NOT communicated to printer
            # Must be True for headers/footers to appear in PDF
            try:
                app.PrintCommunication = True
            except:
                pass
            
            # Determine quality
            quality = xlQualityStandard
            if settings.optimization.image_quality == "low":
                quality = xlQualityMinimum

            if len(sheets) == 1:
                # Export single sheet directly
                logger.info(f"Exporting sheet '{sheets[0].Name}' to PDF...")
                
                sheets[0].ExportAsFixedFormat(
                    Type=xlTypePDF,
                    Filename=output_path,
                    Quality=quality,
                    IncludeDocProperties=settings.metadata.include_properties,
                    IgnorePrintAreas=False,
                    OpenAfterPublish=False
                )
                
                logger.debug(f"Sheet '{sheets[0].Name}' exported successfully")
            else:
                # Multiple sheets: Copy to new temporary workbook iteratively
                logger.debug(f"Preparing to copy {len(sheets)} sheets to new workbook.")
                
                temp_wb = None
                try:
                    # Copy first sheet -> Creates new Workbook
                    sheets[0].Copy()
                    
                    # Get the new workbook
                    try:
                        temp_wb = workbook.Application.ActiveWorkbook
                        _ = temp_wb.Sheets.Count  # Validate connection
                    except Exception as e:
                        raise COMDisconnectedError(
                            f"Failed to access temp workbook: {e}"
                        ) from e
                    
                    logger.debug(f"Created temp WB. Sheets count: {temp_wb.Sheets.Count}")
                    
                    # Copy remaining sheets into the new workbook
                    for idx, s in enumerate(sheets[1:], start=2):
                        try:
                            last_sheet = temp_wb.Sheets(temp_wb.Sheets.Count)
                            # Copy after last_sheet
                            s.Copy(None, last_sheet)
                            logger.debug(
                                f"Copied sheet {idx}/{len(sheets)}. "
                                f"New count: {temp_wb.Sheets.Count}"
                            )
                        except Exception as copy_err:
                            logger.error(f"Failed to copy sheet {idx}: {copy_err}")
                            raise ValueError(
                                f"Failed to copy sheet {idx}/{len(sheets)} "
                                "into the temporary export workbook"
                            ) from copy_err
                    
                    # Export workbook - all sheets will be included automatically
                    count = temp_wb.Sheets.Count
                    logger.debug(f"Exporting created workbook with {count} sheets.")
                    
                    logger.info(f"Exporting {count} sheets to PDF...")
                    
                    temp_wb.ExportAsFixedFormat(
                        Type=xlTypePDF,
                        Filename=output_path,
                        Quality=quality,
                        IncludeDocProperties=settings.metadata.include_properties,
                        IgnorePrintAreas=False,
                        OpenAfterPublish=False
                    )
                    
                    logger.debug("Multi-sheet export completed successfully")
                finally:
                    if temp_wb:
                        try:
                            temp_wb.Close(SaveChanges=False)
                        except Exception as close_err:
                            logger.debug(f"Failed to close temp workbook: {close_err}")

            
        except COMDisconnectedError:
            raise  # Re-raise to caller
        except Exception as e:
            logger.error(f"Failed to export to PDF: {e}")
            raise

    def _copy_region_sheet(self, workbook, source_sheet, region: SheetRegion):
        """Copy a worksheet and assign one verified, independent print region."""
        last_sheet = workbook.Sheets(workbook.Sheets.Count)
        source_sheet.Copy(None, last_sheet)
        copied = workbook.Sheets(workbook.Sheets.Count)
        first_col = self._col_num_to_letter(region.first_col)
        last_col = self._col_num_to_letter(region.last_col)
        print_area = (
            f"${first_col}${region.first_row}:${last_col}${region.last_row}"
        )
        self._required_set_page_property(copied.PageSetup, "PrintArea", print_area)
        return copied

    def _resolve_sheet_regions(self, sheet, policy: str) -> List[SheetRegion]:
        """Resolve preserved Range.Areas or automatic cell/shape content bounds."""
        if policy == "preserve":
            try:
                print_area = str(sheet.PageSetup.PrintArea or "").strip()
                if print_area:
                    areas = sheet.Range(print_area).Areas
                    regions = []
                    for index in range(1, int(areas.Count) + 1):
                        area = areas(index)
                        regions.append(SheetRegion(
                            int(area.Row), int(area.Column),
                            int(area.Row + area.Rows.Count - 1),
                            int(area.Column + area.Columns.Count - 1),
                        ))
                    if regions:
                        return regions
            except Exception as exc:
                logger.warning(
                    f"Sheet '{sheet.Name}': invalid PrintArea ignored: {exc}"
                )

        first_row = first_col = None
        last_row = last_col = None
        # Search formulas as well as displayed values. Page breaks are deliberately
        # excluded: they describe pagination, not visible content.
        for look_in in (-4123, -4163):  # xlFormulas, xlValues
            try:
                first_r = sheet.Cells.Find(
                    What="*", After=sheet.Range("A1"), LookIn=look_in, LookAt=2,
                    SearchOrder=self.xlByRows, SearchDirection=1,
                )
                last_r = sheet.Cells.Find(
                    What="*", After=sheet.Range("A1"), LookIn=look_in, LookAt=2,
                    SearchOrder=self.xlByRows, SearchDirection=self.xlPrevious,
                )
                first_c = sheet.Cells.Find(
                    What="*", After=sheet.Range("A1"), LookIn=look_in, LookAt=2,
                    SearchOrder=self.xlByColumns, SearchDirection=1,
                )
                last_c = sheet.Cells.Find(
                    What="*", After=sheet.Range("A1"), LookIn=look_in, LookAt=2,
                    SearchOrder=self.xlByColumns, SearchDirection=self.xlPrevious,
                )
                if all((first_r, last_r, first_c, last_c)):
                    first_row = min(first_row or first_r.Row, int(first_r.Row))
                    last_row = max(last_row or last_r.Row, int(last_r.Row))
                    first_col = min(first_col or first_c.Column, int(first_c.Column))
                    last_col = max(last_col or last_c.Column, int(last_c.Column))
            except Exception:
                continue

        # Include visible shapes by their anchor cells. Hidden shapes do not render.
        try:
            for index in range(1, int(sheet.Shapes.Count) + 1):
                shape = sheet.Shapes(index)
                if hasattr(shape, "Visible") and not bool(shape.Visible):
                    continue
                top_left = shape.TopLeftCell
                bottom_right = shape.BottomRightCell
                first_row = min(first_row or top_left.Row, int(top_left.Row))
                first_col = min(first_col or top_left.Column, int(top_left.Column))
                last_row = max(last_row or bottom_right.Row, int(bottom_right.Row))
                last_col = max(last_col or bottom_right.Column, int(bottom_right.Column))
        except Exception:
            pass
        if None in (first_row, first_col, last_row, last_col):
            return []
        return [SheetRegion(first_row, first_col, last_row, last_col)]

    def _get_print_area_bounds(self, sheet) -> Tuple[int, int]:
        """
        Get bounds from existing PrintArea if set by user.
        
        This respects user-defined print area settings which have highest priority.
        
        Returns:
            Tuple of (last_row, last_col) from PrintArea, or (0, 0) if not set.
        """
        try:
            print_area = sheet.PageSetup.PrintArea
            if print_area and print_area.strip():
                # PrintArea format: "$A$1:$Z$100" or "A1:Z100"
                # Parse the end cell to get bounds
                import re
                # Remove sheet name prefix if present (e.g., "Sheet1!$A$1:$Z$100")
                if '!' in print_area:
                    print_area = print_area.split('!')[-1]
                
                # Match pattern like $A$1:$Z$100 or A1:Z100
                match = re.search(r':?\$?([A-Z]+)\$?(\d+)$', print_area.upper())
                if match:
                    col_letters = match.group(1)
                    row_num = int(match.group(2))
                    
                    # Convert column letters to number (A=1, Z=26, AA=27, etc.)
                    col_num = 0
                    for char in col_letters:
                        col_num = col_num * 26 + (ord(char) - ord('A') + 1)
                    
                    logger.debug(f"Sheet '{sheet.Name}' has PrintArea set: {print_area} -> Row={row_num}, Col={col_num}")
                    return row_num, col_num
        except Exception as e:
            logger.debug(f"Could not parse PrintArea: {e}")
        
        return 0, 0
    
    def _get_page_break_bounds(self, sheet) -> Tuple[int, int]:
        """
        Get bounds from vertical/horizontal page breaks if set.
        
        This uses the rightmost vertical page break as the column bound.
        
        Returns:
            Tuple of (last_row, last_col) from page breaks, or (0, 0) if none.
        """
        last_row = 0
        last_col = 0
        
        try:
            # Check VPageBreaks (vertical page breaks define column boundaries)
            v_breaks = sheet.VPageBreaks
            if v_breaks and v_breaks.Count > 0:
                # Get the rightmost break location
                max_break_col = 0
                for i in range(1, v_breaks.Count + 1):
                    try:
                        break_loc = v_breaks(i).Location
                        if break_loc and break_loc.Column > max_break_col:
                            max_break_col = break_loc.Column
                    except Exception:
                        continue
                if max_break_col > 0:
                    last_col = max_break_col - 1  # Break is BEFORE this column
                    logger.debug(f"Sheet '{sheet.Name}' VPageBreak found at column {max_break_col}")
        except Exception as e:
            logger.debug(f"Could not read VPageBreaks: {e}")
        
        try:
            # Check HPageBreaks (horizontal page breaks define row boundaries)
            h_breaks = sheet.HPageBreaks
            if h_breaks and h_breaks.Count > 0:
                max_break_row = 0
                for i in range(1, h_breaks.Count + 1):
                    try:
                        break_loc = h_breaks(i).Location
                        if break_loc and break_loc.Row > max_break_row:
                            max_break_row = break_loc.Row
                    except Exception:
                        continue
                if max_break_row > 0:
                    last_row = max_break_row - 1  # Break is BEFORE this row
                    logger.debug(f"Sheet '{sheet.Name}' HPageBreak found at row {max_break_row}")
        except Exception as e:
            logger.debug(f"Could not read HPageBreaks: {e}")
        
        return last_row, last_col

    def _find_longest_text_column(self, sheet, search_last_row: int, search_last_col: int) -> Tuple[int, int, float]:
        """
        Find text that extends beyond column width using row sampling.
        
        Handles merged cells by calculating the total width of the merge area.
        Samples first N, last N, and middle rows for better coverage.
        
        Returns:
            Tuple of (extended_col, max_text_length, required_extra_width_points)
        """
        max_text_extended_col = 0
        max_text_len = 0
        required_extra_width = 0.0
        
        AVG_CHAR_WIDTH_POINTS = 7.2
        DEFAULT_COL_WIDTH = 64.0
        SAMPLE_ROWS = 50
        
        try:
            max_cols = search_last_col + 20 
            
            # Cache column widths
            col_widths = []
            for col_idx in range(1, max_cols + 1):
                try:
                    col_widths.append(sheet.Columns(col_idx).Width)
                except Exception:
                    col_widths.append(DEFAULT_COL_WIDTH)
            
            # Select rows to check
            rows_to_check = set()
            for r in range(1, min(SAMPLE_ROWS + 1, search_last_row + 1)):
                rows_to_check.add(r)
            for r in range(max(1, search_last_row - SAMPLE_ROWS + 1), search_last_row + 1):
                rows_to_check.add(r)
            if search_last_row > SAMPLE_ROWS * 3:
                mid = search_last_row // 2
                for r in range(max(1, mid - 5), min(search_last_row, mid + 5)):
                    rows_to_check.add(r)
            
            check_list = sorted(list(rows_to_check))
            
            for row_idx in check_list:
                try:
                    row_range = sheet.Range(
                        sheet.Cells(row_idx, 1),
                        sheet.Cells(row_idx, min(max_cols, search_last_col + 10))
                    )
                    row_values = row_range.Value
                    
                    if row_values is None:
                        continue

                    if isinstance(row_values, tuple):
                        if isinstance(row_values[0], tuple):
                            row_values = row_values[0]
                    else:
                        row_values = (row_values,)
                    
                    for col_idx, value in enumerate(row_values, start=1):
                        if value is None or not isinstance(value, (str, float, int)):
                            continue
                        
                        text = str(value)
                        text_len = len(text)
                        
                        if text_len > 15:
                            # Check if this cell has wrap text enabled - if so, skip overflow detection
                            try:
                                cell = sheet.Cells(row_idx, col_idx)
                                if cell.WrapText:
                                    # Text wraps within the column, no horizontal overflow
                                    continue
                            except Exception:
                                pass
                            
                            estimated_width = text_len * AVG_CHAR_WIDTH_POINTS
                            
                            # Check if this cell is merged and calculate merged width
                            try:
                                cell = sheet.Cells(row_idx, col_idx)
                                merge_area = cell.MergeArea
                                if merge_area.Columns.Count > 1:
                                    # Sum widths of all merged columns
                                    base_width = 0.0
                                    merge_start_col = merge_area.Column
                                    merge_end_col = merge_start_col + merge_area.Columns.Count - 1
                                    for mc in range(merge_start_col, merge_end_col + 1):
                                        if mc <= len(col_widths):
                                            base_width += col_widths[mc - 1]
                                        else:
                                            base_width += DEFAULT_COL_WIDTH
                                    # The extended column should start after the merge area
                                    effective_col = merge_end_col
                                else:
                                    base_width = col_widths[col_idx - 1] if col_idx <= len(col_widths) else DEFAULT_COL_WIDTH
                                    effective_col = col_idx
                            except Exception:
                                base_width = col_widths[col_idx - 1] if col_idx <= len(col_widths) else DEFAULT_COL_WIDTH
                                effective_col = col_idx
                            
                            if estimated_width > base_width:
                                overflow = estimated_width - base_width
                                
                                extended_col = effective_col
                                accumulated = 0.0
                                for nc in range(effective_col, len(col_widths)):
                                    accumulated += col_widths[nc]
                                    extended_col = nc + 1
                                    if accumulated >= overflow:
                                        break
                                
                                if extended_col > max_text_extended_col:
                                    max_text_extended_col = extended_col
                                    max_text_len = text_len
                                    required_extra_width = overflow
                                    
                except Exception:
                    continue
            
            if max_text_len > 0:
                # Log column widths for debugging
                col_width_summary = ", ".join([f"Col{i+1}:{col_widths[i]:.1f}pt" for i in range(min(search_last_col, len(col_widths)))])
                total_width_pts = sum(col_widths[:search_last_col])
                logger.debug(
                    f"Sheet '{sheet.Name}' text overflow detected: {max_text_len} chars extending to col {max_text_extended_col}. "
                    f"Column widths (1-{search_last_col}): [{col_width_summary}], Total: {total_width_pts:.1f}pt ({total_width_pts/72:.2f}in)"
                )
                        
        except Exception as e:
            logger.debug(f"Text overflow detection sampling failed: {e}")
        
        return max_text_extended_col, max_text_len, required_extra_width


    def _get_content_dimensions_points(self, sheet) -> Tuple[float, float, int, int]:
        """
        Calculate total content width and height in points by summing column widths.
        
        Priority order for determining bounds:
        1. PrintArea (if set by user) - highest priority
        2. Page breaks (VPageBreaks/HPageBreaks)
        3. Cells.Find + longest text detection (fallback)
        
        Returns (max_width_points, max_height_points, last_row, last_col).
        """
        max_width = 0.0
        max_height = 0.0
        
        POINTS_PER_INCH = 72.0
        
        try:
            last_row = 1
            last_col = 1
            bounds_source = "default"
            
            # Priority 1: Check for PrintArea
            print_row, print_col = self._get_print_area_bounds(sheet)
            if print_row > 0 and print_col > 0:
                last_row = print_row
                last_col = print_col
                bounds_source = "PrintArea"
                logger.info(f"Sheet '{sheet.Name}' using PrintArea bounds: Row={last_row}, Col={last_col}")
            else:
                # Priority 2: Check for page breaks
                break_row, break_col = self._get_page_break_bounds(sheet)
                if break_row > 0 or break_col > 0:
                    bounds_source = "PageBreaks"
                
                # Priority 3: Use Cells.Find for base detection
                try:
                    last_row_cell = sheet.Cells.Find(
                        What="*",
                        After=sheet.Range("A1"),
                        LookIn=-4163,  # xlValues
                        LookAt=2,      # xlPart
                        SearchOrder=self.xlByRows,
                        SearchDirection=self.xlPrevious
                    )
                    if last_row_cell:
                        last_row = last_row_cell.Row
                except Exception:
                    last_row = sheet.UsedRange.Rows.Count
                
                try:
                    last_col_cell = sheet.Cells.Find(
                        What="*",
                        After=sheet.Range("A1"),
                        LookIn=-4163,  # xlValues
                        LookAt=2,      # xlPart
                        SearchOrder=self.xlByColumns,
                        SearchDirection=self.xlPrevious
                    )
                    if last_col_cell:
                        last_col = last_col_cell.Column
                except Exception:
                    last_col = sheet.UsedRange.Columns.Count
                
                # Apply page break bounds if they are larger
                if break_row > last_row:
                    last_row = break_row
                if break_col > last_col:
                    last_col = break_col
                    bounds_source = "PageBreaks"
                
                # Priority 3b: Detect longest text and extend bounds if needed
                # Skip text overflow detection if VPageBreak defines column boundary
                overflow_extra_width = 0.0
                if break_col > 0:
                    logger.debug(f"Sheet '{sheet.Name}' skipping text overflow detection - VPageBreak defines column boundary at {break_col}")
                else:
                    text_col, text_len, overflow_extra_width = self._find_longest_text_column(sheet, last_row, last_col)
                    if text_col > last_col:
                        logger.info(f"Sheet '{sheet.Name}' extending column bound from {last_col} to {text_col} for text overflow")
                        last_col = text_col
                        bounds_source = "TextOverflow"
            
            logger.debug(f"Sheet '{sheet.Name}' bounds source: {bounds_source}")
            
            # Sum width of each column (in points)
            total_width_points = 0.0
            
            for col_idx in range(1, last_col + 1):
                try:
                    col_width = sheet.Columns(col_idx).Width
                    total_width_points += col_width
                except Exception:
                    total_width_points += 64.0  # Default column width
            
            # Add extra width for text overflow if detected
            if bounds_source != "PrintArea" and overflow_extra_width > 0:
                total_width_points += overflow_extra_width
                logger.debug(f"Added {overflow_extra_width:.1f}pt for text overflow")
            
            # Sum height of each row (in points)
            total_height_points = 0.0
            
            for row_idx in range(1, last_row + 1):
                try:
                    row_height = sheet.Rows(row_idx).Height
                    total_height_points += row_height
                except Exception:
                    total_height_points += 15.0  # Default row height
            
            max_width = total_width_points
            max_height = total_height_points
            
            logger.debug(
                f"Sheet '{sheet.Name}' Column Sum: "
                f"Cols=1-{last_col}, Total Width={total_width_points:.1f}pt ({total_width_points/POINTS_PER_INCH:.2f}in) | "
                f"Rows=1-{last_row}, Total Height={total_height_points:.1f}pt ({total_height_points/POINTS_PER_INCH:.2f}in)"
            )
            
            # Expand for Shapes (Charts, Images) with safe iteration
            max_width, max_height, last_row, last_col = self._expand_bounds_for_shapes(
                sheet, max_width, max_height, last_row, last_col, POINTS_PER_INCH
            )
            
            logger.info(
                f"Sheet '{sheet.Name}' Final Content Dimensions: "
                f"{max_width:.1f}pt ({max_width/POINTS_PER_INCH:.2f}in) x {max_height:.1f}pt ({max_height/POINTS_PER_INCH:.2f}in)"
            )
                    
        except Exception as e:
            logger.warning(f"Failed to calculate geometry dimensions: {e}")
            
        return max_width, max_height, last_row, last_col
