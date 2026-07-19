"""Authored PageSetup inspection and fingerprint verification."""

from __future__ import annotations

import zipfile
import posixpath
from pathlib import Path
from typing import Any, Set, Tuple
from xml.etree import ElementTree as ET

from .models import AuthoredLayoutSnapshot


_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def _has_persisted_layout(root: ET.Element) -> bool:
    """Document this Excel pipeline operation and its side effects."""
    setup = root.find(f"{{{_MAIN}}}pageSetup")
    if setup is not None and any(
        name in setup.attrib for name in (
            "paperSize", "orientation", "scale", "fitToWidth", "fitToHeight",
            "pageOrder", "blackAndWhite", "draft", "firstPageNumber",
        )
    ):
        return True
    options = root.find(f"{{{_MAIN}}}printOptions")
    if options is not None and bool(options.attrib):
        return True
    header_footer = root.find(f"{{{_MAIN}}}headerFooter")
    if header_footer is not None and (
        bool(header_footer.attrib)
        or any((child.text or "").strip() for child in header_footer)
    ):
        return True
    for name in ("rowBreaks", "colBreaks"):
        breaks = root.find(f"{{{_MAIN}}}{name}")
        if breaks is not None and len(breaks):
            return True
    margins = root.find(f"{{{_MAIN}}}pageMargins")
    if margins is not None:
        defaults = {
            "left": 0.7, "right": 0.7, "top": 0.75, "bottom": 0.75,
            "header": 0.3, "footer": 0.3,
        }
        for name, default in defaults.items():
            try:
                if abs(float(margins.attrib[name]) - default) > 1e-6:
                    return True
            except (KeyError, TypeError, ValueError):
                continue
    return False


def _safe_value(owner: Any, name: str, errors: list[str], default: Any = None) -> Any:
    """Document this Excel pipeline operation and its side effects."""
    try:
        return getattr(owner, name)
    except Exception as exc:
        errors.append(f"{name}: {exc}")
        return default


def _break_ids(collection: Any, errors: list[str], label: str) -> Tuple[int, ...]:
    """Document this Excel pipeline operation and its side effects."""
    try:
        values = []
        for index in range(1, int(collection.Count) + 1):
            item = collection.Item(index)
            location = getattr(item, "Location", None)
            values.append(int(getattr(location, "Row", getattr(location, "Column", 0))))
        return tuple(values)
    except Exception as exc:
        errors.append(f"{label}: {exc}")
        return ()


def persisted_print_sheets(path: Path) -> Set[str]:
    """Return OOXML sheet names that contain persisted print-layout metadata."""
    if path.suffix.lower() not in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        return set()
    try:
        with zipfile.ZipFile(path) as package:
            workbook = ET.fromstring(package.read("xl/workbook.xml"))
            relationships = ET.fromstring(
                package.read("xl/_rels/workbook.xml.rels")
            )
            targets = {
                relation.attrib["Id"]: relation.attrib["Target"]
                for relation in relationships.findall(f"{{{_PKG_REL}}}Relationship")
            }
            authored: Set[str] = set()
            for sheet in workbook.findall(f".//{{{_MAIN}}}sheet"):
                name = sheet.attrib.get("name", "")
                rel_id = sheet.attrib.get(f"{{{_DOC_REL}}}id", "")
                target = targets.get(rel_id, "")
                if not target:
                    continue
                normalized = posixpath.normpath(target.lstrip("/"))
                if not normalized.startswith("xl/"):
                    normalized = "xl/" + normalized
                root = ET.fromstring(package.read(normalized))
                if _has_persisted_layout(root):
                    authored.add(name)
            defined = workbook.find(f"{{{_MAIN}}}definedNames")
            if defined is not None:
                for item in defined:
                    name = item.attrib.get("name", "")
                    if name not in {"_xlnm.Print_Area", "_xlnm.Print_Titles"}:
                        continue
                    local_id = item.attrib.get("localSheetId")
                    if local_id is not None:
                        sheets = workbook.findall(f".//{{{_MAIN}}}sheet")
                        try:
                            authored.add(sheets[int(local_id)].attrib["name"])
                        except (IndexError, ValueError, KeyError):
                            pass
            return authored
    except (OSError, KeyError, zipfile.BadZipFile, ET.ParseError):
        return set()


class AuthoredLayoutInspector:
    """Classify and snapshot a sheet without changing PageSetup."""

    def __init__(self, workbook_path: Path):
        """Document this Excel pipeline operation and its side effects."""
        self.workbook_path = Path(workbook_path)
        self._persisted = persisted_print_sheets(self.workbook_path)
        self._ooxml = self.workbook_path.suffix.lower() in {
            ".xlsx", ".xlsm", ".xltx", ".xltm"
        }

    def inspect(self, sheet: Any) -> AuthoredLayoutSnapshot:
        """Document this Excel pipeline operation and its side effects."""
        errors: list[str] = []
        setup = _safe_value(sheet, "PageSetup", errors)
        if setup is None:
            return AuthoredLayoutSnapshot(
                "invalid", "uncertain", "com", errors=tuple(errors)
            )
        print_area = str(_safe_value(setup, "PrintArea", errors, "") or "")
        title_rows = str(_safe_value(setup, "PrintTitleRows", errors, "") or "")
        title_columns = str(_safe_value(setup, "PrintTitleColumns", errors, "") or "")
        headers = tuple(str(_safe_value(setup, key, errors, "") or "") for key in (
            "LeftHeader", "CenterHeader", "RightHeader"
        ))
        footers = tuple(str(_safe_value(setup, key, errors, "") or "") for key in (
            "LeftFooter", "CenterFooter", "RightFooter"
        ))
        row_breaks = _break_ids(
            _safe_value(sheet, "HPageBreaks", errors), errors, "HPageBreaks"
        ) if hasattr(sheet, "HPageBreaks") else ()
        col_breaks = _break_ids(
            _safe_value(sheet, "VPageBreaks", errors), errors, "VPageBreaks"
        ) if hasattr(sheet, "VPageBreaks") else ()
        name = str(_safe_value(sheet, "Name", errors, "") or "")
        positive_signal = bool(
            print_area or title_rows or title_columns or any(headers) or any(footers)
            or row_breaks or col_breaks
        )
        if self._ooxml:
            authored = name in self._persisted
            source = "ooxml"
            confidence = "certain"
        else:
            authored = positive_signal
            source = "com-signals"
            confidence = "conservative"
        classification = "authored" if authored else "missing"
        if print_area:
            try:
                sheet.Range(print_area)
            except Exception as exc:
                errors.append(f"PrintArea: {exc}")
                classification = "invalid"
        if errors and classification == "authored":
            confidence = "uncertain"
        return AuthoredLayoutSnapshot(
            classification=classification,
            confidence=confidence,
            source=source,
            print_area=print_area,
            paper_size=_safe_value(setup, "PaperSize", errors),
            orientation=_safe_value(setup, "Orientation", errors),
            margins_points=tuple(float(_safe_value(setup, key, errors, 0.0) or 0.0) for key in (
                "LeftMargin", "RightMargin", "TopMargin", "BottomMargin",
                "HeaderMargin", "FooterMargin",
            )),
            zoom=_safe_value(setup, "Zoom", errors),
            fit_to_pages_wide=_safe_value(setup, "FitToPagesWide", errors),
            fit_to_pages_tall=_safe_value(setup, "FitToPagesTall", errors),
            print_title_rows=title_rows,
            print_title_columns=title_columns,
            page_order=_safe_value(setup, "Order", errors),
            manual_row_breaks=row_breaks,
            manual_column_breaks=col_breaks,
            headers=headers,
            footers=footers,
            black_and_white=_safe_value(setup, "BlackAndWhite", errors),
            draft=_safe_value(setup, "Draft", errors),
            errors=tuple(errors),
        )
