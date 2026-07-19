"""Printer identity and hard-imageable-area discovery."""

from __future__ import annotations

import copy
from typing import Any, Dict, Optional, Tuple

import win32gui
import win32print
import win32ui

from .models import PrinterCapability, PrinterFormCapability


DC_PAPERS = 2
DC_PAPERSIZE = 3
DC_PAPERNAMES = 16
HORZRES = 8
VERTRES = 10
LOGPIXELSX = 88
LOGPIXELSY = 90
PHYSICALWIDTH = 110
PHYSICALHEIGHT = 111
PHYSICALOFFSETX = 112
PHYSICALOFFSETY = 113


def excel_printer_name(value: str) -> str:
    """Document this Excel pipeline operation and its side effects."""
    return value.rsplit(" on ", 1)[0].strip()


class PrinterCapabilityProvider:
    def __init__(self) -> None:
        """Document this Excel pipeline operation and its side effects."""
        self._cache: Dict[str, PrinterCapability] = {}
        self._hard_margin_cache: Dict[
            Tuple[str, str, str, int, int],
            Tuple[float, float, float, float],
        ] = {}

    def enforce(self, app: Any, settings: Any) -> PrinterCapability:
        """Document this Excel pipeline operation and its side effects."""
        requested = settings.printer_name.strip()
        # Excel exposes printer names with a port suffix; compare normalized identities.
        requested_base = excel_printer_name(requested)
        active = excel_printer_name(str(getattr(app, "ActivePrinter", "") or ""))
        fallback: Optional[str] = None
        if (
            settings.printer_policy in {"required", "configured_fallback"}
            and active.casefold() != requested_base.casefold()
        ):
            candidates = [requested]
            try:
                flags = (
                    win32print.PRINTER_ENUM_LOCAL
                    | win32print.PRINTER_ENUM_CONNECTIONS
                )
                for printer in win32print.EnumPrinters(flags, None, 2):
                    name = str(printer.get("pPrinterName", "") or "")
                    if name.casefold() != requested_base.casefold():
                        continue
                    port = str(printer.get("pPortName", "") or "")
                    if port:
                        candidates.append(f"{name} on {port}")
            except Exception:
                pass
            failure: Optional[Exception] = None
            for candidate in dict.fromkeys(candidates):
                # Try the configured spelling and each discovered port spelling.
                try:
                    app.ActivePrinter = candidate
                    active = excel_printer_name(str(app.ActivePrinter or ""))
                    if active.casefold() == requested_base.casefold():
                        failure = None
                        break
                except Exception as exc:
                    failure = exc
            if active.casefold() != requested_base.casefold():
                detail = failure or "Excel rejected every configured printer spelling"
                if settings.printer_policy == "required":
                    raise RuntimeError(
                        f"Required Excel printer {requested!r} is unavailable: {detail}"
                    ) from failure
                fallback = f"configured printer unavailable: {detail}"
        if (
            settings.printer_policy == "required"
            and active.casefold() != requested_base.casefold()
        ):
            raise RuntimeError(
                f"Excel retained printer {active!r}; required {requested_base!r}"
            )
        if not active:
            raise RuntimeError("Excel ActivePrinter is empty")
        # Inspect once and cache form/geometry metadata for deterministic planning.
        capability = self.inspect(active)
        if settings.printer_policy == "required" and (
            not capability.forms or capability.errors
        ):
            detail = "; ".join(capability.errors) or "no advertised forms"
            raise RuntimeError(
                f"Required printer capabilities are unavailable: {detail}"
            )
        if fallback:
            capability = PrinterCapability(
                capability.name, capability.driver, capability.driver_version,
                capability.port, capability.forms, fallback, capability.errors,
            )
        return capability

    def inspect(self, printer_name: str) -> PrinterCapability:
        """Document this Excel pipeline operation and its side effects."""
        cached = self._cache.get(printer_name)
        if cached is not None:
            return cached
        handle = None
        errors = []
        forms = []
        driver = version = port = ""
        try:
            handle = win32print.OpenPrinter(printer_name)
            info = win32print.GetPrinter(handle, 2)
            driver = str(info.get("pDriverName", "") or "")
            port = str(info.get("pPortName", "") or "")
            version = str(info.get("cVersion", "") or "")
            ids = win32print.DeviceCapabilities(printer_name, port, DC_PAPERS) or []
            sizes = win32print.DeviceCapabilities(printer_name, port, DC_PAPERSIZE) or []
            names = win32print.DeviceCapabilities(printer_name, port, DC_PAPERNAMES) or []
            for index, paper_id in enumerate(ids):
                try:
                    width, height = sizes[index]
                    width_in, height_in = float(width) / 254.0, float(height) / 254.0
                    if width_in <= 0 or height_in <= 0:
                        continue
                    forms.append(PrinterFormCapability(
                        int(paper_id), str(names[index]).strip(), width_in, height_in
                    ))
                except (IndexError, TypeError, ValueError, OverflowError) as exc:
                    errors.append(f"form[{index}]: {exc}")
        except Exception as exc:
            errors.append(str(exc))
        finally:
            if handle is not None:
                try:
                    win32print.ClosePrinter(handle)
                except Exception:
                    pass
        result = PrinterCapability(
            printer_name, driver, version, port, tuple(forms), errors=tuple(errors)
        )
        self._cache[printer_name] = result
        return result

    def hard_margins_points(
        self, printer_name: str, paper_enum: int, orientation: int,
    ) -> Tuple[float, float, float, float]:
        """Return left/right/top/bottom hard margins for a form/orientation."""
        base = excel_printer_name(printer_name)
        handle = win32print.OpenPrinter(base)
        hdc = None
        dc = None
        try:
            info = win32print.GetPrinter(handle, 2)
            driver = str(info.get("pDriverName", "") or "")
            port = str(info.get("pPortName", "") or "")
            cache_key = (base.casefold(), driver, port, int(paper_enum), int(orientation))
            cached = self._hard_margin_cache.get(cache_key)
            if cached is not None:
                return cached
            devmode = copy.copy(info.get("pDevMode"))
            if devmode is None:
                raise RuntimeError("printer does not expose a DEVMODE")
            devmode.PaperSize = int(paper_enum)
            devmode.Orientation = int(orientation)
            devmode.Fields |= 0x00000002 | 0x00000001
            hdc = win32gui.CreateDC(driver, base, port, devmode)
            dc = win32ui.CreateDCFromHandle(hdc)
            _, _, margins = self.imageable_geometry(dc)
            result = tuple(value * 72.0 for value in margins)
            self._hard_margin_cache[cache_key] = result
            return result
        finally:
            if dc is not None:
                try:
                    dc.Detach()
                except Exception:
                    pass
            if hdc is not None:
                try:
                    win32gui.DeleteDC(hdc)
                except Exception:
                    pass
            win32print.ClosePrinter(handle)

    @staticmethod
    def imageable_geometry(device_context: Any) -> Tuple[float, float, Tuple[float, float, float, float]]:
        """Convert a configured printer DC's hard geometry to inches."""
        dpi_x = float(device_context.GetDeviceCaps(LOGPIXELSX))
        dpi_y = float(device_context.GetDeviceCaps(LOGPIXELSY))
        if dpi_x <= 0 or dpi_y <= 0:
            raise ValueError("printer DC reported invalid DPI")
        physical_w = float(device_context.GetDeviceCaps(PHYSICALWIDTH)) / dpi_x
        physical_h = float(device_context.GetDeviceCaps(PHYSICALHEIGHT)) / dpi_y
        offset_x = float(device_context.GetDeviceCaps(PHYSICALOFFSETX)) / dpi_x
        offset_y = float(device_context.GetDeviceCaps(PHYSICALOFFSETY)) / dpi_y
        imageable_w = float(device_context.GetDeviceCaps(HORZRES)) / dpi_x
        imageable_h = float(device_context.GetDeviceCaps(VERTRES)) / dpi_y
        right = physical_w - offset_x - imageable_w
        bottom = physical_h - offset_y - imageable_h
        values = (imageable_w, imageable_h, offset_x, right, offset_y, bottom)
        if any(value < 0 for value in values) or imageable_w <= 0 or imageable_h <= 0:
            raise ValueError("printer DC reported impossible imageable geometry")
        return imageable_w, imageable_h, (offset_x, right, offset_y, bottom)
