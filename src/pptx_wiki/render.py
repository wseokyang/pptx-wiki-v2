from __future__ import annotations

import gc
import math
import os
import shutil
import subprocess
import tempfile
import threading
import time
from pathlib import Path
from contextlib import contextmanager
from typing import Any, Iterable, Iterator

from PIL import Image

from .models import DeckRecord, Element
from .roi import emu_to_pixels, non_overlapping_padded_bbox


class RenderError(RuntimeError):
    pass


# Office enum values are deliberately kept local rather than importing the
# generated pywin32 constants cache. DispatchEx then works on a clean machine
# without requiring makepy or mutating the user's gen_py cache.
_MSO_FALSE = 0
_MSO_TRUE = -1
_MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
_PP_ALERTS_NONE = 1
_POWERPOINT_PROG_ID = "PowerPoint.Application"
_MAX_POWERPOINT_EXPORT_PIXELS = 100_000_000
_POWERPOINT_COM_LOCK = threading.Lock()
_RPC_BUSY_HRESULTS = {0x80010001, 0x8001010A}


@contextmanager
def _scrub_process_environment(names: Iterable[str]) -> Iterator[None]:
    """Temporarily remove configured secrets while Office is being launched."""

    selected = tuple(dict.fromkeys(name for name in names if name))
    saved = {name: os.environ[name] for name in selected if name in os.environ}
    for name in selected:
        os.environ.pop(name, None)
    try:
        yield
    finally:
        for name in selected:
            os.environ.pop(name, None)
        os.environ.update(saved)


def _load_powerpoint_com() -> tuple[Any, Any]:
    """Load pywin32 lazily so importing this module remains portable."""

    try:
        import pythoncom
        from win32com import client as win32_client
    except ImportError as exc:  # pragma: no cover - exercised through a patched loader
        raise RenderError(
            "PowerPoint rendering requires Windows, Microsoft PowerPoint, and "
            "pywin32; install the project with the 'windows' extra"
        ) from exc
    return pythoncom, win32_client


def _powerpoint_pixel_size(page_setup: Any, dpi: int) -> tuple[int, int]:
    try:
        width_points = float(page_setup.SlideWidth)
        height_points = float(page_setup.SlideHeight)
    except (AttributeError, TypeError, ValueError) as exc:
        raise RenderError("PowerPoint returned invalid slide dimensions") from exc
    if not math.isfinite(width_points) or not math.isfinite(height_points):
        raise RenderError("PowerPoint returned non-finite slide dimensions")
    if width_points <= 0 or height_points <= 0:
        raise RenderError("PowerPoint returned non-positive slide dimensions")

    # PowerPoint PageSetup dimensions are points (72 points per inch). Avoid
    # Python's bankers-rounding at exact .5 values so the result is stable.
    width = max(1, math.floor(width_points * dpi / 72.0 + 0.5))
    height = max(1, math.floor(height_points * dpi / 72.0 + 0.5))
    if width * height > _MAX_POWERPOINT_EXPORT_PIXELS:
        raise RenderError(
            f"requested PowerPoint export is {width}x{height} pixels; lower --dpi "
            f"to stay below {_MAX_POWERPOINT_EXPORT_PIXELS:,} pixels per slide"
        )
    return width, height


def _verify_powerpoint_export(path: Path, expected_size: tuple[int, int]) -> None:
    if not path.is_file() or path.stat().st_size <= 0:
        raise RenderError(f"PowerPoint did not create the expected slide image: {path.name}")
    try:
        with Image.open(path) as image:
            actual_size = image.size
            image.verify()
    except (OSError, ValueError) as exc:
        raise RenderError(f"PowerPoint created an invalid PNG: {path.name}") from exc
    if actual_size != expected_size:
        raise RenderError(
            f"PowerPoint exported {path.name} at {actual_size[0]}x{actual_size[1]}, "
            f"expected {expected_size[0]}x{expected_size[1]}"
        )


def _com_hresult(error: Exception) -> int | None:
    value = getattr(error, "hresult", None)
    if value is None and error.args and isinstance(error.args[0], int):
        value = error.args[0]
    return int(value) & 0xFFFFFFFF if isinstance(value, int) else None


def _pump_com_messages(pythoncom: Any) -> None:
    pump = getattr(pythoncom, "PumpWaitingMessages", None)
    if callable(pump):
        pump()


def _powerpoint_call(operation: Any, pythoncom: Any, busy_timeout_seconds: float) -> Any:
    """Retry only the two standard 'Office is busy' HRESULTs."""

    deadline = time.monotonic() + busy_timeout_seconds
    delay = 0.1
    while True:
        try:
            return operation()
        except Exception as exc:
            if _com_hresult(exc) not in _RPC_BUSY_HRESULTS or time.monotonic() >= deadline:
                raise
            _pump_com_messages(pythoncom)
            time.sleep(min(delay, max(0.0, deadline - time.monotonic())))
            delay = min(0.5, delay * 1.5)


def _wait_for_full_download(
    presentation: Any,
    pythoncom: Any,
    *,
    timeout_seconds: float,
    busy_timeout_seconds: float,
) -> None:
    """Wait for OneDrive/partial presentations before calling Slide.Export."""

    deadline = time.monotonic() + timeout_seconds
    while True:
        try:
            complete = bool(
                _powerpoint_call(
                    lambda: presentation.IsFullyDownloaded,
                    pythoncom,
                    busy_timeout_seconds,
                )
            )
        except AttributeError:
            # Older PowerPoint object models do not expose this property.
            return
        if complete:
            return
        if time.monotonic() >= deadline:
            raise RenderError("PowerPoint presentation was not fully downloaded before timeout")
        _pump_com_messages(pythoncom)
        time.sleep(min(0.1, max(0.0, deadline - time.monotonic())))


def render_pptx_powerpoint(
    pptx_path: str | Path,
    output_dir: str | Path,
    *,
    dpi: int = 300,
    with_window: bool = False,
    quit_application: bool = False,
    busy_timeout_seconds: float = 30.0,
    download_timeout_seconds: float = 60.0,
    scrub_env_vars: Iterable[str] = (),
) -> list[Path]:
    """Render slides with desktop Microsoft PowerPoint through COM.

    This renderer is intended for an interactive Windows desktop installation,
    where PowerPoint generally preserves fonts and layout more faithfully than
    LibreOffice. It is not suitable for an unattended Windows service: Office
    automation can still display modal dialogs and COM calls have no safe
    in-process timeout.

    PowerPoint is a multi-use, single-instance COM server, so even ``DispatchEx``
    can refer to the same application process the user has open. This function
    therefore closes only its own read-only presentation and does *not* call
    ``Application.Quit`` by default. ``quit_application=True`` is only safe for
    a dedicated interactive worker account/VM where the caller owns the entire
    PowerPoint session. Global alert and macro-security settings are restored.

    In headless mode ``Presentations.Open(..., WithWindow=0)`` is used and
    ``Application.Visible`` is never assigned. ``with_window=True`` changes only
    the Open argument and is useful for troubleshooting desktop-only failures.

    Slides are exported one at a time to explicit ASCII filenames. This avoids
    the localized ``Slide1.PNG``/``슬라이드1.PNG`` names produced by
    ``Presentation.Export`` and includes hidden slides in deck order.
    """

    source = Path(pptx_path).resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    if dpi <= 0:
        raise ValueError("dpi must be positive")
    if busy_timeout_seconds < 0 or download_timeout_seconds < 0:
        raise ValueError("PowerPoint timeout values must be non-negative")
    destination = Path(output_dir).resolve()
    destination.mkdir(parents=True, exist_ok=True)

    pythoncom, win32_client = _load_powerpoint_com()
    # PowerPoint is single-instance and non-reentrant. This lock serializes COM
    # access within this process; separate processes should likewise use one
    # queue/worker rather than attempting parallel PowerPoint automation.
    with _POWERPOINT_COM_LOCK:
        with _scrub_process_environment(scrub_env_vars):
            return _render_pptx_powerpoint_com(
                source,
                destination,
                dpi=dpi,
                with_window=with_window,
                quit_application=quit_application,
                busy_timeout_seconds=busy_timeout_seconds,
                download_timeout_seconds=download_timeout_seconds,
                pythoncom=pythoncom,
                win32_client=win32_client,
            )


def _render_pptx_powerpoint_com(
    source: Path,
    destination: Path,
    *,
    dpi: int,
    with_window: bool,
    quit_application: bool,
    busy_timeout_seconds: float,
    download_timeout_seconds: float,
    pythoncom: Any,
    win32_client: Any,
) -> list[Path]:
    application: Any | None = None
    presentation: Any | None = None
    slides: Any | None = None
    slide: Any | None = None
    com_initialized = False
    previous_alerts: Any | None = None
    previous_security: Any | None = None
    alerts_changed = False
    security_changed = False
    stage = "initializing COM"
    failure: Exception | None = None
    cleanup_errors: list[str] = []
    staged_images: list[Path] = []

    with tempfile.TemporaryDirectory(prefix="pptx-wiki-powerpoint-") as tmp_name:
        staging = Path(tmp_name)
        try:
            co_initialize_ex = getattr(pythoncom, "CoInitializeEx", None)
            if callable(co_initialize_ex):
                co_initialize_ex(getattr(pythoncom, "COINIT_APARTMENTTHREADED", 2))
            else:  # pragma: no cover - modern pywin32 always has CoInitializeEx
                pythoncom.CoInitialize()
            com_initialized = True

            stage = "starting Microsoft PowerPoint"
            application = _powerpoint_call(
                lambda: win32_client.DispatchEx(_POWERPOINT_PROG_ID),
                pythoncom,
                busy_timeout_seconds,
            )
            # Do not use the generated constants module: it requires makepy.
            previous_alerts = _powerpoint_call(
                lambda: application.DisplayAlerts,
                pythoncom,
                busy_timeout_seconds,
            )
            previous_security = _powerpoint_call(
                lambda: application.AutomationSecurity,
                pythoncom,
                busy_timeout_seconds,
            )
            _powerpoint_call(
                lambda: setattr(application, "DisplayAlerts", _PP_ALERTS_NONE),
                pythoncom,
                busy_timeout_seconds,
            )
            alerts_changed = True
            _powerpoint_call(
                lambda: setattr(
                    application,
                    "AutomationSecurity",
                    _MSO_AUTOMATION_SECURITY_FORCE_DISABLE,
                ),
                pythoncom,
                busy_timeout_seconds,
            )
            security_changed = True

            stage = f"opening {source.name}"
            # Positional arguments work with both dynamic and makepy dispatch:
            # FileName, ReadOnly, Untitled, WithWindow.
            presentation = _powerpoint_call(
                lambda: application.Presentations.Open(
                    str(source),
                    _MSO_TRUE,
                    _MSO_FALSE,
                    _MSO_TRUE if with_window else _MSO_FALSE,
                ),
                pythoncom,
                busy_timeout_seconds,
            )
            # Macro security is application-global. Restore it immediately
            # after Open rather than leaving the user's session modified while
            # potentially slow image exports run.
            _powerpoint_call(
                lambda: setattr(application, "AutomationSecurity", previous_security),
                pythoncom,
                busy_timeout_seconds,
            )
            security_changed = False
            _powerpoint_call(
                lambda: setattr(application, "DisplayAlerts", previous_alerts),
                pythoncom,
                busy_timeout_seconds,
            )
            alerts_changed = False
            stage = f"waiting for {source.name} to download"
            _wait_for_full_download(
                presentation,
                pythoncom,
                timeout_seconds=download_timeout_seconds,
                busy_timeout_seconds=busy_timeout_seconds,
            )
            expected_size = _powerpoint_pixel_size(presentation.PageSetup, dpi)
            slides = presentation.Slides
            slide_count = int(slides.Count)
            if slide_count <= 0:
                raise RenderError("PowerPoint presentation contains no slides")

            for index in range(1, slide_count + 1):
                stage = f"exporting slide {index} of {slide_count}"
                slide = _powerpoint_call(
                    lambda index=index: slides.Item(index),
                    pythoncom,
                    busy_timeout_seconds,
                )
                # Slide.Export accepts a complete filename, unlike
                # Presentation.Export, which creates locale-dependent names.
                target = staging / f"slide-{index:04d}.png"
                _powerpoint_call(
                    lambda current_slide=slide, current_target=target: current_slide.Export(
                        str(current_target),
                        "PNG",
                        expected_size[0],
                        expected_size[1],
                    ),
                    pythoncom,
                    busy_timeout_seconds,
                )
                _verify_powerpoint_export(target, expected_size)
                staged_images.append(target)
                slide = None
        except Exception as exc:  # noqa: BLE001 - COM errors have backend-specific classes
            failure = exc
        finally:
            # Release proxies from the leaves inward before COM is uninitialized.
            slide = None
            slides = None
            if presentation is not None:
                try:
                    _powerpoint_call(
                        lambda: presentation.Close(),
                        pythoncom,
                        busy_timeout_seconds,
                    )
                except Exception as exc:  # noqa: BLE001 - cleanup must continue after COM failure
                    cleanup_errors.append(f"Presentation.Close failed: {exc}")
                presentation = None
            if application is not None:
                if security_changed:
                    try:
                        _powerpoint_call(
                            lambda: setattr(
                                application,
                                "AutomationSecurity",
                                previous_security,
                            ),
                            pythoncom,
                            busy_timeout_seconds,
                        )
                        security_changed = False
                    except Exception as exc:  # noqa: BLE001 - restore global state if possible
                        cleanup_errors.append(f"AutomationSecurity restore failed: {exc}")
                if alerts_changed:
                    try:
                        _powerpoint_call(
                            lambda: setattr(application, "DisplayAlerts", previous_alerts),
                            pythoncom,
                            busy_timeout_seconds,
                        )
                        alerts_changed = False
                    except Exception as exc:  # noqa: BLE001 - restore global state if possible
                        cleanup_errors.append(f"DisplayAlerts restore failed: {exc}")
            if application is not None and quit_application:
                try:
                    _powerpoint_call(
                        lambda: application.Quit(),
                        pythoncom,
                        busy_timeout_seconds,
                    )
                except Exception as exc:  # noqa: BLE001 - cleanup must continue after COM failure
                    cleanup_errors.append(f"Application.Quit failed: {exc}")
            application = None
            gc.collect()
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception as exc:  # noqa: BLE001 - cleanup must not mask render failure
                    cleanup_errors.append(f"CoUninitialize failed: {exc}")

        if failure is not None:
            cleanup_detail = f"; cleanup: {'; '.join(cleanup_errors)}" if cleanup_errors else ""
            if isinstance(failure, RenderError):
                message = str(failure)
            else:
                message = f"PowerPoint COM failed while {stage}: {failure}"
            raise RenderError(f"{message}{cleanup_detail}") from failure
        if cleanup_errors:
            raise RenderError("PowerPoint rendered the deck but cleanup failed: " + "; ".join(cleanup_errors))

        rendered: list[Path] = []
        for index, image_path in enumerate(staged_images, start=1):
            target = destination / f"slide-{index:04d}.png"
            shutil.copy2(image_path, target)
            rendered.append(target)

    if not rendered:
        raise RenderError("PowerPoint renderer produced no slide images")
    return rendered


def _require_binary(explicit: str | None, candidates: tuple[str, ...], purpose: str) -> str:
    if explicit:
        resolved = shutil.which(explicit) or (explicit if Path(explicit).is_file() else None)
    else:
        resolved = next((path for name in candidates if (path := shutil.which(name))), None)
    if not resolved:
        joined = "/".join(candidates)
        raise RenderError(f"{purpose} requires {joined}; install it or pass its executable path")
    return resolved


def render_pptx(
    pptx_path: str | Path,
    output_dir: str | Path,
    *,
    dpi: int = 300,
    office_binary: str | None = None,
    pdf_binary: str | None = None,
    scrub_env_vars: Iterable[str] = (),
) -> list[Path]:
    """Render slides through LibreOffice and Poppler.

    Native extraction is independent of rendering. Rendering is only needed for
    pictures, screenshots, SmartArt appearance, or visual QA.
    """

    source = Path(pptx_path).resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    if dpi <= 0:
        raise ValueError("dpi must be positive")
    office = _require_binary(office_binary, ("libreoffice", "soffice"), "PPTX rendering")
    pdftocairo = _require_binary(pdf_binary, ("pdftocairo",), "PDF rasterization")
    destination = Path(output_dir).resolve()
    destination.mkdir(parents=True, exist_ok=True)
    child_env = os.environ.copy()
    for name in scrub_env_vars:
        if name:
            child_env.pop(name, None)

    with tempfile.TemporaryDirectory(prefix="pptx-wiki-render-") as tmp_name:
        tmp = Path(tmp_name)
        profile = tmp / "lo-profile"
        convert = subprocess.run(
            [
                office,
                f"-env:UserInstallation={profile.as_uri()}",
                "--headless",
                "--invisible",
                "--norestore",
                "--nodefault",
                "--nolockcheck",
                "--convert-to",
                "pdf",
                "--outdir",
                str(tmp),
                str(source),
            ],
            capture_output=True,
            text=True,
            env=child_env,
            timeout=300,
            check=False,
        )
        pdf = tmp / f"{source.stem}.pdf"
        if convert.returncode != 0 or not pdf.is_file():
            detail = (convert.stderr or convert.stdout).strip()
            raise RenderError(f"LibreOffice failed to render {source.name}: {detail}")

        prefix = tmp / "slide"
        raster = subprocess.run(
            [pdftocairo, "-png", "-r", str(dpi), str(pdf), str(prefix)],
            capture_output=True,
            text=True,
            env=child_env,
            timeout=600,
            check=False,
        )
        if raster.returncode != 0:
            detail = (raster.stderr or raster.stdout).strip()
            raise RenderError(f"pdftocairo failed: {detail}")

        rendered: list[Path] = []
        pages = sorted(tmp.glob("slide-*.png"), key=lambda path: int(path.stem.rsplit("-", 1)[1]))
        for index, page in enumerate(pages, start=1):
            target = destination / f"slide-{index:04d}.png"
            shutil.copy2(page, target)
            rendered.append(target)
    if not rendered:
        raise RenderError("renderer produced no slide images")
    return rendered


def _is_blocker(element: Element) -> bool:
    return element.kind not in {"background", "connector", "line"} and element.bbox.area > 0


def create_element_crops(
    deck: DeckRecord,
    rendered_slides: list[str | Path],
    output_dir: str | Path,
    *,
    element_kinds: set[str] | None = None,
    padding_ratio: float = 0.01,
    model_padding_px: int = 24,
) -> dict[str, Path]:
    """Create one non-overlapping rendered crop per selected PPT element.

    Source pixels are partitioned at neighbour mid-gaps. Any extra margin the
    OCR model needs is added afterwards as synthetic white pixels, never by
    stealing pixels from the neighbouring object.
    """

    if len(rendered_slides) != len(deck.slides):
        raise ValueError("rendered slide count does not match the PPTX")
    if model_padding_px < 0:
        raise ValueError("model_padding_px must be non-negative")
    # ``image`` is the canonical native-extractor kind. Keep ``picture`` as a
    # compatibility alias for hand-built/interchange records.
    selected = element_kinds or {
        "image",
        "picture",
        "chart",
        "diagram",
        "media",
        "ole",
        "unknown",
    }
    destination = Path(output_dir)
    destination.mkdir(parents=True, exist_ok=True)
    results: dict[str, Path] = {}

    for slide, image_path in zip(deck.slides, rendered_slides, strict=True):
        with Image.open(image_path) as image:
            image.load()
            blockers = [element.bbox for element in slide.elements if _is_blocker(element)]
            for element in slide.elements:
                if element.kind not in selected or element.bbox.area == 0:
                    continue
                crop_emu = non_overlapping_padded_bbox(
                    element.bbox,
                    blockers,
                    slide_width=slide.width,
                    slide_height=slide.height,
                    padding_ratio=padding_ratio,
                )
                crop_px = emu_to_pixels(
                    crop_emu,
                    slide_width=slide.width,
                    slide_height=slide.height,
                    image_width=image.width,
                    image_height=image.height,
                )
                target = destination / f"{element.id}.png"
                capture = image.crop((crop_px.left, crop_px.top, crop_px.right, crop_px.bottom))
                model_image = Image.new(
                    "RGB",
                    (capture.width + 2 * model_padding_px, capture.height + 2 * model_padding_px),
                    "white",
                )
                if capture.mode != "RGB":
                    capture = capture.convert("RGB")
                model_image.paste(capture, (model_padding_px, model_padding_px))
                model_image.save(target)
                if element.asset_path:
                    element.metadata.setdefault("original_asset_path", element.asset_path)
                element.asset_path = str(target)
                element.metadata["ocr_crop_bbox_emu"] = {
                    "x": crop_emu.x,
                    "y": crop_emu.y,
                    "width": crop_emu.width,
                    "height": crop_emu.height,
                }
                element.metadata["ocr_crop_bbox_px"] = {
                    "left": crop_px.left,
                    "top": crop_px.top,
                    "right": crop_px.right,
                    "bottom": crop_px.bottom,
                }
                element.metadata["model_padding_px"] = model_padding_px
                results[element.id] = target
    return results
