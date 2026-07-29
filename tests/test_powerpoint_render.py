from __future__ import annotations

from pathlib import Path
import os
from typing import NoReturn

import pytest
from PIL import Image

import pptx_wiki.render as render_module
from pptx_wiki.render import RenderError, render_pptx_powerpoint


class _FakePythonCom:
    COINIT_APARTMENTTHREADED = 2

    def __init__(self) -> None:
        self.initialized = 0
        self.uninitialized = 0
        self.initialize_flags: list[int] = []

    def CoInitializeEx(self, flag: int) -> None:
        self.initialized += 1
        self.initialize_flags.append(flag)

    def CoInitialize(self) -> None:
        self.initialized += 1

    def CoUninitialize(self) -> None:
        self.uninitialized += 1


class _FakeSlide:
    def __init__(self, index: int, calls: list[tuple[object, ...]], *, fail: bool = False) -> None:
        self.index = index
        self.calls = calls
        self.fail = fail

    def Export(self, path: str, filter_name: str, width: int, height: int) -> None:
        self.calls.append((self.index, Path(path).name, filter_name, width, height))
        if self.fail:
            raise RuntimeError("synthetic COM export failure")
        Image.new("RGB", (width, height), "white").save(path, format="PNG")


class _FakeSlides:
    def __init__(self, calls: list[tuple[object, ...]], *, count: int = 2, fail_at: int | None = None) -> None:
        self.Count = count
        self._items = {
            index: _FakeSlide(index, calls, fail=index == fail_at)
            for index in range(1, count + 1)
        }

    def Item(self, index: int) -> _FakeSlide:
        return self._items[index]


class _FakePageSetup:
    SlideWidth = 72.0
    SlideHeight = 36.0


class _FakePresentation:
    def __init__(self, calls: list[tuple[object, ...]], *, count: int = 2, fail_at: int | None = None) -> None:
        self.PageSetup = _FakePageSetup()
        self.Slides = _FakeSlides(calls, count=count, fail_at=fail_at)
        self.closed = False

    def Close(self) -> None:
        self.closed = True


class _FakePresentations:
    def __init__(self, presentation: _FakePresentation) -> None:
        self.presentation = presentation
        self.open_calls: list[tuple[object, ...]] = []

    def Open(self, *args: object) -> _FakePresentation:
        self.open_calls.append(args)
        return self.presentation


class _FakeApplication:
    def __init__(self, presentation: _FakePresentation) -> None:
        self.Presentations = _FakePresentations(presentation)
        self._display_alerts = 2
        self._automation_security = 1
        self.alert_writes: list[int] = []
        self.security_writes: list[int] = []
        self.visible_writes: list[int] = []
        self.quit_called = False

    @property
    def DisplayAlerts(self) -> int:
        return self._display_alerts

    @DisplayAlerts.setter
    def DisplayAlerts(self, value: int) -> None:
        self._display_alerts = value
        self.alert_writes.append(value)

    @property
    def AutomationSecurity(self) -> int:
        return self._automation_security

    @AutomationSecurity.setter
    def AutomationSecurity(self, value: int) -> None:
        self._automation_security = value
        self.security_writes.append(value)

    @property
    def Visible(self) -> int | None:
        return self.visible_writes[-1] if self.visible_writes else None

    @Visible.setter
    def Visible(self, value: int) -> None:
        self.visible_writes.append(value)

    def Quit(self) -> None:
        self.quit_called = True


class _FakeWin32Client:
    def __init__(self, application: _FakeApplication) -> None:
        self.application = application
        self.dispatch_calls: list[str] = []

    def DispatchEx(self, prog_id: str) -> _FakeApplication:
        self.dispatch_calls.append(prog_id)
        return self.application


def _fake_com(*, count: int = 2, fail_at: int | None = None):
    export_calls: list[tuple[object, ...]] = []
    pythoncom = _FakePythonCom()
    presentation = _FakePresentation(export_calls, count=count, fail_at=fail_at)
    application = _FakeApplication(presentation)
    client = _FakeWin32Client(application)
    return pythoncom, client, application, presentation, export_calls


def test_powerpoint_renderer_is_headless_read_only_and_locale_independent(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "한국어 자료.pptx"
    source.write_bytes(b"fake pptx for COM double")
    pythoncom, client, application, presentation, calls = _fake_com()
    monkeypatch.setattr(render_module, "_load_powerpoint_com", lambda: (pythoncom, client))

    output = render_pptx_powerpoint(source, tmp_path / "결과", dpi=72)

    assert [path.name for path in output] == ["slide-0001.png", "slide-0002.png"]
    assert all(path.is_file() for path in output)
    assert calls == [
        (1, "slide-0001.png", "PNG", 72, 36),
        (2, "slide-0002.png", "PNG", 72, 36),
    ]
    assert client.dispatch_calls == ["PowerPoint.Application"]
    assert application.Presentations.open_calls == [(str(source.resolve()), -1, 0, 0)]
    assert application.alert_writes == [1, 2]
    assert application.security_writes == [3, 1]
    assert application.DisplayAlerts == 2
    assert application.AutomationSecurity == 1
    assert application.visible_writes == []
    assert presentation.closed and not application.quit_called
    assert pythoncom.initialized == pythoncom.uninitialized == 1
    assert pythoncom.initialize_flags == [pythoncom.COINIT_APARTMENTTHREADED]


def test_powerpoint_renderer_visible_diagnostic_mode(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "fixture.pptx"
    source.write_bytes(b"fake")
    pythoncom, client, application, _, _ = _fake_com(count=1)
    monkeypatch.setattr(render_module, "_load_powerpoint_com", lambda: (pythoncom, client))

    render_pptx_powerpoint(source, tmp_path / "rendered", dpi=72, with_window=True)

    # WithWindow controls this presentation without changing the global
    # Application.Visible property of a potentially shared PowerPoint session.
    assert application.visible_writes == []
    assert application.Presentations.open_calls[0][1:] == (-1, 0, -1)


def test_powerpoint_renderer_scrubs_configured_secret_during_com_launch(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "fixture.pptx"
    source.write_bytes(b"fake")
    pythoncom, client, _, _, _ = _fake_com(count=1)
    observed: list[str | None] = []
    original_dispatch = client.DispatchEx

    def dispatch(prog_id: str):
        observed.append(os.getenv("PPTX_WIKI_TEST_SECRET"))
        return original_dispatch(prog_id)

    client.DispatchEx = dispatch  # type: ignore[method-assign]
    monkeypatch.setenv("PPTX_WIKI_TEST_SECRET", "do-not-inherit")
    monkeypatch.setattr(render_module, "_load_powerpoint_com", lambda: (pythoncom, client))

    render_pptx_powerpoint(
        source,
        tmp_path / "rendered",
        dpi=72,
        scrub_env_vars=("PPTX_WIKI_TEST_SECRET",),
    )

    assert observed == [None]
    assert os.environ["PPTX_WIKI_TEST_SECRET"] == "do-not-inherit"


def test_powerpoint_renderer_cleans_up_and_publishes_nothing_after_export_failure(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "fixture.pptx"
    source.write_bytes(b"fake")
    pythoncom, client, application, presentation, calls = _fake_com(fail_at=2)
    monkeypatch.setattr(render_module, "_load_powerpoint_com", lambda: (pythoncom, client))
    destination = tmp_path / "rendered"

    with pytest.raises(RenderError, match="exporting slide 2 of 2"):
        render_pptx_powerpoint(source, destination, dpi=72, quit_application=True)

    assert [call[0] for call in calls] == [1, 2]
    assert presentation.closed and application.quit_called
    assert pythoncom.initialized == pythoncom.uninitialized == 1
    assert list(destination.glob("*.png")) == []


def test_powerpoint_renderer_rejects_excessive_pixel_dimensions(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "fixture.pptx"
    source.write_bytes(b"fake")
    pythoncom, client, application, presentation, calls = _fake_com(count=1)
    monkeypatch.setattr(render_module, "_load_powerpoint_com", lambda: (pythoncom, client))

    with pytest.raises(RenderError, match="lower --dpi"):
        render_pptx_powerpoint(
            source,
            tmp_path / "rendered",
            dpi=20_000,
            quit_application=True,
        )

    assert calls == []
    assert presentation.closed and application.quit_called
    assert pythoncom.initialized == pythoncom.uninitialized == 1


def test_powerpoint_validation_fails_before_loading_com(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    def unexpected_loader() -> NoReturn:
        raise AssertionError("COM loader must not run for invalid arguments")

    monkeypatch.setattr(render_module, "_load_powerpoint_com", unexpected_loader)
    with pytest.raises(FileNotFoundError):
        render_pptx_powerpoint(tmp_path / "missing.pptx", tmp_path / "rendered")

    source = tmp_path / "fixture.pptx"
    source.write_bytes(b"fake")
    with pytest.raises(ValueError, match="dpi"):
        render_pptx_powerpoint(source, tmp_path / "rendered", dpi=0)


def test_powerpoint_initialize_failure_is_not_uninitialized(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "fixture.pptx"
    source.write_bytes(b"fake")
    pythoncom, client, _, _, _ = _fake_com(count=1)

    def fail_initialize(_flag: int) -> NoReturn:
        raise RuntimeError("wrong COM apartment")

    pythoncom.CoInitializeEx = fail_initialize  # type: ignore[method-assign]
    monkeypatch.setattr(render_module, "_load_powerpoint_com", lambda: (pythoncom, client))

    with pytest.raises(RenderError, match="initializing COM"):
        render_pptx_powerpoint(source, tmp_path / "rendered", dpi=72)

    assert pythoncom.uninitialized == 0
    assert client.dispatch_calls == []


def test_powerpoint_cleanup_failure_does_not_mask_export_failure(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "fixture.pptx"
    source.write_bytes(b"fake")
    pythoncom, client, application, presentation, _ = _fake_com(fail_at=1)

    def fail_close() -> NoReturn:
        raise RuntimeError("synthetic close failure")

    presentation.Close = fail_close  # type: ignore[method-assign]
    monkeypatch.setattr(render_module, "_load_powerpoint_com", lambda: (pythoncom, client))

    with pytest.raises(RenderError) as captured:
        render_pptx_powerpoint(
            source,
            tmp_path / "rendered",
            dpi=72,
            quit_application=True,
        )

    assert "exporting slide 1 of 2" in str(captured.value)
    assert "Presentation.Close failed" in str(captured.value)
    assert application.quit_called
    assert pythoncom.initialized == pythoncom.uninitialized == 1


def test_powerpoint_failure_preserves_existing_destination_files(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "fixture.pptx"
    source.write_bytes(b"fake")
    destination = tmp_path / "rendered"
    destination.mkdir()
    existing = destination / "slide-0001.png"
    existing.write_bytes(b"user-owned previous result")
    pythoncom, client, _, _, _ = _fake_com(fail_at=2)
    monkeypatch.setattr(render_module, "_load_powerpoint_com", lambda: (pythoncom, client))

    with pytest.raises(RenderError):
        render_pptx_powerpoint(source, destination, dpi=72)

    assert existing.read_bytes() == b"user-owned previous result"
    assert not (destination / "slide-0002.png").exists()
