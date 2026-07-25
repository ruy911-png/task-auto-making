"""SAP GUI 창을 이미지로 캡처한다 (윈도우 캡처, OS 레벨).

텍스트/실행 로그 대신 조회조건 화면과 조회결과 화면을 각각 캡처해서 감사 증적으로
결과 엑셀에 삽입하는 데 사용한다. Windows 전용 - pywin32 + Pillow 필요.
"""
from __future__ import annotations

from pathlib import Path


def capture_window(window_title_hint: str, output_path: str | Path) -> Path:
    """제목에 window_title_hint를 포함하는 창을 찾아 PNG로 캡처한다."""
    try:
        import win32gui
        import win32ui
        from ctypes import windll
        from PIL import Image
    except ImportError as exc:  # pragma: no cover - Windows 환경 밖에서는 항상 발생
        raise RuntimeError(
            "화면 캡처에는 pywin32와 Pillow가 필요합니다. Windows 환경에서 설치 후 실행하세요."
        ) from exc

    hwnd = _find_window(window_title_hint)
    if hwnd is None:
        raise RuntimeError(f"'{window_title_hint}' 제목을 포함한 SAP 창을 찾지 못했습니다.")

    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width, height = right - left, bottom - top

    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
    save_dc.SelectObject(bitmap)
    windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 2)

    bmpinfo = bitmap.GetInfo()
    bmpstr = bitmap.GetBitmapBits(True)
    image = Image.frombuffer(
        "RGB", (bmpinfo["bmWidth"], bmpinfo["bmHeight"]), bmpstr, "raw", "BGRX", 0, 1
    )

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(output_path)

    win32gui.DeleteObject(bitmap.GetHandle())
    save_dc.DeleteDC()
    mfc_dc.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwnd_dc)

    return output_path


def _find_window(title_hint: str) -> int | None:
    import win32gui

    matches: list[int] = []

    def _enum_handler(hwnd: int, _: object) -> None:
        if win32gui.IsWindowVisible(hwnd) and title_hint in win32gui.GetWindowText(hwnd):
            matches.append(hwnd)

    win32gui.EnumWindows(_enum_handler, None)
    return matches[0] if matches else None
