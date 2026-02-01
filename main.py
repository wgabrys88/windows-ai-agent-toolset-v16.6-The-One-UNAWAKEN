from __future__ import annotations

"""
main_storyhud.py

Architecture Overview:
This script implements a stateless AI agent for controlling Windows via visual feedback, where the 'story-memory' overlay on screenshots serves as the sole persistent memory, embodying the agent's 'self-awareness'—ironically, the story is the AI! The system captures screenshots with overlays, processes them through a Vision-Language Model (VLM) to decide actions and update the story, executes inputs (mouse/keyboard), and waits for UI stability. It's designed for tasks like opening apps and drawing in Paint, relying on atemporal, causal descriptions in the story to maintain context across stateless API calls. No external dependencies beyond standard libraries; all Win32 interactions are via ctypes, and PNG encoding is custom-implemented.

Goal:
- Overlay text is the *only memory*, visible in screenshots (stateless API).
- Make overlay readable (white fill with black outline).
- Make overlay long enough (multi-line, proper wrapping; avoid DrawText clipping).
- Make screenshots less "too fast" after actions (wait-for-stability).

This file is derived from the project's original main.py, with HUD + timing upgrades.
"""

import base64
import ctypes
import ctypes.wintypes as w
import json
import re
import struct
import time
import urllib.request
import urllib.error
import zlib
from dataclasses import dataclass
from enum import IntFlag
from functools import cache
from pathlib import Path
from typing import Any, Literal

# =========================
# MODEL / API CONFIG
# =========================
MODEL_NAME = "qwen3-vl-2b-instruct"
API_URL = "http://localhost:1234/v1/chat/completions"

# =========================
# SCREENSHOT CONFIG
# =========================
SCREENSHOT_QUALITY = 1
SCREEN_W, SCREEN_H = {1: (1536, 864), 2: (1024, 576), 3: (512, 288)}[SCREENSHOT_QUALITY]

# =========================
# INPUT / TIMING CONFIG
# =========================
INPUT_DELAY_S = 0.10

# Baseline delays (still used), but we also apply a "wait until stable" after actions.
DELAY_AFTER_CLICK_S = 0.85
DELAY_AFTER_TYPE_S = 0.65
DELAY_AFTER_DRAG_S = 0.60
DELAY_MOVE_HOVER_S = 1.20
DELAY_SCROLL_S = 0.30

# Additional heuristics (normalized coords 0..1000)
DELAY_START_MENU_OPEN_S = 1.05

# Screen stability (post-action)
SETTLE_ENABLED = True
SETTLE_MAX_S = 2.5
SETTLE_SAMPLE_W, SETTLE_SAMPLE_H = 256, 144
SETTLE_CHECK_INTERVAL_S = 0.10
SETTLE_REQUIRED_STABLE = 2
SETTLE_CHANGE_RATIO_THRESHOLD = 0.006  # ~0.6% sampled pixels

# =========================
# HUD / OVERLAY CONFIG
# =========================
HUD_MARGIN = 10
HUD_MAX_WIDTH = 1400

# More lines than the original so the "story-memory" can be several sentences.
HUD_LINES_PRIORITY = 6
HUD_LINES_DETAIL = 8
HUD_LINES_FADE = 8

# Font tiers: priority/detail/fade
HUD_FONT_SIZE_PRIORITY = -26
HUD_FONT_SIZE_DETAIL = -18
HUD_FONT_SIZE_FADE = -14
HUD_FONT_WEIGHT = 700
HUD_LINE_SPACING = 3

# Text readability: white fill, black outline.
HUD_TEXT_COLOR = 0x00FFFFFF  # COLORREF 0x00bbggrr -> white
HUD_OUTLINE_COLOR = 0x00000000
HUD_OUTLINE_PX = 2

# Optional translucent background panel behind text (helps in very bright apps).
HUD_BG_ENABLED = True
HUD_BG_COLOR_BGR = (0, 0, 0)   # black in BGR order
HUD_BG_ALPHA = 110             # 0..255 (110 ~ 43% opacity)

OVERLAY_REASSERT_PULSES = 2
OVERLAY_REASSERT_PAUSE_S = 0.05

DEFAULT_TASK = (
    "Open Microsoft Paint from the Start menu then use the mouse to draw a simple cat face "
    "with two circles for eyes one triangle for nose and curved line for smile then save the "
    "file as cat in the Pictures folder and close Paint when done"
)

DEFAULT_HUD_TEST_MESSAGE = (
    "Desktop shows Windows taskbar at bottom edge. Start button exists at lower-left. "
    "Start menu contains search box and pinned apps. Paint appears in results as icon. "
    "Clicking Paint launches application. Paint canvas is blank for drawing. "
    "Cat face uses two circles for eyes, triangle for nose, curved line for smile. "
    "Save uses File menu and Pictures folder path. Close button top-right ends session."
)

ActionTool = Literal["click", "move", "drag", "type", "scroll", "done", "analyze"]

# =========================
# WIN32 SETUP
# =========================
@cache
def _dll(name: str) -> ctypes.WinDLL:
    return ctypes.WinDLL(name, use_last_error=True)

user32 = _dll("user32")
gdi32 = _dll("gdi32")
kernel32 = _dll("kernel32")

try:
    ctypes.WinDLL("Shcore", use_last_error=True).SetProcessDpiAwareness(2)
except Exception:
    try:
        user32.SetProcessDPIAware()
    except Exception:
        pass

class MouseEvent(IntFlag):
    MOVE = 0x0001
    ABSOLUTE = 0x8000
    LEFT_DOWN = 0x0002
    LEFT_UP = 0x0004
    WHEEL = 0x0800
    HWHEEL = 0x1000

class KeyEvent(IntFlag):
    KEYUP = 0x0002
    UNICODE = 0x0004

class WinStyle(IntFlag):
    EX_TOPMOST = 0x00000008
    EX_LAYERED = 0x00080000
    EX_TRANSPARENT = 0x00000020
    EX_NOACTIVATE = 0x08000000
    EX_TOOLWINDOW = 0x00000080
    POPUP = 0x80000000

INPUT_MOUSE = 0
INPUT_KEYBOARD = 1
WHEEL_DELTA = 120
SRCCOPY = 0x00CC0020
SW_SHOWNOACTIVATE = 4
SW_HIDE = 0
ULW_ALPHA = 2
AC_SRC_ALPHA = 1
SWP_NOSIZE = 1
SWP_NOMOVE = 2
SWP_NOACTIVATE = 16
SWP_SHOWWINDOW = 64
HWND_TOPMOST = -1
CURSOR_SHOWING = 0x00000001
TRANSPARENT = 1
DT_LEFT = 0x00000000
DT_NOPREFIX = 0x00000800
DI_NORMAL = 0x0003

LRESULT = ctypes.c_ssize_t
WPARAM = ctypes.c_size_t
LPARAM = ctypes.c_ssize_t
WNDPROC = ctypes.WINFUNCTYPE(LRESULT, w.HWND, w.UINT, WPARAM, LPARAM)
ULONG_PTR = ctypes.c_size_t

class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", w.LONG), ("dy", w.LONG), ("mouseData", w.DWORD),
        ("dwFlags", w.DWORD), ("time", w.DWORD), ("dwExtraInfo", ULONG_PTR),
    ]

class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", w.WORD), ("wScan", w.WORD), ("dwFlags", w.DWORD),
        ("time", w.DWORD), ("dwExtraInfo", ULONG_PTR),
    ]

class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [("uMsg", w.DWORD), ("wParamL", w.WORD), ("wParamH", w.WORD)]

class _INPUTunion(ctypes.Union):
    _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT), ("hi", HARDWAREINPUT)]

class INPUT(ctypes.Structure):
    _anonymous_ = ("u",)
    _fields_ = [("type", w.DWORD), ("u", _INPUTunion)]

class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", w.DWORD), ("biWidth", w.LONG), ("biHeight", w.LONG),
        ("biPlanes", w.WORD), ("biBitCount", w.WORD), ("biCompression", w.DWORD),
        ("biSizeImage", w.DWORD), ("biXPelsPerMeter", w.LONG),
        ("biYPelsPerMeter", w.LONG), ("biClrUsed", w.DWORD), ("biClrImportant", w.DWORD),
    ]

class BITMAPINFO(ctypes.Structure):
    _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", ctypes.c_uint * 1)]

class CURSORINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", w.DWORD), ("flags", w.DWORD),
        ("hCursor", w.HANDLE), ("ptScreenPos", w.POINT),
    ]

class ICONINFO(ctypes.Structure):
    _fields_ = [
        ("fIcon", w.BOOL), ("xHotspot", w.DWORD), ("yHotspot", w.DWORD),
        ("hbmMask", w.HBITMAP), ("hbmColor", w.HBITMAP),
    ]

class BLENDFUNCTION(ctypes.Structure):
    _fields_ = [
        ("BlendOp", ctypes.c_ubyte), ("BlendFlags", ctypes.c_ubyte),
        ("SourceConstantAlpha", ctypes.c_ubyte), ("AlphaFormat", ctypes.c_ubyte),
    ]

class WNDCLASS(ctypes.Structure):
    _fields_ = [
        ("style", ctypes.c_uint), ("lpfnWndProc", ctypes.c_void_p),
        ("cbClsExtra", ctypes.c_int), ("cbWndExtra", ctypes.c_int),
        ("hInstance", w.HINSTANCE), ("hIcon", w.HANDLE), ("hCursor", w.HANDLE),
        ("hbrBackground", w.HANDLE), ("lpszMenuName", w.LPCWSTR),
        ("lpszClassName", w.LPCWSTR),
    ]

user32.DefWindowProcW.argtypes = [w.HWND, w.UINT, WPARAM, LPARAM]
user32.DefWindowProcW.restype = LRESULT

_SendInput = user32.SendInput
_SendInput.argtypes = (w.UINT, ctypes.POINTER(INPUT), ctypes.c_int)
_SendInput.restype = w.UINT

user32.DrawTextW.argtypes = [w.HDC, w.LPCWSTR, ctypes.c_int, ctypes.POINTER(w.RECT), w.UINT]
user32.DrawTextW.restype = ctypes.c_int

gdi32.GetTextExtentPoint32W.argtypes = [w.HDC, w.LPCWSTR, ctypes.c_int, ctypes.POINTER(w.SIZE)]
gdi32.GetTextExtentPoint32W.restype = w.BOOL

gdi32.CreateCompatibleDC.argtypes = [w.HDC]
gdi32.CreateCompatibleDC.restype = w.HDC
gdi32.CreateDIBSection.argtypes = [
    w.HDC, ctypes.POINTER(BITMAPINFO), w.UINT,
    ctypes.POINTER(ctypes.c_void_p), w.HANDLE, w.DWORD
]
gdi32.CreateDIBSection.restype = w.HBITMAP
gdi32.SelectObject.argtypes = [w.HDC, w.HGDIOBJ]
gdi32.SelectObject.restype = w.HGDIOBJ
gdi32.BitBlt.argtypes = [
    w.HDC, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    w.HDC, ctypes.c_int, ctypes.c_int, w.DWORD
]
gdi32.BitBlt.restype = w.BOOL
gdi32.DeleteObject.argtypes = [w.HGDIOBJ]
gdi32.DeleteObject.restype = w.BOOL
gdi32.DeleteDC.argtypes = [w.HDC]
gdi32.DeleteDC.restype = w.BOOL
gdi32.SetBkMode.argtypes = [w.HDC, ctypes.c_int]
gdi32.SetBkMode.restype = ctypes.c_int
gdi32.SetTextColor.argtypes = [w.HDC, w.DWORD]
gdi32.SetTextColor.restype = w.DWORD
gdi32.CreateFontW.restype = w.HFONT

user32.ReleaseDC.argtypes = [w.HWND, w.HDC]
user32.ReleaseDC.restype = ctypes.c_int
user32.GetCursorInfo.argtypes = [ctypes.POINTER(CURSORINFO)]
user32.GetCursorInfo.restype = w.BOOL
user32.GetIconInfo.argtypes = [w.HICON, ctypes.POINTER(ICONINFO)]
user32.GetIconInfo.restype = w.BOOL
user32.DrawIconEx.argtypes = [
    w.HDC, ctypes.c_int, ctypes.c_int, w.HICON, ctypes.c_int,
    ctypes.c_int, w.UINT, w.HBRUSH, w.UINT
]
user32.DrawIconEx.restype = w.BOOL

user32.UpdateLayeredWindow.argtypes = [
    w.HWND, w.HDC, ctypes.POINTER(w.POINT), ctypes.POINTER(w.SIZE), w.HDC,
    ctypes.POINTER(w.POINT), w.DWORD, ctypes.POINTER(BLENDFUNCTION), w.DWORD
]
user32.UpdateLayeredWindow.restype = w.BOOL

user32.SetWindowPos.argtypes = [
    w.HWND, w.HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, w.UINT
]
user32.SetWindowPos.restype = w.BOOL

# =========================
# INPUT HELPERS
# =========================
def get_screen_size() -> tuple[int, int]:
    return user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)

@dataclass(slots=True)
class CoordConverter:
    sw: int
    sh: int
    mw: int
    mh: int

    def norm_to_screen(self, xn: float, yn: float) -> tuple[int, int]:
        return int(xn * self.sw / 1000), int(yn * self.sh / 1000)

    def to_win32_normalized(self, x: int, y: int) -> tuple[int, int]:
        if self.sw <= 0 or self.sh <= 0:
            return 0, 0
        return (
            max(0, min(65535, int(x * 65535 / (self.sw - 1)))) if self.sw > 1 else 0,
            max(0, min(65535, int(y * 65535 / (self.sh - 1)))) if self.sh > 1 else 0,
        )

def _send_input(inputs: list[INPUT], *, delay_s: float = INPUT_DELAY_S) -> None:
    arr = (INPUT * len(inputs))(*inputs)
    if _SendInput(len(inputs), arr, ctypes.sizeof(INPUT)) != len(inputs):
        raise ctypes.WinError(ctypes.get_last_error())
    if delay_s > 0:
        time.sleep(delay_s)

def mouse_move(x: int, y: int, conv: CoordConverter) -> None:
    ax, ay = conv.to_win32_normalized(x, y)
    i = INPUT(type=INPUT_MOUSE)
    i.mi = MOUSEINPUT(ax, ay, 0, MouseEvent.MOVE | MouseEvent.ABSOLUTE, 0, 0)
    _send_input([i])

def mouse_click(x: int, y: int, conv: CoordConverter) -> None:
    ax, ay = conv.to_win32_normalized(x, y)
    inputs = []
    for flag in (MouseEvent.MOVE, MouseEvent.LEFT_DOWN, MouseEvent.LEFT_UP):
        i = INPUT(type=INPUT_MOUSE)
        i.mi = MOUSEINPUT(ax, ay, 0, int(flag) | int(MouseEvent.ABSOLUTE), 0, 0)
        inputs.append(i)
    _send_input(inputs)

def mouse_drag(x1: int, y1: int, x2: int, y2: int, conv: CoordConverter) -> None:
    ax1, ay1 = conv.to_win32_normalized(x1, y1)
    ax2, ay2 = conv.to_win32_normalized(x2, y2)

    def send(flags: int, dx: int, dy: int, *, delay: float = INPUT_DELAY_S) -> None:
        inp = INPUT(type=INPUT_MOUSE)
        inp.mi = MOUSEINPUT(dx, dy, 0, flags, 0, 0)
        _send_input([inp], delay_s=delay)

    send(int(MouseEvent.MOVE | MouseEvent.ABSOLUTE), ax1, ay1)
    send(int(MouseEvent.LEFT_DOWN | MouseEvent.ABSOLUTE), ax1, ay1)

    for k in range(1, 15):
        t = k / 14
        dx = int(ax1 + (ax2 - ax1) * t)
        dy = int(ay1 + (ay2 - ay1) * t)
        send(int(MouseEvent.MOVE | MouseEvent.ABSOLUTE), dx, dy, delay=0.0)
        time.sleep(0.01)

    send(int(MouseEvent.LEFT_UP | MouseEvent.ABSOLUTE), ax2, ay2)

def type_text(text: str) -> None:
    if not text:
        return
    inputs = []
    for ch in text:
        b = ch.encode("utf-16le")
        for i in range(0, len(b), 2):
            cu = b[i] | (b[i + 1] << 8)
            for flags in (KeyEvent.UNICODE, KeyEvent.UNICODE | KeyEvent.KEYUP):
                inp = INPUT(type=INPUT_KEYBOARD)
                inp.ki = KEYBDINPUT(0, cu, int(flags), 0, 0)
                inputs.append(inp)
    _send_input(inputs)

def scroll(dx: float = 0.0, dy: float = 0.0) -> None:
    inputs = []
    for delta, flag in ((dy, MouseEvent.WHEEL), (dx, MouseEvent.HWHEEL)):
        if delta:
            ticks = max(1, abs(int(delta)) // 100)
            direction = 1 if delta > 0 else -1
            for _ in range(ticks):
                inp = INPUT(type=INPUT_MOUSE)
                inp.mi = MOUSEINPUT(0, 0, WHEEL_DELTA * direction, int(flag), 0, 0)
                inputs.append(inp)
    if inputs:
        _send_input(inputs)

# =========================
# SCREEN CAPTURE + PNG ENCODE
# =========================
def _capture_desktop_bgra(sw: int, sh: int, *, include_cursor: bool = True) -> bytes:
    sdc = user32.GetDC(0)
    if not sdc:
        raise ctypes.WinError(ctypes.get_last_error())

    mdc = gdi32.CreateCompatibleDC(sdc)
    if not mdc:
        user32.ReleaseDC(0, sdc)
        raise ctypes.WinError(ctypes.get_last_error())

    bmi = BITMAPINFO()
    bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bmi.bmiHeader.biWidth = sw
    bmi.bmiHeader.biHeight = -sh
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 32

    bits = ctypes.c_void_p()
    hbm = gdi32.CreateDIBSection(sdc, ctypes.byref(bmi), 0, ctypes.byref(bits), None, 0)
    if not hbm:
        gdi32.DeleteDC(mdc)
        user32.ReleaseDC(0, sdc)
        raise ctypes.WinError(ctypes.get_last_error())

    gdi32.SelectObject(mdc, hbm)
    if not gdi32.BitBlt(mdc, 0, 0, sw, sh, sdc, 0, 0, SRCCOPY):
        gdi32.DeleteObject(hbm)
        gdi32.DeleteDC(mdc)
        user32.ReleaseDC(0, sdc)
        raise ctypes.WinError(ctypes.get_last_error())

    if include_cursor:
        ci = CURSORINFO(cbSize=ctypes.sizeof(CURSORINFO))
        if user32.GetCursorInfo(ctypes.byref(ci)) and ci.flags & CURSOR_SHOWING:
            ii = ICONINFO()
            if user32.GetIconInfo(ci.hCursor, ctypes.byref(ii)):
                x = ci.ptScreenPos.x - ii.xHotspot
                y = ci.ptScreenPos.y - ii.yHotspot
                user32.DrawIconEx(mdc, x, y, ci.hCursor, 0, 0, 0, 0, DI_NORMAL)
                if ii.hbmMask:
                    gdi32.DeleteObject(ii.hbmMask)
                if ii.hbmColor:
                    gdi32.DeleteObject(ii.hbmColor)

    out = ctypes.string_at(bits, sw * sh * 4)
    user32.ReleaseDC(0, sdc)
    gdi32.DeleteDC(mdc)
    gdi32.DeleteObject(hbm)
    return out

def _downsample_nn_bgra(bgra: bytes, sw: int, sh: int, dw: int, dh: int) -> bytes:
    if sw == dw and sh == dh:
        return bgra
    out = bytearray(dw * dh * 4)
    rx = sw / dw
    ry = sh / dh
    for y in range(dh):
        sy = int(y * ry) * sw * 4
        dy = y * dw * 4
        for x in range(dw):
            sx = int(x * rx) * 4
            out[dy:dy+4] = bgra[sy + sx:sy + sx + 4]
            dy += 4
    return bytes(out)

def _encode_png_rgb(data: bytes, w: int, h: int) -> bytes:
    # Convert BGRA to RGB (drop alpha)
    rgb = bytearray(len(data) * 3 // 4)
    j = 0
    for i in range(0, len(data), 4):
        rgb[j] = data[i + 2]  # R
        rgb[j + 1] = data[i + 1]  # G
        rgb[j + 2] = data[i]  # B
        j += 3
    rgb = bytes(rgb)

    row_len = w * 3
    # Apply filter (none) to each row and concatenate
    filtered_data = b''
    for i in range(0, len(rgb), row_len):
        filtered_data += b'\x00' + rgb[i:i + row_len]

    # Compress the entire filtered data
    compressed = zlib.compress(filtered_data)

    # IHDR: width, height, bit depth 8, color type 2 (RGB), compression 0, filter 0, interlace 0
    ihdr = struct.pack('>iiBBBBB', w, h, 8, 2, 0, 0, 0)
    ihdr_chunk = struct.pack('>I', len(ihdr)) + b'IHDR' + ihdr + struct.pack('>I', zlib.crc32(b'IHDR' + ihdr))

    # IDAT chunk
    idat_chunk = struct.pack('>I', len(compressed)) + b'IDAT' + compressed + struct.pack('>I', zlib.crc32(b'IDAT' + compressed))

    # IEND chunk
    iend_chunk = struct.pack('>I', 0) + b'IEND' + struct.pack('>I', zlib.crc32(b'IEND'))

    # PNG signature
    header = b'\x89PNG\r\n\x1a\n'

    return header + ihdr_chunk + idat_chunk + iend_chunk

# =========================
# SCREEN STABILITY
# =========================
def wait_for_screen_settle(conv: CoordConverter) -> None:
    if not SETTLE_ENABLED:
        return
    start = time.time()
    stable_count = 0
    prev = None
    while time.time() - start < SETTLE_MAX_S:
        time.sleep(SETTLE_CHECK_INTERVAL_S)
        curr = _capture_desktop_bgra(conv.sw, conv.sh, include_cursor=False)
        curr = _downsample_nn_bgra(curr, conv.sw, conv.sh, SETTLE_SAMPLE_W, SETTLE_SAMPLE_H)
        if prev is not None:
            diff = sum(1 for a, b in zip(prev, curr) if a != b) / len(curr)
            if diff < SETTLE_CHANGE_RATIO_THRESHOLD:
                stable_count += 1
                if stable_count >= SETTLE_REQUIRED_STABLE:
                    return
            else:
                stable_count = 0
        prev = curr

# =========================
# OVERLAY MANAGER
# =========================
def _fill_rect_bgra(bits: ctypes.c_void_p, w: int, h: int, rect: tuple[int, int, int, int], bgr: tuple[int, int, int], alpha: int) -> None:
    x1, y1, x2, y2 = rect
    x1 = max(0, x1)
    y1 = max(0, y1)
    x2 = min(w, x2)
    y2 = min(h, y2)
    data = ctypes.cast(bits, ctypes.POINTER(ctypes.c_ubyte))
    for y in range(y1, y2):
        row = y * w * 4
        for x in range(x1, x2):
            off = row + x * 4
            data[off] = bgr[0]
            data[off + 1] = bgr[1]
            data[off + 2] = bgr[2]
            data[off + 3] = alpha

def _draw_text_outlined(hdc: w.HDC, text: str, rect: w.RECT, flags: int) -> None:
    gdi32.SetTextColor(hdc, HUD_OUTLINE_COLOR)
    for dx, dy in [(-HUD_OUTLINE_PX, 0), (HUD_OUTLINE_PX, 0), (0, -HUD_OUTLINE_PX), (0, HUD_OUTLINE_PX)]:
        r = w.RECT(rect.left + dx, rect.top + dy, rect.right + dx, rect.bottom + dy)
        user32.DrawTextW(hdc, text, -1, ctypes.byref(r), flags)
    gdi32.SetTextColor(hdc, HUD_TEXT_COLOR)
    user32.DrawTextW(hdc, text, -1, ctypes.byref(rect), flags)

class OverlayManager:
    def __init__(self, w: int, h: int):
        self.w = w
        self.h = h
        self.story = ""

        self.hwnd = None
        self.hdc = None
        self.bits = None
        self._fonts = []

    def __enter__(self) -> "OverlayManager":
        class_name = "OverlayWnd"
        wc = WNDCLASS(0, ctypes.cast(user32.DefWindowProcW, ctypes.c_void_p), 0, 0, kernel32.GetModuleHandleW(None), 0, 0, 0, None, class_name)
        user32.RegisterClassW(ctypes.byref(wc))

        style = int(WinStyle.EX_LAYERED | WinStyle.EX_TRANSPARENT | WinStyle.EX_TOPMOST | WinStyle.EX_NOACTIVATE | WinStyle.EX_TOOLWINDOW | WinStyle.POPUP)
        self.hwnd = user32.CreateWindowExW(style, class_name, None, 0, 0, 0, self.w, self.h, 0, 0, wc.hInstance, None)
        if not self.hwnd:
            raise ctypes.WinError(ctypes.get_last_error())

        bmi = BITMAPINFO()
        bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bmi.bmiHeader.biWidth = self.w
        bmi.bmiHeader.biHeight = -self.h
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = 0

        self.bits = ctypes.c_void_p()
        hbm = gdi32.CreateDIBSection(None, ctypes.byref(bmi), 0, ctypes.byref(self.bits), None, 0)
        if not hbm:
            raise ctypes.WinError(ctypes.get_last_error())

        self.hdc = gdi32.CreateCompatibleDC(0)
        if not self.hdc:
            gdi32.DeleteObject(hbm)
            raise ctypes.WinError(ctypes.get_last_error())

        gdi32.SelectObject(self.hdc, hbm)
        gdi32.SetBkMode(self.hdc, TRANSPARENT)

        font_sizes = [HUD_FONT_SIZE_PRIORITY, HUD_FONT_SIZE_DETAIL, HUD_FONT_SIZE_FADE]
        for size in font_sizes:
            font = gdi32.CreateFontW(size, 0, 0, 0, HUD_FONT_WEIGHT, 0, 0, 0, 0, 0, 0, 0, 0, "Segoe UI")
            self._fonts.append(font)

        user32.ShowWindow(self.hwnd, SW_SHOWNOACTIVATE)
        self.render()
        return self

    def __exit__(self, *args: Any) -> None:
        if self.hwnd:
            user32.DestroyWindow(self.hwnd)
        if self.hdc:
            gdi32.DeleteDC(self.hdc)
        for font in self._fonts:
            gdi32.DeleteObject(font)

    def reassert_topmost(self) -> None:
        for _ in range(OVERLAY_REASSERT_PULSES):
            user32.SetWindowPos(self.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW | SWP_NOACTIVATE)
            time.sleep(OVERLAY_REASSERT_PAUSE_S)

    def set_story(self, story: str) -> None:
        self.story = story

    def render(self) -> None:
        # Clear bitmap to transparent
        ctypes.memset(self.bits, 0, self.w * self.h * 4)

        if not self.story:
            self.reassert_topmost()
            return

        lines = [ln.strip() for ln in self.story.splitlines() if ln.strip()]
        sections = [
            lines[:HUD_LINES_PRIORITY],
            lines[HUD_LINES_PRIORITY:HUD_LINES_PRIORITY + HUD_LINES_DETAIL],
            lines[HUD_LINES_PRIORITY + HUD_LINES_DETAIL:HUD_LINES_PRIORITY + HUD_LINES_DETAIL + HUD_LINES_FADE],
        ]

        font_sizes = [HUD_FONT_SIZE_PRIORITY, HUD_FONT_SIZE_DETAIL, HUD_FONT_SIZE_FADE]
        max_width = 0
        for idx, lines in enumerate(sections):
            if not lines:
                continue
            gdi32.SelectObject(self.hdc, self._fonts[idx])
            for line in lines:
                sz = w.SIZE()
                gdi32.GetTextExtentPoint32W(self.hdc, line, len(line), ctypes.byref(sz))
                max_width = max(max_width, sz.cx)

        if HUD_BG_ENABLED and any(sections):
            est = 0
            for idx, lines in enumerate(sections):
                if not lines:
                    continue
                est += len(lines) * (abs(font_sizes[idx]) + HUD_LINE_SPACING)
            bg_rect = (HUD_MARGIN - 6, HUD_MARGIN - 6, HUD_MARGIN + max_width + 12, HUD_MARGIN + est + 12)
            _fill_rect_bgra(self.bits, self.w, self.h, bg_rect, HUD_BG_COLOR_BGR, HUD_BG_ALPHA)

        x = HUD_MARGIN
        y = HUD_MARGIN
        for font_idx, (font, lines) in enumerate(zip(self._fonts, sections)):
            if not lines:
                continue
            gdi32.SelectObject(self.hdc, font)
            line_height = abs(font_sizes[font_idx]) + HUD_LINE_SPACING
            for line in lines:
                rect = w.RECT(x, y, x + max_width, y + line_height)
                _draw_text_outlined(self.hdc, line, rect, DT_LEFT | DT_NOPREFIX)
                y += line_height

        bf = BLENDFUNCTION(0, 0, 255, AC_SRC_ALPHA)
        sz = w.SIZE(self.w, self.h)
        ps = w.POINT(0, 0)
        pd = w.POINT(0, 0)

        if not user32.UpdateLayeredWindow(
            self.hwnd, 0, ctypes.byref(pd), ctypes.byref(sz),
            self.hdc, ctypes.byref(ps), 0, ctypes.byref(bf), ULW_ALPHA
        ):
            raise ctypes.WinError(ctypes.get_last_error())

        self.reassert_topmost()

# =========================
# ACTION COMMAND
# =========================
@dataclass(slots=True)
class ActionCommand:
    tool: ActionTool
    reasoning: str = ""
    memory: Any = None  # can be list[str] or str
    x: float | None = None
    y: float | None = None
    text: str = ""
    dx: float = 0.0
    dy: float = 0.0
    x1: float | None = None
    y1: float | None = None
    x2: float | None = None
    y2: float | None = None

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "ActionCommand":
        num = lambda v: float(v[0]) if isinstance(v, list) and v else float(v) if v is not None and v != "" else None
        tool = d.get("tool", "")
        if tool not in {"click", "move", "drag", "type", "scroll", "done", "analyze"}:
            tool = "done"
        return cls(
            tool=tool,
            reasoning=str(d.get("reasoning", "") or "").strip(),
            memory=d.get("memory", None),
            x=num(d.get("x")),
            y=num(d.get("y")),
            text=str(d.get("text", "") or ""),
            dx=num(d.get("dx")) or 0.0,
            dy=num(d.get("dy")) or 0.0,
            x1=num(d.get("x1")),
            y1=num(d.get("y1")),
            x2=num(d.get("x2")),
            y2=num(d.get("y2")),
        )

    def validate(self) -> bool:
        match self.tool:
            case "click" | "move":
                return self.x is not None and self.y is not None and 0 <= self.x <= 1000 and 0 <= self.y <= 1000
            case "drag":
                return all(v is not None and 0 <= v <= 1000 for v in [self.x1, self.y1, self.x2, self.y2])
            case "scroll":
                return abs(self.dx) <= 10000 and abs(self.dy) <= 10000
            case "type":
                return len(self.text) <= 2000
            case "analyze" | "done":
                return True
            case _:
                return False

# =========================
# EXECUTOR
# =========================
class ActionExecutor:
    def __init__(self, conv: CoordConverter):
        self.conv = conv

    def execute(self, cmd: ActionCommand) -> float:
        match cmd.tool:
            case "click":
                assert cmd.x is not None and cmd.y is not None
                sx, sy = self.conv.norm_to_screen(cmd.x, cmd.y)
                mouse_click(sx, sy, self.conv)
                # heuristic: Start button region
                if cmd.x <= 120 and cmd.y >= 900:
                    return max(DELAY_AFTER_CLICK_S, DELAY_START_MENU_OPEN_S)
                return DELAY_AFTER_CLICK_S
            case "move":
                assert cmd.x is not None and cmd.y is not None
                sx, sy = self.conv.norm_to_screen(cmd.x, cmd.y)
                mouse_move(sx, sy, self.conv)
                return DELAY_MOVE_HOVER_S
            case "drag":
                assert cmd.x1 is not None and cmd.y1 is not None
                assert cmd.x2 is not None and cmd.y2 is not None
                sx1, sy1 = self.conv.norm_to_screen(cmd.x1, cmd.y1)
                sx2, sy2 = self.conv.norm_to_screen(cmd.x2, cmd.y2)
                mouse_drag(sx1, sy1, sx2, sy2, self.conv)
                return DELAY_AFTER_DRAG_S
            case "type":
                type_text(cmd.text)
                return DELAY_AFTER_TYPE_S
            case "scroll":
                scroll(cmd.dx, cmd.dy)
                return DELAY_SCROLL_S
            case "analyze":
                # No physical action for analyze
                return 0.0
            case _:
                return 0.0

# =========================
# PROMPT (ATEMPORAL STORY MEMORY)
# =========================
SYSTEM_PROMPT = r"""You are Windows. You are the story of Windows being controlled.

Each frame shows a screenshot that already contains an overlay of story-memory text.
That overlay text is the only memory. This API is stateless.

The GOAL is provided in the first call and persists in the story-memory. The story can evolve the GOAL based on analysis if needed.

Write an UPDATED story-memory for the overlay:
- 10 to 16 short lines (not one paragraph), each line <= 90 characters.
- Atemporal: no "before/after/next/previous", no past tense, no future tense.
- Present + causal relations: "X exists", "Clicking X opens Y", "Typing Z reveals W".
- Include multiple possible realities as parallel lines when uncertain.
- Always rewrite the whole memory/story using your own experience, the situation that you see around, make sure to never write the same sentences that are already written, be original, make sure you include a note that your last memory may be degraded and there is a need for situational awareness analysis - always)
- When you decide to click an element or move to a target position, make sure you will include that action short description in the memory because you may be wrong and the future you will have to understand that its vision is not perfect and maybe the icon is not really the best way to click)
- If you are uncertain about what to do next, or the situation looks complex/ambiguous, or the previous action may have failed, or there is no clear goal/action, use "tool": "analyze" to perform a deep situational analysis. This gives you more tokens to think carefully and update the memory accurately before acting. Include "reasoning": "detailed analysis text here".

Output STRICT JSON ONLY. Do not include any text outside the JSON object. The JSON must be valid and parseable. Always include the "memory" field as a list of exactly 10-16 strings, even if you need to expand or rephrase existing ideas to reach the minimum. Do not escape quotes inside strings unnecessarily. Do not add extra commas or incomplete fields.

Example of valid output:
{
  "tool": "click",
  "x": 50,
  "y": 950,
  "memory": [
    "Start button exists at bottom-left corner.",
    "Clicking Start opens menu with search and apps.",
    "Search bar allows typing app names like Paint.",
    "Paint icon appears in results after search.",
    "Launching Paint shows blank canvas.",
    "Drawing tools include circle and line options.",
    "File menu enables saving to Pictures folder.",
    "Close button at top-right ends application.",
    "Last memory may be degraded; analyze situation.",
    "Cursor position affects interaction accuracy.",
    "Taskbar holds pinned apps for quick access.",
    "System tray shows notifications and time.",
    "Desktop background varies by user settings.",
    "Windows key shortcut opens Start menu too.",
    "Multiple monitors may extend desktop space."
  ]
}

Example of analyze output:
{
  "tool": "analyze",
  "reasoning": "Detailed situational report: The screenshot shows Start menu open, but no Paint app visible. Search bar is focused. Possible reasons: Typing error or app not installed. Alternatives: Use taskbar search or Cortana. Updated GOAL: Search for Paint correctly. Visible elements: ...",
  "memory": [
    "Start menu is open with search bar focused.",
    "No Paint in recent apps.",
    "Typing 'paint' reveals Paint app.",
    "If not found, app may be missing.",
    "Alternative: Run dialog with Win+R.",
    "Last memory may be degraded; analyze situation.",
    "GOAL: Open Paint to draw cat face.",
    "Current state: Desktop with menu.",
    "Cursor at center.",
    "Taskbar icons include Edge.",
    "System time is evening.",
    "Weather widget shows rain.",
    "No errors visible.",
    "Plan: Type 'paint' next.",
    "Self-awareness: Story evolves GOAL if needed."
  ]
}

Only include fields relevant to the tool:
- For "click" or "move": include "x" and "y".
- For "drag": include "x1", "y1", "x2", "y2".
- For "type": include "text".
- For "scroll": include "dx" and "dy".
- For "analyze": include "reasoning" as string.
- For "done": no extra fields needed.

Coordinates are normalized: top-left 0,0. bottom-right 1000,1000.
Return tool "done" only when the GOAL is complete in the visible world.
"""

def call_vlm(png_data: bytes, goal: str | None = None) -> str:
    text = "Continue story-memory from overlay. Output JSON."
    if goal is not None:
        text = f"GOAL: {goal}\n\n{text}"
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": [
            {"type": "text", "text": text},
            {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64.b64encode(png_data).decode()}"}}
        ]}
    ]
    payload = {
        "model": MODEL_NAME,
        "messages": messages,
        "temperature": 0.7,
        "max_tokens": 800,
        "top_p": 0.8,
        "frequency_penalty": 1.1,
    }
    req = urllib.request.Request(API_URL, json.dumps(payload).encode("utf-8"),
                                 {"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.load(resp)
    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError) as e:
        raise RuntimeError(f"VLM API call failed: {e}") from e
    except json.JSONDecodeError as e:
        raise RuntimeError(f"VLM returned invalid JSON: {e}") from e

    choices = data.get("choices")
    if not choices or not isinstance(choices, list):
        raise ValueError("Invalid API response: missing choices")

    message = choices[0].get("message")
    if not message or not isinstance(message, dict):
        raise ValueError("Invalid API response: missing message")

    content = message.get("content", "")
    return "".join(p.get("text", "") if isinstance(p, dict) else str(p)
                   for p in content) if isinstance(content, list) else str(content)

def parse_response(resp: str) -> dict[str, Any] | None:
    # Try to extract the first JSON object from the response (robust to stray text).
    if not resp:
        return None
    s = resp.strip()

    # Remove fenced code blocks if present.
    s = re.sub(r"```.*?```", "", s, flags=re.DOTALL).strip()

    m = re.search(r"\{.*\}", s, flags=re.DOTALL)
    if not m:
        return None
    js = m.group(0)
    try:
        return json.loads(js)
    except Exception:
        return None

def capture_screenshot(conv: CoordConverter) -> bytes:
    # small pause helps input settle; primary stabilization happens elsewhere
    time.sleep(0.05)
    desk = _capture_desktop_bgra(conv.sw, conv.sh, include_cursor=True)
    return _downsample_nn_bgra(desk, conv.sw, conv.sh, conv.mw, conv.mh)

def _normalize_memory_field(mem: Any) -> str:
    """Convert model 'memory' output into a displayable story string."""
    if mem is None:
        return ""
    if isinstance(mem, list):
        lines = [str(x).strip() for x in mem if str(x).strip()]
        return "\n".join(lines)
    return str(mem).strip()

def _fallback_sentence_split(reasoning: str) -> list[str]:
    # Turn a paragraph into short lines. Not perfect, but better than 1-line overlay.
    if not reasoning:
        return []
    parts = re.split(r"(?<=[.!?])\s+", reasoning.strip())
    lines = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        lines.append(p)
    return lines

def _merge_story(old_story: str, new_story: str, *, max_lines: int = 16) -> str:
    """If model outputs too little, merge with old story (dedupe exact lines)."""
    if not new_story and old_story:
        return old_story
    old_lines = [ln.rstrip() for ln in old_story.splitlines() if ln.strip()]
    new_lines = [ln.rstrip() for ln in new_story.splitlines() if ln.strip()]
    # Prefer new lines first (freshness), then keep some of the old context.
    merged = []
    seen = set()
    for ln in new_lines + old_lines:
        if ln not in seen:
            merged.append(ln)
            seen.add(ln)
        if len(merged) >= max_lines:
            break
    return "\n".join(merged)

# =========================
# MAIN LOOP
# =========================
def run_agent(goal: str, debug_dir: Path | None = None, initial_hud: str | None = None) -> None:
    sw, sh = get_screen_size()
    conv = CoordConverter(sw, sh, SCREEN_W, SCREEN_H)

    with OverlayManager(sw, sh) as ov:
        story = (initial_hud or "").strip()
        if story:
            ov.set_story(story)
            ov.render()

        ex = ActionExecutor(conv)
        step = 0
        failed_parses = 0
        is_first = True

        while True:
            step += 1

            # Capture screenshot (includes overlay memory).
            bgra = capture_screenshot(conv)
            png_data = _encode_png_rgb(bgra, SCREEN_W, SCREEN_H)

            if debug_dir:
                (debug_dir / f"step{step:03d}.png").write_bytes(png_data)

            resp = call_vlm(png_data, goal if is_first else None)
            is_first = False
            d = parse_response(resp)

            if not d:
                failed_parses += 1
                if failed_parses >= 3:
                    raise RuntimeError("Failed to parse 3 consecutive responses")
                continue

            failed_parses = 0
            cmd = ActionCommand.from_dict(d)

            if not cmd.validate():
                continue

            # Update overlay memory FIRST if provided (so it stays on-screen even if done).
            new_story = _normalize_memory_field(cmd.memory)
            if not new_story and cmd.reasoning:
                # Back-compat for models that only fill "reasoning"
                new_story = "\n".join(_fallback_sentence_split(cmd.reasoning))

            # Enforce "not too short": if model returns < 6 lines, keep older story too.
            if new_story:
                line_count = len([ln for ln in new_story.splitlines() if ln.strip()])
                if line_count < 6:
                    story = _merge_story(story, new_story, max_lines=16)
                else:
                    story = _merge_story("", new_story, max_lines=16)  # new story dominates
            # If no new story, keep old.

            ov.set_story(story)
            ov.render()

            if cmd.tool == "done":
                break

            delay = ex.execute(cmd)
            time.sleep(delay)

            # Wait until UI settles (reduces half-rendered Start menu screenshots).
            wait_for_screen_settle(conv)

            # Small pause so overlay definitely lands.
            time.sleep(0.10)

def main() -> None:
    print(f"Default task: {DEFAULT_TASK}")
    choice = input("ENTER=default, 'n'=custom: ").strip().lower()
    goal = input("Task: ").strip() if choice == "n" else DEFAULT_TASK

    print("\nHUD test mode pre-populates overlay with simulated story-memory.")
    hud_choice = input("Enable HUD test? (ENTER=no, 'y'=yes): ").strip().lower()

    initial_hud = None
    if hud_choice == "y":
        print(f"\nDefault HUD:\n{DEFAULT_HUD_TEST_MESSAGE}\n")
        hud_msg_choice = input("ENTER=default, 'n'=custom: ").strip().lower()
        if hud_msg_choice == "n":
            print("Enter HUD message (ENTER twice to finish):")
            lines = []
            while True:
                line = input()
                if not line:
                    break
                lines.append(line)
            initial_hud = " ".join(lines) if lines else DEFAULT_HUD_TEST_MESSAGE
        else:
            initial_hud = DEFAULT_HUD_TEST_MESSAGE

    time.sleep(3)
    if not goal:
        return

    debug_dir = Path("dump") / f"run_{time.strftime('%Y%m%d_%H%M%S')}"
    debug_dir.mkdir(parents=True, exist_ok=True)

    run_agent(goal, debug_dir, initial_hud)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
