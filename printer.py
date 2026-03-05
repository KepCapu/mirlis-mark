# printer.py
from __future__ import annotations

import os
import win32print
from typing import Union, Optional

try:
    from PIL import Image, ImageDraw, ImageFont
except Exception:
    Image = None
    ImageDraw = None
    ImageFont = None


# ================================
# RAW печать
# ================================

def _to_crlf_bytes(data: Union[str, bytes]) -> bytes:
    if isinstance(data, bytes):
        return data
    s = data.replace("\r\n", "\n").replace("\n", "\r\n")
    return s.encode("ascii", errors="ignore")


def print_raw(printer_name: str, data: Union[str, bytes]) -> int:
    """Отправка RAW на принтер. data: bytes — без изменений; str — CRLF + ASCII."""
    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        win32print.StartDocPrinter(hPrinter, 1, ("MirlisMarkLabel", None, "RAW"))
        try:
            win32print.StartPagePrinter(hPrinter)

            raw_bytes = data if isinstance(data, bytes) else _to_crlf_bytes(data)
            written = win32print.WritePrinter(hPrinter, raw_bytes)

            win32print.EndPagePrinter(hPrinter)
            return written

        finally:
            win32print.EndDocPrinter(hPrinter)
    finally:
        win32print.ClosePrinter(hPrinter)


# ================================
# Поиск шрифта Windows
# ================================

def _find_windows_font() -> Optional[str]:
    fonts_dir = r"C:\Windows\Fonts"

    candidates = [
        os.path.join(fonts_dir, "segoeui.ttf"),
        os.path.join(fonts_dir, "arial.ttf"),
        os.path.join(fonts_dir, "calibri.ttf"),
        os.path.join(fonts_dir, "tahoma.ttf"),
    ]

    for f in candidates:
        if os.path.exists(f):
            return f

    return None


# ================================
# mm → px
# ================================

def _mm_to_px(mm: float, dpi: int) -> int:
    return int(round(mm / 25.4 * dpi))


# ================================
# перенос текста
# ================================

def _wrap_text(draw, text, font, max_width):

    lines = []

    for raw in text.splitlines():

        words = raw.split(" ")

        if not words:
            lines.append("")
            continue

        cur = words[0]

        for w in words[1:]:

            test = cur + " " + w

            bbox = draw.textbbox((0, 0), test, font=font)

            if bbox[2] - bbox[0] <= max_width:
                cur = test
            else:
                lines.append(cur)
                cur = w

        lines.append(cur)

    return "\n".join(lines)


# ================================
# bitmap → TSPL payload
# ================================

def _image_to_tspl_bitmap_payload(img_1bit):

    w, h = img_1bit.size
    width_bytes = (w + 7) // 8

    px = img_1bit.load()

    out = bytearray()

    for y in range(h):

        for xb in range(width_bytes):

            b = 0

            for bit in range(8):

                x = xb * 8 + bit

                if x >= w:
                    continue

                # PIL mode "1": 0 = чёрный, 255 = белый. TSPL BITMAP: бит 1 = чёрная точка.
                # На 4B-2054L бит 1 = белый (не печатать), 0 = чёрный — ставим бит для белого.
                is_white = (px[x, y] != 0)
                if is_white:
                    b |= (1 << (7 - bit))

            out.append(b)

    return width_bytes, h, bytes(out)


# ================================
# создание TSPL
# ================================

def build_bitmap_tspl(
    text: str,
    label_w_mm: float = 58,
    label_h_mm: float = 80,
    dpi: int = 203,
    padding_mm: float = 0.2,
    margin_mm: float = 0,
    font_size_pt: int = 28,
    density: int = 10,
    speed: int = 4,
    threshold: int = 200,
):

    if Image is None:
        raise RuntimeError("Pillow не установлен. python -m pip install pillow")

    pad_px = _mm_to_px(padding_mm, dpi)

    W_full = _mm_to_px(label_w_mm, dpi)
    H_full = _mm_to_px(label_h_mm, dpi)

    # полный размер этикетки, BITMAP ставим от (0,0) без сдвигов
    img = Image.new("L", (W_full, H_full), 255)

    draw = ImageDraw.Draw(img)

    font_path = _find_windows_font()

    if not font_path:
        raise RuntimeError("Не найден системный шрифт")

    font = ImageFont.truetype(font_path, font_size_pt)

    max_w = W_full - 2 * pad_px

    wrapped = _wrap_text(draw, text, font, max_w)

    draw.multiline_text(
        (pad_px, pad_px),
        wrapped,
        font=font,
        fill=0,
        spacing=6,
        align="left"
    )

    thr = max(0, min(255, threshold))
    # Бинаризация без dithering: фон белый (255), текст чёрный (0). Порог убирает грязь.
    img_1 = img.point(lambda p: 255 if p > thr else 0, mode="1")

    width_bytes, height, raster = _image_to_tspl_bitmap_payload(img_1)

    header = (
        f"SIZE {label_w_mm} mm, {label_h_mm} mm\r\n"
        f"GAP 2 mm, 0 mm\r\n"
        f"SPEED {speed}\r\n"
        f"DENSITY {density}\r\n"
        f"DIRECTION 1\r\n"
        f"REFERENCE 0,0\r\n"
        f"CLS\r\n"
        f"BITMAP 0,0,{width_bytes},{height},0,"
    ).encode("ascii")

    tail = b"\r\nPRINT 1\r\n"

    return header + raster + tail


# ================================
# публичная функция печати
# ================================

def print_text_as_bitmap_tspl(
    printer_name: str,
    text: str,
    label_w_mm: float = 58,
    label_h_mm: float = 80,
    padding_mm: float = 0.2,
    font_size_pt: int = 28,
):

    tspl = build_bitmap_tspl(
        text,
        label_w_mm=label_w_mm,
        label_h_mm=label_h_mm,
        padding_mm=padding_mm,
        font_size_pt=font_size_pt,
        threshold=200,
    )

    return print_raw(printer_name, tspl)