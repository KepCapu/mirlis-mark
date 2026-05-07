from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
import math
from datetime import datetime, date
from statistics_data import PrintRecord, normalize_stat_key, compute_shift_totals

from PyQt5.QtCore import Qt, QRect, QSize, QPointF
from PyQt5.QtWidgets import (
    QWidget,
    QDialog,
    QFrame,
    QLabel,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QPushButton,
    QMessageBox,
    QCheckBox,
    QRadioButton,
    QButtonGroup,
    QToolButton,
    QScrollArea,
)
from PyQt5.QtGui import QPainter, QLinearGradient, QRadialGradient, QPixmap
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog
from PyQt5.QtGui import QColor, QFont, QFontMetrics, QPen, QIcon


# Keep styling consistent with statistics UI, but avoid importing statistics_page.py (cycle).
_STATS_FONT_FAMILY = '"Inter","Segoe UI","Manrope","Arial",sans-serif'
_C_TITLE = "#1E2F45"
_C_TEXT = "#24364D"
_C_SUB = "#6B7C93"
_C_MUTED = "#7C8CA3"

_PAGE_BG = "#f1f5f9"
_PAGE_BORDER = "#e5e7eb"
_SECTION_BORDER = "#eef2f7"

# Print/preview pie charts: larger diameter vs previous defaults, thin dividers + outline.
_PIE_PRINT_DIAMETER_SCALE = 1.50
# Extra vertical padding inside the pie widget reserved for external percent labels
# (above and below the circle), so callouts don't hit the widget bounds.
# We intentionally reserve MORE space on top: small-slice labels tend to cluster there.
_PIE_EXTERNAL_LABEL_TOP_VPAD = 72
_PIE_EXTERNAL_LABEL_BOTTOM_VPAD = 60

# Pie external percent labels (callouts). Very small slices create visual noise,
# but they remain present in the pie and in the legend list.
_MIN_EXTERNAL_LABEL_PERCENT = 1.0

# Shift the circle slightly upwards inside its content area.
# Keeps top reserve for callouts but reduces excessive empty space above the pie.
_PIE_CIRCLE_VERTICAL_SHIFT_PX = 180

# Unified series palette (copied from statistics_page.py to match dashboard coloring logic).
_SERIES_PALETTE_HEX = [
    "#3B66C3", "#2FAE7A", "#C35A6B", "#2F97AE", "#B88A2D",
    "#7A4EC3", "#2FAE4B", "#C35A2F", "#2F6AAE", "#6FAE2F",
    "#C34EB6", "#2FAEA8", "#C37D2F", "#4A2FAE", "#3EAE2F",
    "#C33D7D", "#2F7EAE", "#9EAE2F", "#AE2FA6", "#2FAE90",
    "#C34A3B", "#2F56AE", "#57AE2F", "#C32F92", "#2F9E9E",
    "#C3A23B", "#8C2FAE", "#2FAE60", "#C32F4A", "#2F86AE",
    "#86AE2F", "#AE2F8C", "#2FAEAA", "#C3652F", "#2F3EAE",
    "#4AAE2F", "#C32F6A", "#2F9AAE", "#B3AE2F", "#6A2FAE",
    "#2FAE73", "#C33B3B", "#2F72AE", "#7EAE2F", "#C32FAE",
    "#2FAE9E", "#C38A3B", "#5A2FAE", "#2FAE3E", "#C33B86",
]

# Increase "document" text in print preview (headings, section titles, tables)
# without affecting fonts inside chart renderers.
_DOC_FONT_SCALE = 1.25  # ~+25%


def _doc_pt(size_pt: int) -> int:
    return max(1, int(round(float(size_pt) * float(_DOC_FONT_SCALE))))


def _doc_px(size_px: int) -> int:
    return max(1, int(round(float(size_px) * float(_DOC_FONT_SCALE))))


def _font_doc(size: int, weight: int = 600) -> QFont:
    return _font(_doc_pt(int(size)), weight)


# Slightly increase text inside bar charts (axes/value labels) for print/preview readability
# without affecting the surrounding "document" UI text.
_BAR_CHART_FONT_SCALE = 1.10  # ~+10%


def _bar_pt(size_pt: int) -> int:
    return max(1, int(round(float(size_pt) * float(_BAR_CHART_FONT_SCALE))))


def _font_bar(size: int, weight: int = 600) -> QFont:
    return _font(_bar_pt(int(size)), weight)


def _font(size: int, weight: int = 600) -> QFont:
    f = QFont()
    # Prefer first family in the list.
    fam = _STATS_FONT_FAMILY.replace('"', "").split(",")[0].strip()
    if fam:
        f.setFamily(fam)
    f.setPointSize(int(size))
    f.setWeight(int(weight))
    return f


def _label(text: str, *, size: int, weight: int, color: str, wrap: bool = False) -> QLabel:
    lb = QLabel(text)
    lb.setWordWrap(bool(wrap))
    lb.setStyleSheet("background: transparent; border: none;")
    lb.setFont(_font_doc(size, weight))
    try:
        pal = lb.palette()
        pal.setColor(lb.foregroundRole(), QColor(color))
        lb.setPalette(pal)
    except Exception:
        pass
    lb.setAttribute(Qt.WA_TranslucentBackground, True)
    return lb

@dataclass(frozen=True)
class PrintReportOptions:
    period_label: str
    period_subtitle: str
    selected_blocks: tuple[str, ...]
    format_label: str
    orientation_label: str
    # Per-section chart representation modes for top blocks.
    # key in {"top_products","top_staff","workshops"}, modes in {"bar","pie"}.
    chart_modes: tuple[tuple[str, tuple[str, ...]], ...] = ()


def _get_chart_modes(options: PrintReportOptions, key: str) -> set[str]:
    try:
        for k, modes in (options.chart_modes or ()):
            if k == key:
                return {str(m) for m in (modes or ())}
    except Exception:
        return set()
    return set()


def _set_chart_modes(options: PrintReportOptions, key: str, modes: set[str]) -> PrintReportOptions:
    clean = tuple(sorted({m for m in modes if m in ("bar", "pie")}))
    items = [(k, tuple(v)) for (k, v) in (options.chart_modes or ())]
    found = False
    for i, (k, _v) in enumerate(items):
        if k == key:
            items[i] = (k, clean)
            found = True
            break
    if not found:
        items.append((key, clean))
    items.sort(key=lambda x: x[0])
    return PrintReportOptions(
        period_label=options.period_label,
        period_subtitle=options.period_subtitle,
        selected_blocks=options.selected_blocks,
        format_label=options.format_label,
        orientation_label=options.orientation_label,
        chart_modes=tuple((k, tuple(v)) for (k, v) in items),
    )


def _round_to_nearest_5(value: float) -> int:
    return int(round(float(value) / 5.0) * 5.0)


def _print_time_chart_y_scale_max(all_values: list[int]) -> int:
    """
    Shared Y-axis top for all «Динамика по дням» print chunks (full period, not per chunk):
    global_max * 1.20, then nearest multiple of 5; never below the data maximum.
    """
    mx = max((int(v) for v in (all_values or [])), default=0)
    if mx <= 0:
        return 5
    scaled = float(mx) * 1.20
    cap = _round_to_nearest_5(scaled)
    if cap < mx:
        cap = int(math.ceil(float(mx) / 5.0) * 5.0)
    return max(cap, 5)


# Печать «Динамика по дням» / гистограмма по времени: не более 30 точек на одном chart (и тот же размер для layout_bar_count).
_PRINT_TIME_SERIES_MAX_POINTS_PER_CHUNK = 30


def _pie_print_reference_legend_row_count() -> int:
    """
    Число строк 2-колоночной легенды под pie, от которого считается «эталонная» высота блока круга.
    Совпадает с max_rows в estimate_pie_legend_capacity: тот же pie, что при заполненной первой странице (Топ продуктов).
    """
    legend_font = _font_bar(8, 600)
    fm_cap = QFontMetrics(legend_font)
    row_h_cap = max(22, int(fm_cap.height()) + 8)
    gap_v_cap = 10
    min_pie_side_cap = int(round(220 * _PIE_PRINT_DIAMETER_SCALE))
    pie_h_cap = (
        int(round(460 * _PIE_PRINT_DIAMETER_SCALE))
        + int(_PIE_EXTERNAL_LABEL_TOP_VPAD)
        + int(_PIE_EXTERNAL_LABEL_BOTTOM_VPAD)
    )
    inner_h_cap = max(1, pie_h_cap - 28)
    pie_area_h_cap = max(
        1,
        inner_h_cap
        - gap_v_cap
        - min_pie_side_cap
        - int(_PIE_EXTERNAL_LABEL_TOP_VPAD)
        - int(_PIE_EXTERNAL_LABEL_BOTTOM_VPAD),
    )
    return max(1, int(pie_area_h_cap // row_h_cap))


class _MiniBarChart(QWidget):
    """Simple bar chart for report preview/printing (no dependency on dashboard widgets)."""

    def __init__(
        self,
        title: str,
        labels: list[str],
        values: list[int],
        *,
        horizontal: bool = False,
        bar_color: QColor | None = None,
        bar_colors: list[QColor] | None = None,
        grid_color: QColor | None = None,
        warm_gradient: bool = False,
        value_scale_max: int | None = None,
        layout_bar_count: int | None = None,
        show_y_axis_guides: bool = False,
        parent=None,
    ):
        super().__init__(parent)
        self._title = title
        self._labels = list(labels or [])
        self._values = [int(v) for v in (values or [])]
        self._horizontal = bool(horizontal)
        self._bar_color = bar_color or QColor("#93c5fd")
        self._bar_colors = list(bar_colors) if bar_colors else None
        self._grid_color = grid_color or QColor("#eef2f7")
        self._warm_gradient = bool(warm_gradient)
        self._value_scale_max = None if value_scale_max is None else max(1, int(value_scale_max))
        self._layout_bar_count = None if layout_bar_count is None else max(1, int(layout_bar_count))
        self._show_y_axis_guides = bool(show_y_axis_guides)
        # Keep a reasonable floor but allow content-driven height via sizeHint().
        self.setMinimumHeight(180 if not horizontal else 120)
        self.setSizePolicy(self.sizePolicy().Expanding, self.sizePolicy().Fixed)

    def sizeHint(self) -> QSize:  # type: ignore[override]
        pad = 8
        top_extra = 22 if (self._title or "").strip() else 0
        if not self._horizontal:
            h = 190
            return QSize(640, h)
        n = min(len(self._labels or []), len(self._values or []))
        n = max(1, int(n))
        # Horizontal bars: content-driven height to avoid large empty gaps.
        # paintEvent uses row_h=22 and inner margins add ~12px.
        row_h = 22
        h = pad * 2 + top_extra + 12 + (n * row_h) + 8
        return QSize(640, int(h))

    def paintEvent(self, _ev):  # type: ignore[override]
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)
        p.fillRect(self.rect(), Qt.white)

        pad = 8
        r = self.rect().adjusted(pad, pad, -pad, -pad)

        top_extra = 0
        if (self._title or "").strip():
            p.setFont(_font_bar(10, 750))
            p.setPen(QColor(_C_TEXT))
            p.drawText(r.left(), r.top(), r.width(), 18, Qt.AlignLeft | Qt.AlignVCenter, self._title)
            top_extra = 22

        # Document-style plot area: no "card" frame, just a clean plotting rect.
        chart = QRect(r.left(), r.top() + top_extra, r.width(), r.height() - top_extra)

        vals = self._values[:]
        labs = self._labels[:]
        if not vals or not labs:
            p.setPen(QColor(_C_MUTED))
            p.drawText(chart, Qt.AlignCenter, "Нет данных")
            p.end()
            return

        n = min(len(vals), len(labs))
        vals = vals[:n]
        labs = labs[:n]
        data_max = max(vals) if vals else 0
        y_scale = max(1, int(self._value_scale_max)) if self._value_scale_max is not None else max(1, data_max)
        y_scale = max(y_scale, data_max, 1)

        inner = chart.adjusted(4, 6, -4, -6)
        p.setPen(Qt.NoPen)
        bar_color_default = QColor(self._bar_color)

        p.setFont(_font_bar(8, 600))
        fm = QFontMetrics(p.font())
        # A bit bolder for print readability
        value_font = _font_bar(8, 900)
        fm_val = QFontMetrics(value_font)

        if not self._horizontal:
            gap = 8
            # Reserve explicit zones:
            # - top for value labels above bars
            # - bottom for X labels
            top_value_h = fm_val.height() + 6

            def _format_x_label(raw: str) -> str:
                """
                Report X-axis labels:
                - keep full 'dd.mm' (with leading zeros) for day buckets
                - keep other label kinds intact (hours '00', week ranges 'dd.mm–dd.mm', etc.)
                """
                s = (raw or "").strip()
                return s

            def _is_day_label(raw: str) -> bool:
                s = (raw or "").strip()
                return bool(len(s) == 5 and s[2] == "." and s[:2].isdigit() and s[3:].isdigit())

            # Special, predictable strategy for "Динамика по дням":
            # - show label under EACH bar (requirement for print)
            # Additionally: keep formatting dd.mm -> day number.
            is_by_days = (n > 0) and all(_is_day_label(str(x)) for x in labs)
            x_font = _font_bar(8, 600)
            if is_by_days and n >= 20:
                # More labels -> slightly smaller font to keep under-bar dates readable on A4.
                x_font = _font_bar(7, 650)
            x_fm = QFontMetrics(x_font)
            x_label_h = max(18, int(x_fm.height()) + 8)
            plot_frame = inner.adjusted(0, top_value_h, 0, -(x_label_h + 6))

            axis_w = 0
            tick_font = _font_bar(7, 600)
            fm_tick = QFontMetrics(tick_font)
            y_ticks: list[int] = []
            if self._show_y_axis_guides:
                y_ticks = [int(round(y_scale * i / 5)) for i in range(6)]
                y_ticks[-1] = y_scale
                y_ticks = sorted(set(y_ticks))
                axis_w = max(int(fm_tick.horizontalAdvance(str(t))) for t in y_ticks) + 10
                axis_w = max(axis_w, 36)

            plot_bars = plot_frame.adjusted(axis_w, 0, -axis_w, 0)

            # --- Adaptive bar density for report preview/print ---
            # Some sections pass layout_bar_count (e.g. 30) to keep stable sizing across chunks.
            # For short periods this can make bars too thin and leave a large empty area on the right.
            # When actual bar count is small, fill the available width by treating it as an n-cell grid.
            n_layout = max(n, int(self._layout_bar_count)) if self._layout_bar_count else n
            n_layout = max(1, n_layout)

            fill_to_width = (n_layout > n) and (n <= 14)
            if fill_to_width:
                cell_w = float(max(1, plot_bars.width())) / float(max(1, n))
                # Keep bars pleasantly wide but not "full cell" (leave some breathing room).
                bw = max(10, int(cell_w * 0.78))
                bw = min(bw, int(cell_w) - 2) if int(cell_w) >= 12 else bw
                gap = max(4, int(max(0.0, cell_w - float(bw))))
            else:
                # Classic layout: honor n_layout (can be > n to keep chunk sizing stable).
                # Must be safe for any n (including 30) and never reference bw before computing it.
                bw = int((plot_bars.width() - gap * (n_layout - 1)) / n_layout) if n_layout > 1 else int(plot_bars.width())
                bw = max(4, bw)
                cell_w = float(max(1, bw + gap))

            plot_h = max(1, plot_bars.height() - 4)
            baseline_y = plot_bars.bottom()

            if self._show_y_axis_guides and y_ticks:
                for tv in y_ticks:
                    yy = baseline_y - int((float(tv) / float(y_scale)) * float(plot_h))
                    if tv == 0:
                        p.setPen(QPen(QColor(self._grid_color), 1))
                        p.drawLine(plot_bars.left(), yy, plot_bars.right(), yy)
                    else:
                        p.setPen(QPen(QColor("#dce3eb"), 1, Qt.DotLine))
                        p.drawLine(plot_bars.left(), yy, plot_bars.right(), yy)
                p.setPen(QColor(_C_SUB))
                p.setFont(tick_font)
                for tv in y_ticks:
                    yy = baseline_y - int((float(tv) / float(y_scale)) * float(plot_h))
                    txt = str(tv)
                    th = fm_tick.height()
                    p.drawText(
                        QRect(plot_frame.left(), yy - th // 2, axis_w - 4, th),
                        Qt.AlignRight | Qt.AlignVCenter,
                        txt,
                    )
                    p.drawText(
                        QRect(plot_frame.right() - axis_w + 4, yy - th // 2, axis_w - 4, th),
                        Qt.AlignLeft | Qt.AlignVCenter,
                        txt,
                    )
                p.setPen(Qt.NoPen)
            else:
                p.setPen(QPen(QColor(self._grid_color), 1))
                p.drawLine(plot_bars.left(), baseline_y, plot_bars.right(), baseline_y)
                p.setPen(Qt.NoPen)

            # Compute a stable show_every once (based on real bar step and label text width).
            step_px = float(max(1, cell_w))
            max_lab_w = 0
            for raw in labs:
                t = _format_x_label(str(raw))
                if not t:
                    continue
                max_lab_w = max(max_lab_w, int(x_fm.horizontalAdvance(t)))
            min_cell_w = float(max_lab_w + 6)  # padding inside the cell
            width_based_every = max(1, int(math.ceil(min_cell_w / step_px)))

            if is_by_days:
                show_every = 1
            else:
                show_every = width_based_every
                # Guardrails for very dense charts (keeps labels readable on A4).
                if n > 16:
                    show_every = max(show_every, 2)
                if n > 28:
                    show_every = max(show_every, 3)
            for i in range(n):
                v = vals[i]
                bh = int((float(v) / float(y_scale)) * float(plot_h))
                if fill_to_width:
                    x = int(plot_bars.left() + float(i) * cell_w + (cell_w - float(bw)) * 0.5)
                else:
                    x = plot_bars.left() + i * (bw + gap)
                y = baseline_y - bh
                # Color logic aligned to dashboard:
                # - warm gradient (week/month/custom) for time chart
                # - per-bar palette when bar_colors provided
                # - single bar_color fallback
                if self._warm_gradient:
                    grad = QLinearGradient(float(x), float(y), float(x), float(y + bh))
                    grad.setColorAt(0.0, QColor("#E85D4A"))
                    grad.setColorAt(0.5, QColor("#F29E4C"))
                    grad.setColorAt(1.0, QColor("#F4C20D"))
                    outline = QColor("#c2410c")
                    outline.setAlpha(180)
                    pen = QPen(outline)
                    pen.setWidthF(1.0)
                    p.setPen(pen)
                    p.setBrush(grad)
                else:
                    fill = QColor(self._bar_colors[i]) if (self._bar_colors and i < len(self._bar_colors)) else QColor(bar_color_default)
                    p.setPen(Qt.NoPen)
                    p.setBrush(fill)
                p.drawRoundedRect(QRect(x, y, bw, bh), 6, 6)
                # X labels: date-aware formatting + stable skip step.
                show_this = (i == 0) or (i == n - 1) or (i % show_every == 0)
                lab = _format_x_label(str(labs[i])) if show_this else ""
                p.setPen(QColor(_C_SUB))
                p.setFont(x_font)
                if fill_to_width:
                    cell_left = int(plot_bars.left() + float(i) * cell_w)
                    cell_right = int(plot_bars.left() + float(i + 1) * cell_w)
                else:
                    cell_left = x - (gap // 2)
                    cell_right = cell_left + (bw + gap)
                if i == 0:
                    cell_left = plot_bars.left()
                if i == n - 1:
                    cell_right = plot_bars.right()
                p.drawText(QRect(int(cell_left), baseline_y + 2, int(cell_right - cell_left), x_label_h), Qt.AlignCenter, lab)
                p.setFont(_font_bar(8, 600))
                # value label (above bar, inside reserved top zone)
                if bh >= 24:
                    p.setPen(QColor(_C_TEXT))
                    p.setFont(value_font)
                    y_txt = max(inner.top(), y - fm_val.height() - 4)
                    p.drawText(QRect(x, int(y_txt), bw, fm_val.height() + 2), Qt.AlignCenter, str(v))
                    p.setFont(_font_bar(8, 600))
                p.setPen(Qt.NoPen)
        else:
            # For print/preview readability we keep a stable row height.
            vmax = max(1, max(vals))
            row_h = 22
            for i in range(n):
                v = vals[i]
                label_w = 160
                right_pad = 72
                bar_max_w = max(10, inner.width() - label_w - right_pad)
                w = int((v / vmax) * bar_max_w)
                y = inner.top() + i * row_h
                p.setPen(QColor(_C_SUB))
                lab = fm.elidedText(str(labs[i]), Qt.ElideRight, label_w - 6)
                p.drawText(QRect(inner.left(), y, label_w, row_h), Qt.AlignVCenter | Qt.AlignLeft, lab)
                p.setPen(Qt.NoPen)
                fill = QColor(self._bar_colors[i]) if (self._bar_colors and i < len(self._bar_colors)) else QColor(bar_color_default)
                p.setBrush(fill)
                bar_x = inner.left() + label_w
                p.drawRoundedRect(QRect(bar_x, y + 4, w, row_h - 8), 6, 6)
                p.setPen(QColor(_C_TEXT))
                # value column aligned to the right with padding
                p.drawText(
                    QRect(inner.right() - right_pad, y, right_pad, row_h),
                    Qt.AlignVCenter | Qt.AlignRight,
                    str(v),
                )
                p.setPen(Qt.NoPen)

        p.end()


def _mini_pie_print_height_for_width(width_px: int, *, label_count: int, legend_max_items: int | None) -> int:
    """
    Высота виджета печатного pie (inner + 28), согласованная с paintEvent и эталоном «Топ продуктов»:
    та же база, что inner_h_cap в _pie_print_reference_legend_row_count (460*scale + callout-pad),
    плюс добавка, если на первой странице легенды больше, чем в эталонной вместимости.
    Не использовать минимальный inner от «первого подходящего» — он ужимал круг до мизерного размера.
    width_px оставлен в сигнатуре для совместимости вызовов; высота по вертикали от ширины не зависит.
    """
    _ = int(width_px)
    legend_font = _font_bar(8, 600)
    fm = QFontMetrics(legend_font)
    row_h = max(22, int(fm.height()) + 8)
    n = max(0, int(label_count))
    max_items = n
    if legend_max_items is not None:
        max_items = min(max_items, max(0, int(legend_max_items)))
    rows = int(math.ceil(float(max_items) / 2.0)) if max_items > 0 else 0
    legend_h = rows * row_h
    top_r = int(_PIE_EXTERNAL_LABEL_TOP_VPAD)
    bot_r = int(_PIE_EXTERNAL_LABEL_BOTTOM_VPAD)

    if n == 0:
        return 120

    pie_h_cap = (
        int(round(460 * _PIE_PRINT_DIAMETER_SCALE))
        + top_r
        + bot_r
    )
    inner_ref = max(1, pie_h_cap - 28)
    ref_legend_h = _pie_print_reference_legend_row_count() * row_h
    inner_h = inner_ref + max(0, legend_h - ref_legend_h)
    return int(inner_h) + 28


class _MiniPieChart(QWidget):
    def __init__(
        self,
        title: str,
        labels: list[str],
        values: list[int],
        *,
        colors: list[QColor] | None = None,
        legend_offset: int = 0,
        legend_max_items: int | None = None,
        parent=None,
    ):
        super().__init__(parent)
        self._title = title
        self._labels = list(labels or [])
        self._values = [int(v) for v in (values or [])]
        self._colors = list(colors) if colors else None
        self._legend_offset = max(0, int(legend_offset))
        self._legend_max_items = None if legend_max_items is None else max(0, int(legend_max_items))
        # Нижняя граница как у эталонного pie до ужатия: иначе layout мог дать слишком низкий виджет.
        _pie_h_cap = (
            int(round(460 * _PIE_PRINT_DIAMETER_SCALE))
            + int(_PIE_EXTERNAL_LABEL_TOP_VPAD)
            + int(_PIE_EXTERNAL_LABEL_BOTTOM_VPAD)
        )
        self.setMinimumHeight(_pie_h_cap)
        self.setSizePolicy(self.sizePolicy().Expanding, self.sizePolicy().Fixed)

    def hasHeightForWidth(self) -> bool:  # type: ignore[override]
        return True

    def heightForWidth(self, w: int) -> int:  # type: ignore[override]
        ww = max(80, int(w))
        return _mini_pie_print_height_for_width(
            ww,
            label_count=len(self._labels),
            legend_max_items=self._legend_max_items,
        )

    def sizeHint(self) -> QSize:  # type: ignore[override]
        base = max(200, self.width()) if self.width() > 0 else 636
        h = self.heightForWidth(base)
        return QSize(int(base), int(h))

    def paintEvent(self, _ev):  # type: ignore[override]
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)
        p.fillRect(self.rect(), Qt.white)

        pad = 8
        r = self.rect().adjusted(pad, pad, -pad, -pad)
        top_extra = 0
        if (self._title or "").strip():
            p.setFont(_font(10, 750))
            p.setPen(QColor(_C_TEXT))
            p.drawText(r.left(), r.top(), r.width(), 18, Qt.AlignLeft | Qt.AlignVCenter, self._title)
            top_extra = 22

        chart = QRect(r.left(), r.top() + top_extra, r.width(), r.height() - top_extra)

        vals = self._values[:]
        labs = self._labels[:]
        n = min(len(vals), len(labs))
        vals = vals[:n]
        labs = labs[:n]
        total = sum(max(0, int(v)) for v in vals)
        if n == 0 or total <= 0:
            p.setPen(QColor(_C_MUTED))
            p.drawText(chart, Qt.AlignCenter, "Нет данных")
            p.end()
            return

        inner = chart.adjusted(6, 6, -6, -6)
        # Layout: pie on top, legend UNDER the pie (no right-side legend).
        # Use the same font helper as horizontal bar chart (left labels + right values).
        legend_font = _font_bar(8, 600)
        p.setFont(legend_font)
        fm = QFontMetrics(p.font())
        # Match horizontal bar chart minimum row height (see _MiniBarChart horizontal branch).
        row_h = max(22, int(fm.height()) + 8)
        gap_v = 10

        # Render only requested legend slice (chart itself still uses full data).
        legend_start = min(max(0, int(self._legend_offset)), n)
        legend_available = max(0, n - legend_start)
        max_items = legend_available
        if self._legend_max_items is not None:
            max_items = min(max_items, max(0, int(self._legend_max_items)))
        # Keep pie noticeably large in print/preview.
        min_pie_side = int(round(220 * _PIE_PRINT_DIAMETER_SCALE))
        # 2-column legend under the pie: left column fills first, then right.
        rows = int(math.ceil(float(max_items) / 2.0)) if max_items > 0 else 0
        legend_h = max(0, rows * row_h)
        pie_area_h = inner.height() - legend_h - (gap_v if rows > 0 else 0)
        # Одинаковый диаметр круга для всех top pie: не расширять область круга из‑за короткой легенди
        # (иначе «Топ сотрудников» / цеха получают больший круг, чем «Топ продуктов»).
        if rows > 0:
            ref_legend_h = _pie_print_reference_legend_row_count() * row_h
            pie_pack_h = max(1, min(inner.height() - gap_v - ref_legend_h, pie_area_h))
        else:
            pie_pack_h = max(1, pie_area_h)
        pie_area_rect = QRect(inner.left(), inner.top(), inner.width(), pie_pack_h)

        # Reserve extra top/bottom space for external percentage labels.
        top_label_reserve = int(_PIE_EXTERNAL_LABEL_TOP_VPAD)
        bottom_label_reserve = int(_PIE_EXTERNAL_LABEL_BOTTOM_VPAD)
        if pie_area_h <= (top_label_reserve + bottom_label_reserve + 40):
            # If widget is too short, reduce reserves gracefully.
            rr = max(10, int(pie_area_h * 0.10))
            top_label_reserve = rr
            bottom_label_reserve = rr

        pie_content_rect = pie_area_rect.adjusted(
            0,
            top_label_reserve,
            0,
            -bottom_label_reserve,
        )
        pie_area_h_for_circle = max(1, pie_content_rect.height())
        if rows > 0 and pie_area_h < min_pie_side:
            # In the unlikely case the widget is too short, reduce rows (not order),
            # so pie remains readable.
            max_rows = max(1, int((inner.height() - gap_v - min_pie_side) // row_h))
            rows = min(rows, max_rows)
            legend_h = max(0, rows * row_h)
            pie_area_h = max(1, inner.height() - legend_h - gap_v)
            max_items = min(max_items, rows * 2)

        pie_side = max(min_pie_side, min(pie_content_rect.width(), pie_area_h_for_circle))
        pie_side = min(int(round(pie_side * _PIE_PRINT_DIAMETER_SCALE)), pie_content_rect.width(), pie_area_h_for_circle)
        pie_left = pie_content_rect.left() + max(0, (pie_content_rect.width() - pie_side) // 2)
        pie_top_centered = pie_content_rect.top() + max(0, (pie_area_h_for_circle - pie_side) // 2)
        pie_top = int(pie_top_centered) - int(_PIE_CIRCLE_VERTICAL_SHIFT_PX)
        # Clamp: keep the circle below the top reserve (space for callouts),
        # and never overlap the legend/list area below.
        pie_top_min = int(pie_content_rect.top()) + 6
        pie_top_max = int(pie_content_rect.bottom() - pie_side)
        pie_top = max(pie_top_min, pie_top)
        pie_top = min(pie_top_max, pie_top)
        pie = QRect(pie_left, pie_top, pie_side, pie_side)

        # Легенда сразу под областью круга (pie_pack_h), а не под pie_area_h: иначе при короткой
        # легенде остаётся полоса пустоты (pie_area_h > pie_pack_h из‑за эталонной высоты для диаметра).
        legend_top = inner.top() + pie_pack_h + (gap_v if rows > 0 else 0)
        legend = QRect(inner.left(), legend_top, inner.width(), max(0, legend_h))

        # ---- Soft shadow under the pie (subtle depth) ----
        try:
            shadow = pie.adjusted(int(pie.width() * 0.06), int(pie.height() * 0.14), -int(pie.width() * 0.06), int(pie.height() * 0.02))
            grad = QRadialGradient(shadow.center(), float(max(1, shadow.width())) * 0.55)
            grad.setColorAt(0.0, QColor(15, 23, 42, 38))
            grad.setColorAt(0.55, QColor(15, 23, 42, 18))
            grad.setColorAt(1.0, QColor(15, 23, 42, 0))
            p.save()
            p.setPen(Qt.NoPen)
            p.setBrush(grad)
            p.drawEllipse(shadow)
            p.restore()
        except Exception:
            pass

        start_angle = 90 * 16
        a = start_angle
        slice_specs: list[tuple[int, QColor, int, int, int]] = []
        for i in range(n):
            v = max(0, int(vals[i]))
            if v <= 0:
                continue
            span = int(round((v / total) * 360.0 * 16.0))
            color = QColor(self._colors[i]) if (self._colors and i < len(self._colors)) else QColor(_SERIES_PALETTE_HEX[i % len(_SERIES_PALETTE_HEX)])
            slice_specs.append((span, color, i, v, a))
            a -= span

        # ---- Slices with soft separators (no harsh outlines) ----
        sep = QPen(QColor(255, 255, 255, 185))
        sep.setWidthF(0.9)
        sep.setCosmetic(True)
        for span, color, _i, _v, st in slice_specs:
            p.setPen(sep)
            p.setBrush(color)
            p.drawPie(pie, st, -span)

        R = float(min(pie.width(), pie.height())) / 2.0
        cx = float(pie.center().x())
        cy = float(pie.center().y())

        def _ray_end(angle_16th: int) -> tuple[float, float]:
            qt_deg = float(angle_16th) / 16.0
            rad = math.radians(qt_deg)
            return cx + R * math.cos(rad), cy - R * math.sin(rad)

        # Soft depth/highlight overlay (front-facing, no tilt).
        p.save()
        try:
            p.setClipRegion(p.clipRegion().intersected(pie.adjusted(0, 0, 1, 1)))
            hl = QRadialGradient(
                QPointF(
                    float(pie.center().x()) - float(pie.width()) * 0.20,
                    float(pie.center().y()) - float(pie.height()) * 0.22,
                ),
                float(max(1, pie.width())) * 0.85,
            )
            hl.setColorAt(0.0, QColor(255, 255, 255, 82))
            hl.setColorAt(0.35, QColor(255, 255, 255, 28))
            hl.setColorAt(1.0, QColor(255, 255, 255, 0))
            p.setPen(Qt.NoPen)
            p.setBrush(hl)
            p.drawEllipse(pie)
        except Exception:
            # Cosmetic only: never break chart rendering.
            pass
        finally:
            p.restore()

        out_pen = QPen(QColor(30, 41, 59, 105))
        out_pen.setWidthF(0.85)
        out_pen.setCosmetic(True)
        p.setPen(out_pen)
        p.setBrush(Qt.NoBrush)
        p.drawEllipse(pie)

        # External percentage labels with leader lines (no on-slice text).
        #
        # Core geometry rule:
        # - Use mid-angle direction as the base (always outward, never inward).
        # - Resolve collisions by extending outward (radius) and sideways (x), NOT by bending back in Y.
        pct_font = _font(8, 800)
        p.setFont(pct_font)
        fm_pct = QFontMetrics(pct_font)
        leader_pen = QPen(QColor(15, 23, 42, 115))
        leader_pen.setWidthF(0.75)
        leader_pen.setCosmetic(True)
        text_clearance_px = 8  # gap between line end and percent text

        label_color = QColor("#0f172a")
        r_outer = float(min(pie.width(), pie.height())) / 2.0
        ext_r_base = r_outer + 14.0
        h_len_base = 22.0

        def _label_rect(side: str, x2: float, y: float, txt: str) -> QRect:
            br = fm_pct.boundingRect(str(txt))
            tw = int(max(1, br.width()))
            th = int(max(1, br.height()))
            if side == "right":
                tx = int(round(x2 + 4.0))
                return QRect(tx, int(round(y - th / 2.0)), tw + 6, th + 4)
            tx = int(round(x2 - 4.0))
            return QRect(tx - tw - 6, int(round(y - th / 2.0)), tw + 6, th + 4)

        # Keep labels within a safe vertical band around the pie area (incl. reserves).
        top_band = float(pie_area_rect.top()) + 4.0
        bot_band = float(pie_area_rect.bottom()) - 4.0

        # Build candidates; each label will be placed along its outward ray.
        raw_items: list[dict] = []
        for span, color, _i, v, st in slice_specs:
            pct = (v / total) * 100.0 if total > 0 else 0.0
            if pct < float(_MIN_EXTERNAL_LABEL_PERCENT):
                continue
            txt = f"{pct:.1f}%"
            br = fm_pct.boundingRect(str(txt))
            tw = int(max(1, br.width()))
            th = int(max(1, br.height()))
            mid_qt_deg = (st - span / 2) / 16.0
            ang = math.radians(mid_qt_deg)
            ca = math.cos(ang)
            sa = math.sin(ang)
            side = "right" if ca >= 0 else "left"
            raw_items.append({"ca": ca, "sa": sa, "side": side, "txt": txt, "tw": tw, "th": th, "color": QColor(color)})

        # Place right side first, then left (keeps symmetry and reduces overlaps).
        placed_rects: list[QRect] = []
        label_items: list[dict] = []

        def place_side(side: str) -> None:
            items = [it for it in raw_items if it["side"] == side]
            # Sort by natural Y (from base ext radius) so placement is stable.
            def natural_y(it: dict) -> float:
                return cy - ext_r_base * float(it["sa"])

            items.sort(key=natural_y)
            for it in items:
                ca = float(it["ca"])
                sa = float(it["sa"])
                # Base anchor on the slice edge.
                x0 = cx + r_outer * ca
                y0 = cy - r_outer * sa

                # Try outward placements: increase radius and horizontal extension until free.
                ext_r = ext_r_base
                h_len = h_len_base
                step_r = 8.0
                step_h = 8.0
                for _try in range(26):
                    x1 = cx + ext_r * ca
                    y1 = cy - ext_r * sa
                    # Keep y on the ray direction (no back-bending).
                    y = max(top_band, min(bot_band, y1))
                    x2 = x1 + (h_len if side == "right" else -h_len)
                    rect = _label_rect(side, x2, y, str(it["txt"]))
                    if not any(rect.intersects(r) for r in placed_rects):
                        placed_rects.append(rect)
                        label_items.append(
                            {
                                "side": side,
                                "txt": it["txt"],
                                "tw": it["tw"],
                                "th": it["th"],
                                "color": it["color"],
                                "x0": x0,
                                "y0": y0,
                                "x1": x1,
                                "y1": y1,
                                "x2": x2,
                                "y": y,
                            }
                        )
                        break
                    # Extend further outward in the same direction.
                    ext_r += step_r
                    h_len += step_h

        place_side("right")
        place_side("left")

        p.setBrush(Qt.NoBrush)
        for it in label_items:
            # Draw leader: anchor -> outward ray point -> horizontal to label.
            y = float(it["y"])
            x0i, y0i = int(round(it["x0"])), int(round(it["y0"]))
            x1i, y1i = int(round(it["x1"])), int(round(it["y1"]))
            x2i, yi = int(round(float(it["x2"]))), int(round(y))

            # Compute real text rect first (using QFontMetrics.boundingRect),
            # then stop the leader line BEFORE it with a guaranteed gap.
            rect = _label_rect(str(it["side"]), float(it["x2"]), float(y), str(it["txt"]))
            safe_rect = rect.adjusted(-1, -1, 1, 1)  # hard guarantee: line never touches pixels of text
            if it["side"] == "right":
                line_end_x = int(safe_rect.left() - text_clearance_px)
            else:
                line_end_x = int(safe_rect.right() + text_clearance_px)

            # Leader line in the slice color (same as the corresponding sector).
            pen = QPen(QColor(it.get("color", QColor(15, 23, 42, 115))))
            try:
                pen.setAlpha(150)  # keep it light
            except Exception:
                pass
            pen.setWidthF(0.75)
            pen.setCosmetic(True)
            p.setPen(pen)

            # Draw outward ray segment.
            p.drawLine(x0i, y0i, x1i, y1i)
            # Small vertical correction (if band clamp adjusted y).
            if yi != y1i:
                # Safety: avoid vertical segment crossing the text rect.
                if safe_rect.left() <= x1i <= safe_rect.right():
                    # Clip vertical segment before it would enter safe_rect.
                    if yi > y1i and y1i <= safe_rect.top() <= yi:
                        yi = int(safe_rect.top() - text_clearance_px)
                    elif yi < y1i and yi <= safe_rect.bottom() <= y1i:
                        yi = int(safe_rect.bottom() + text_clearance_px)
                p.drawLine(x1i, y1i, x1i, yi)
            # Horizontal segment stops before text rect.
            hx2 = int(line_end_x)
            # Safety: avoid horizontal segment crossing the text rect.
            if safe_rect.top() <= yi <= safe_rect.bottom():
                if it["side"] == "right" and x1i <= safe_rect.left() <= hx2:
                    hx2 = int(safe_rect.left() - text_clearance_px)
                if it["side"] == "left" and hx2 <= safe_rect.right() <= x1i:
                    hx2 = int(safe_rect.right() + text_clearance_px)
            p.drawLine(x1i, yi, hx2, yi)

            # Draw text after the line (always on top, no overlap).
            p.setPen(label_color)
            if it["side"] == "right":
                p.drawText(rect, Qt.AlignVCenter | Qt.AlignLeft, str(it["txt"]))
            else:
                p.drawText(rect, Qt.AlignVCenter | Qt.AlignRight, str(it["txt"]))

        # Pie slice labels temporarily switch to a heavier font; restore legend font.
        p.setFont(legend_font)

        # legend (under pie, 2 columns): marker + label + value + percent
        col_gap = 16
        col_w = max(10, (legend.width() - col_gap) // 2)
        value_w = max(92, min(132, int(col_w * 0.42)))  # keep value close to name
        marker_w = 12
        text_pad = 6

        def draw_row(item_i: int, col: int, row: int) -> None:
            x0 = legend.left() + col * (col_w + col_gap)
            y0 = legend.top() + row * row_h
            c = QColor(self._colors[item_i]) if (self._colors and item_i < len(self._colors)) else QColor(_SERIES_PALETTE_HEX[item_i % len(_SERIES_PALETTE_HEX)])

            # marker
            p.setPen(Qt.NoPen)
            p.setBrush(c)
            p.drawEllipse(x0, y0 + (row_h // 2) - 4, 8, 8)

            # label (left side of the column)
            p.setPen(QColor(_C_SUB))
            label_x = x0 + marker_w
            label_w = max(10, col_w - marker_w - value_w - text_pad)
            lab_txt = fm.elidedText(str(labs[item_i]), Qt.ElideRight, label_w)
            p.drawText(QRect(label_x, y0, label_w, row_h), Qt.AlignVCenter | Qt.AlignLeft, lab_txt)

            # value + percent (right side of the same column, close to the label)
            p.setPen(QColor(_C_TEXT))
            vv = max(0, int(vals[item_i]))
            pp = (vv / total) * 100.0 if total > 0 else 0.0
            val_txt = f"{vv} ({pp:.1f}%)"
            val_x = x0 + col_w - value_w
            p.drawText(QRect(val_x, y0, value_w, row_h), Qt.AlignVCenter | Qt.AlignRight, val_txt)

        # Fill left column first, then right.
        items_to_draw = min(max_items, n - legend_start)
        for idx in range(items_to_draw):
            item_i = legend_start + idx
            r = idx if idx < rows else None
            if r is not None and r < rows:
                draw_row(item_i, 0, r)
                continue
            r2 = idx - rows
            if 0 <= r2 < rows:
                draw_row(item_i, 1, r2)

        p.end()


def build_report_pages(options: PrintReportOptions, records: list[PrintRecord]) -> list[QWidget]:
    """
    Build A4-like page widgets for preview/print.
    This is the single source of truth for both preview rendering and QPrinter output.
    """
    orient = (options.orientation_label or "").strip()
    landscape = orient == "Альбомная"

    # Screen preview size (px). Aspect matches A4; used as base for print scaling.
    # Slightly upscale for preview readability.
    base_w, base_h = (1123, 794) if landscape else (794, 1123)
    page_w, page_h = int(base_w * 1.12), int(base_h * 1.12)
    margin = 48  # inner page padding (gives charts more width in preview)
    section_gap = 12
    content_w = page_w - margin * 2
    content_h = page_h - margin * 2
    # header height is dynamic (wrap), measure once with the same width
    _measured_body_h: int | None = None

    def mk_page() -> tuple[QFrame, QVBoxLayout]:
        page = QFrame()
        page.setObjectName("A4Page")
        page.setFixedSize(page_w, page_h)
        page.setStyleSheet(
            "#A4Page {"
            "background: #ffffff;"
            f"border: 1px solid {_PAGE_BORDER};"
            "border-radius: 10px;"
            "}"
        )
        lay = QVBoxLayout(page)
        lay.setContentsMargins(margin, margin, margin, margin)
        lay.setSpacing(0)

        header = QWidget()
        header.setStyleSheet("background: transparent; border: none;")
        hl = QHBoxLayout(header)
        hl.setContentsMargins(0, 0, 0, 8)
        hl.setSpacing(10)
        left = QWidget()
        left.setStyleSheet("background: transparent; border: none;")
        ll = QVBoxLayout(left)
        ll.setContentsMargins(0, 0, 0, 0)
        ll.setSpacing(2)
        ll.addWidget(_label("Отчёт по статистике", size=16, weight=900, color=_C_TITLE), 0, Qt.AlignLeft)
        ll.addWidget(
            _label(f"Период: {options.period_label} · {options.period_subtitle}", size=10, weight=650, color=_C_SUB, wrap=True),
            0,
            Qt.AlignLeft,
        )
        hl.addWidget(left, 1)
        # page number placeholder, updated after pagination
        page_no = _label("", size=10, weight=650, color=_C_MUTED)
        page_no.setObjectName("PageNumberLabel")
        hl.addWidget(page_no, 0, Qt.AlignRight | Qt.AlignTop)
        lay.addWidget(header, 0)

        body = QWidget()
        body.setStyleSheet("background: transparent; border: none;")
        bl = QVBoxLayout(body)
        bl.setContentsMargins(0, 0, 0, 0)
        bl.setSpacing(section_gap)
        lay.addWidget(body, 1)

        nonlocal _measured_body_h
        if _measured_body_h is None:
            # Force the header to compute its height with the fixed content width.
            header.setFixedWidth(content_w)
            header.adjustSize()
            # Только sizeHint: height() после adjustSize() часто выше фактической высоты
            # в итоговом layout страницы → hh завышается → body_h занижается → лишние разрывы
            # (pie «Топ сотрудников» под bar при визуально достаточном месте).
            hh = int(header.sizeHint().height())
            # Available height for body content inside the page.
            _measured_body_h = max(1, content_h - hh)
        return page, bl

    # Measure body height early so section builders can estimate per-page capacities.
    _tmp_pg, _tmp_lay = mk_page()
    _body_h_for_chunks = int(_measured_body_h or (content_h - 64))

    def section_frame(title: str) -> tuple[QWidget, QVBoxLayout]:
        # Document-like section: title + thin divider, no "card" container.
        w = QWidget()
        w.setStyleSheet("background: transparent; border: none;")
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(8)
        title_lb = _label(title, size=12, weight=850, color=_C_TEXT)
        title_lb.setObjectName("SectionTitleLabel")
        lay.addWidget(title_lb, 0, Qt.AlignLeft)
        sep = QFrame()
        sep.setObjectName("SectionDivider")
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"color: {_SECTION_BORDER}; background: {_SECTION_BORDER}; max-height: 1px;")
        sep.setFixedHeight(1)
        lay.addWidget(sep, 0)
        w.setFixedWidth(content_w)
        return w, lay

    def kv_table(rows: list[tuple[str, str]]) -> QWidget:
        w = QWidget()
        w.setStyleSheet("background: transparent; border: none;")
        g = QGridLayout(w)
        g.setContentsMargins(0, 0, 0, 0)
        g.setHorizontalSpacing(12)
        g.setVerticalSpacing(6)
        for i, (k, v) in enumerate(rows):
            lk = QLabel(str(k))
            lk.setStyleSheet(
                f"font-family: {_STATS_FONT_FAMILY}; font-size: {_doc_px(12)}px; font-weight: 650;"
                f"color: {_C_SUB}; background: transparent; border: none;"
            )
            lv = QLabel(str(v))
            lv.setStyleSheet(
                f"font-family: {_STATS_FONT_FAMILY}; font-size: {_doc_px(12)}px; font-weight: 800;"
                f"color: {_C_TEXT}; background: transparent; border: none;"
            )
            g.addWidget(lk, i, 0, 1, 1, Qt.AlignLeft)
            g.addWidget(lv, i, 1, 1, 1, Qt.AlignLeft)
        g.setColumnStretch(0, 0)
        g.setColumnStretch(1, 1)
        w.setFixedWidth(content_w - 28)
        return w

    def kv_table_two_columns(rows: list[tuple[str, str]]) -> QWidget:
        """Compact 2-column KV list for dense time-series tables."""
        rows = list(rows or [])
        n = len(rows)
        if n <= 6:
            return kv_table(rows)
        mid = (n + 1) // 2
        left = rows[:mid]
        right = rows[mid:]

        w = QWidget()
        w.setStyleSheet("background: transparent; border: none;")
        g = QGridLayout(w)
        g.setContentsMargins(0, 0, 0, 0)
        g.setHorizontalSpacing(18)
        g.setVerticalSpacing(6)

        def add_row(rr: int, cc: int, k: str, v: str) -> None:
            lk = QLabel(str(k))
            lk.setStyleSheet(
                f"font-family: {_STATS_FONT_FAMILY}; font-size: {_doc_px(12)}px; font-weight: 650;"
                f"color: {_C_SUB}; background: transparent; border: none;"
            )
            lv = QLabel(str(v))
            lv.setStyleSheet(
                f"font-family: {_STATS_FONT_FAMILY}; font-size: {_doc_px(12)}px; font-weight: 800;"
                f"color: {_C_TEXT}; background: transparent; border: none;"
            )
            g.addWidget(lk, rr, cc, 1, 1, Qt.AlignLeft)
            g.addWidget(lv, rr, cc + 1, 1, 1, Qt.AlignRight)

        for i, (k, v) in enumerate(left):
            add_row(i, 0, str(k), str(v))
        for i, (k, v) in enumerate(right):
            add_row(i, 2, str(k), str(v))

        g.setColumnStretch(0, 0)
        g.setColumnStretch(1, 0)
        g.setColumnStretch(2, 0)
        g.setColumnStretch(3, 0)
        w.setFixedWidth(content_w - 28)
        return w

    def is_tables() -> bool:
        return (options.format_label or "").strip() in ("Только значения", "Графики и значения")

    def is_charts() -> bool:
        # "Только графики" removed: charts are shown only in the combined mode.
        return (options.format_label or "").strip() in ("Графики и значения",)

    # ---- Build sections (widgets) ----
    sections: list[QWidget] = []

    blocks = set(options.selected_blocks or ())

    def build_time_series_for_report() -> tuple[str, list[str], list[int], bool]:
        """
        Time series bucketing for print/preview:
        - День: по часам (24)
        - 7/30 дней и пользовательский период: по календарным дням (полный диапазон)
        Печать не сворачивает длинный период в недели — при нехватке места используется пагинация.
        Returns: (section_title, labels, buckets, warm_gradient)
        """
        pl = (options.period_label or "").strip()
        # base records
        if not records:
            return ("Динамика", [], [], False)

        if pl == "День":
            buckets = [0] * 24
            for r in records:
                buckets[int(r.dt.hour)] += int(r.copies)
            labels = [f"{h:02d}" for h in range(24)]
            return ("По часам", labels, buckets, False)

        # Determine FULL date range for day-bucketing (must include zero days).
        # Important: records list is already filtered by the selected period.
        # For fixed presets (7/30 days) we align to the same "now" anchor as the UI filter:
        # otherwise a single printed day would shrink the chart to 1 bar.
        now_d = datetime.now().date()
        max_d = max(r.dt for r in records).date()
        min_d = min(r.dt for r in records).date()

        def _parse_custom_range_from_subtitle() -> tuple[date | None, date | None]:
            # Expected: "dd.mm.yyyy HH:MM — dd.mm.yyyy HH:MM"
            s = (options.period_subtitle or "").strip()
            if "—" not in s:
                return (None, None)
            try:
                a, b = [x.strip() for x in s.split("—", 1)]
                da = datetime.strptime(a, "%d.%m.%Y %H:%M").date()
                db = datetime.strptime(b, "%d.%m.%Y %H:%M").date()
                return (da, db)
            except Exception:
                return (None, None)

        if pl == "7 дней":
            end_d = now_d
            start_d = end_d.fromordinal(end_d.toordinal() - 6)
            title = "Динамика по дням"
        elif pl == "30 дней":
            end_d = now_d
            start_d = end_d.fromordinal(end_d.toordinal() - 29)
            title = "Динамика по дням"
        else:
            # custom "Период": всегда по дням (как в UI); длинный диапазон — через continuation страниц.
            cs, ce = _parse_custom_range_from_subtitle()
            start_d = cs or min_d
            end_d = ce or max_d
            title = "Динамика по дням"

        # day buckets for [start_d..end_d]
        days: list = []
        cur = start_d
        safety = 0
        while cur <= end_d and safety < 4000:
            safety += 1
            days.append(cur)
            cur = cur.fromordinal(cur.toordinal() + 1)
        idx = {d: i for i, d in enumerate(days)}
        buckets = [0] * len(days)
        for r in records:
            d = r.dt.date()
            if d in idx:
                buckets[idx[d]] += int(r.copies)
        labels = [d.strftime("%d.%m") for d in days]
        return (title, labels, buckets, True)

    if "Сводка" in blocks:
        fr, fr_lay = section_frame("Сводка")
        ops = len(records)
        labels_total = sum(int(r.copies) for r in records)
        day_shift_total, night_shift_total = compute_shift_totals(records)
        if is_tables():
            fr_lay.addWidget(
                kv_table(
                    [
                        ("Всего этикеток", str(labels_total)),
                        ("Всего операций", str(ops)),
                        ("Дневная смена", str(day_shift_total)),
                        ("Ночная смена", str(night_shift_total)),
                    ]
                ),
                0,
            )
        if is_charts():
            # For now: keep charts in other sections; summary is numeric.
            pass
        sections.append(fr)

    if "Гистограмма по времени" in blocks:
        time_title, labels, buckets, warm = build_time_series_for_report()
        chunk_pts = _PRINT_TIME_SERIES_MAX_POINTS_PER_CHUNK
        y_scale_max = _print_time_chart_y_scale_max(list(buckets))

        def _make_time_chunk_chart(off: int, ln: int) -> _MiniBarChart:
            ch0 = _MiniBarChart(
                "",
                labels[off : off + ln],
                buckets[off : off + ln],
                horizontal=False,
                bar_color=QColor("#93c5fd"),
                grid_color=QColor("#eef2f7"),
                warm_gradient=warm,
                value_scale_max=y_scale_max,
                layout_bar_count=chunk_pts,
                show_y_axis_guides=True,
            )
            ch0.setMinimumHeight(320 if (options.period_label or "").strip() != "День" else 280)
            ch0.setFixedWidth(content_w - 4)
            return ch0

        if is_tables():
            rows_all = [(labels[i], str(buckets[i])) for i in range(len(labels))]
            # Каждый блок: своя гистограмма чанка + список того же чанка; один Y-max и ширина бара для всех чанков.
            fm_ts = QFontMetrics(_font_doc(12, 800))
            row_h_ts = max(18, int(fm_ts.height()) + 6)

            def _max_time_kv_items_for_grid_rows(max_rows: int) -> int:
                max_rows = max(1, int(max_rows))
                max_k_a = min(6, max_rows)
                max_k_b = 2 * max_rows if (2 * max_rows >= 7) else 0
                return max(max_k_a, max_k_b)

            def _time_table_chunk_capacity(*, header_h: int, chart_h: int) -> int:
                avail = int(_body_h_for_chunks) - int(header_h) - int(chart_h)
                if chart_h > 0:
                    avail -= 8
                avail -= 4
                grid_rows = max(1, avail // row_h_ts)
                return max(1, _max_time_kv_items_for_grid_rows(grid_rows))

            if not rows_all:
                fr0, fr0_lay = section_frame(time_title)
                if is_charts() and labels:
                    fr0_lay.addWidget(_make_time_chunk_chart(0, len(labels)), 0)
                sections.append(fr0)
            else:
                consumed = 0
                part_idx = 0
                while consumed < len(rows_all):
                    rest = len(rows_all) - consumed
                    sec_title = time_title if consumed == 0 else f"{time_title} (продолжение)"
                    fr_ts, fr_ts_lay = section_frame(sec_title)
                    fr_ts.setProperty("FlowKind", "time_series_table")
                    fr_ts.setProperty("FlowKey", "histogram_time")
                    fr_ts.setProperty("FlowPart", int(part_idx))
                    hdr_h = fr_ts.sizeHint().height()
                    if is_charts():
                        # Ровно до 30 дней/точек на пару chart+list; не дробить чанк из‑за высоты списка — перенос страницы.
                        take = min(rest, chunk_pts)
                    else:
                        table_cap = _time_table_chunk_capacity(header_h=hdr_h, chart_h=0)
                        take = min(rest, table_cap)
                    take = max(1, take)
                    if is_charts():
                        fr_ts_lay.addWidget(_make_time_chunk_chart(consumed, take), 0)
                    fr_ts_lay.addWidget(kv_table_two_columns(rows_all[consumed : consumed + take]), 0)
                    sections.append(fr_ts)
                    consumed += take
                    part_idx += 1
        elif is_charts():
            if not labels:
                fr, fr_lay = section_frame(time_title)
                sections.append(fr)
            else:
                consumed = 0
                part_idx = 0
                n_pts = len(labels)
                while consumed < n_pts:
                    rest = n_pts - consumed
                    take = max(1, min(rest, chunk_pts))
                    sec_title = time_title if consumed == 0 else f"{time_title} (продолжение)"
                    fr, fr_lay = section_frame(sec_title)
                    fr.setProperty("FlowKind", "time_series_table")
                    fr.setProperty("FlowKey", "histogram_time")
                    fr.setProperty("FlowPart", int(part_idx))
                    fr_lay.addWidget(_make_time_chunk_chart(consumed, take), 0)
                    sections.append(fr)
                    consumed += take
                    part_idx += 1
        else:
            fr, fr_lay = section_frame(time_title)
            sections.append(fr)

    def add_top(title: str, section_key: str, key_getter, *, horizontal: bool = True) -> None:
        counter = Counter()
        for r in records:
            k = normalize_stat_key(key_getter(r))
            if k:
                counter[k] += int(r.copies)
        # Print/preview must include full list (no hard top-10 cap).
        top = counter.most_common()
        labs_all = [k for k, _ in top]
        vals_all = [int(v) for _, v in top]
        # Modes are per-section toggles: {"bar","pie"}.
        modes = _get_chart_modes(options, section_key)
        if not modes:
            modes = {"bar", "pie"}
        if not labs_all:
            fr, fr_lay = section_frame(title)
            fr_lay.addWidget(_label("Нет данных", size=10, weight=650, color=_C_MUTED), 0, Qt.AlignLeft)
            sections.append(fr)
            return

        def top_legend_table_chunk(offset: int, limit: int) -> QWidget:
            w = QWidget()
            w.setStyleSheet("background: transparent; border: none;")
            g = QGridLayout(w)
            g.setContentsMargins(0, 0, 0, 0)
            g.setHorizontalSpacing(18)
            g.setVerticalSpacing(6)
            legend_font_px = _bar_pt(8)
            total_all = max(1, sum(vals_all))
            chunk_labs = labs_all[offset:offset + limit]
            chunk_vals = vals_all[offset:offset + limit]
            rows = int(math.ceil(float(len(chunk_labs)) / 2.0))

            def add_item(rr: int, cc: int, abs_i: int, name: str, val: int) -> None:
                color = _SERIES_PALETTE_HEX[abs_i % len(_SERIES_PALETTE_HEX)]
                dot = QLabel("●")
                dot.setStyleSheet(
                    f"font-family: {_STATS_FONT_FAMILY}; font-size: {legend_font_px}px; font-weight: 700;"
                    f"color: {color}; background: transparent; border: none;"
                )
                name_lb = QLabel(str(name))
                name_lb.setStyleSheet(
                    f"font-family: {_STATS_FONT_FAMILY}; font-size: {legend_font_px}px; font-weight: 600;"
                    f"color: {_C_SUB}; background: transparent; border: none;"
                )
                name_lb.setWordWrap(False)
                name_lb.setMinimumWidth(10)
                pp = (float(val) / float(total_all)) * 100.0
                val_lb = QLabel(f"{int(val)} ({pp:.1f}%)")
                val_lb.setStyleSheet(
                    f"font-family: {_STATS_FONT_FAMILY}; font-size: {legend_font_px}px; font-weight: 600;"
                    f"color: {_C_TEXT}; background: transparent; border: none;"
                )
                g.addWidget(dot, rr, cc, 1, 1, Qt.AlignLeft | Qt.AlignVCenter)
                g.addWidget(name_lb, rr, cc + 1, 1, 1, Qt.AlignLeft | Qt.AlignVCenter)
                g.addWidget(val_lb, rr, cc + 2, 1, 1, Qt.AlignRight | Qt.AlignVCenter)

            # Stable 2-column split: left gets ceil(n/2), right gets floor(n/2).
            left_count = rows
            for i in range(min(left_count, len(chunk_labs))):
                add_item(i, 0, offset + i, chunk_labs[i], int(chunk_vals[i]))
            for j in range(max(0, len(chunk_labs) - left_count)):
                idx = left_count + j
                add_item(j, 3, offset + idx, chunk_labs[idx], int(chunk_vals[idx]))

            g.setColumnStretch(0, 0)
            g.setColumnStretch(1, 1)
            g.setColumnStretch(2, 0)
            g.setColumnStretch(3, 0)
            g.setColumnStretch(4, 1)
            g.setColumnStretch(5, 0)
            w.setFixedWidth(content_w - 4)
            return w

        total_items = len(labs_all)
        def estimate_pie_legend_capacity(item_count: int) -> int:
            # Capacity is determined by the same geometry rules as _MiniPieChart.paintEvent().
            # This prevents continuation for odd counts (e.g. 19) when all items actually fit.
            max_rows_cap = _pie_print_reference_legend_row_count()
            return min(max(0, int(item_count)), max_rows_cap * 2)

        pie_first_legend_items = estimate_pie_legend_capacity(total_items)
        list_chunk_items = 28
        # ---- Build physical blocks in strict order: all BAR blocks -> PIE -> PIE legend continuation ----
        if is_charts() and ("bar" in modes):
            # Bar blocks (can span multiple pages). Pie must NOT appear until bars are fully rendered.
            if horizontal:
                # Page-driven capacity based on real sizeHints (avoid magic 10/12).
                # Estimate how many rows fit in ONE bar widget on a fresh page.
                # - header section uses fixed fonts/layout -> take its sizeHint
                # - bar widget overhead from sizeHint(1 row) minus one row height
                fr_tmp, fr_tmp_lay = section_frame(title)
                header_h = fr_tmp.sizeHint().height()
                # Dummy bar widget with 1 item to estimate non-row overhead.
                dummy_colors = [QColor(_SERIES_PALETTE_HEX[0])]
                dummy = _MiniBarChart("", ["_"], [1], horizontal=True, bar_colors=dummy_colors, grid_color=QColor("#f3f4f6"))
                dummy_h = dummy.sizeHint().height()
                row_h = 22
                bar_overhead_h = max(0, int(dummy_h) - row_h)
                rows_per_page = max(1, int((_body_h_for_chunks - header_h - bar_overhead_h) // row_h))
            else:
                rows_per_page = total_items
            bar_consumed = 0
            while bar_consumed < total_items:
                sec_title = title if bar_consumed == 0 else f"{title} (продолжение)"
                fr_bar, fr_bar_lay = section_frame(sec_title)
                fr_bar.setProperty("FlowKind", "top_bar")
                fr_bar.setProperty("FlowKey", str(section_key))
                fr_bar.setProperty("FlowPart", int(bar_consumed))
                bar_take = min(rows_per_page, total_items - bar_consumed)
                grid_c = QColor("#eef2f7") if title == "Распределение по цехам" else QColor("#f3f4f6")
                bar_colors_chunk = [
                    QColor(_SERIES_PALETTE_HEX[(bar_consumed + i) % len(_SERIES_PALETTE_HEX)])
                    for i in range(bar_take)
                ]
                ch = _MiniBarChart(
                    "",
                    labs_all[bar_consumed:bar_consumed + bar_take],
                    vals_all[bar_consumed:bar_consumed + bar_take],
                    horizontal=horizontal,
                    bar_colors=bar_colors_chunk,
                    grid_color=grid_c,
                )
                # Content-driven height: keep chunks visually contiguous.
                ch.setMinimumHeight(ch.sizeHint().height())
                ch.setFixedWidth(content_w - 4)
                fr_bar_lay.addWidget(ch, 0)
                sections.append(fr_bar)
                bar_consumed += bar_take

        if is_charts() and ("pie" in modes):
            # Pie block comes strictly after bars. Keep the same color mapping as bars.
            pie_title = title if ("bar" not in modes) else f"{title} (продолжение)"
            fr_pie, fr_pie_lay = section_frame(pie_title)
            bar_colors_full = [QColor(_SERIES_PALETTE_HEX[i % len(_SERIES_PALETTE_HEX)]) for i in range(total_items)]
            pie = _MiniPieChart(
                "",
                labs_all,
                vals_all,
                colors=bar_colors_full,
                legend_offset=0,
                legend_max_items=pie_first_legend_items,
            )
            pie.setFixedWidth(content_w - 4)
            fr_pie_lay.addWidget(pie, 0)
            fr_pie.setProperty("FlowKind", "top_pie")
            fr_pie.setProperty("FlowKey", str(section_key))
            sections.append(fr_pie)

            # Legend/list continuation blocks (text-only) after the pie.
            consumed = min(total_items, pie_first_legend_items)
            while consumed < total_items:
                fr_next, fr_next_lay = section_frame(f"{title} (продолжение)")
                fr_next_lay.addWidget(top_legend_table_chunk(consumed, list_chunk_items), 0)
                sections.append(fr_next)
                consumed += list_chunk_items

        elif is_tables():
            # Tables-only mode: split long lists into continuation blocks.
            fr, fr_lay = section_frame(title)
            first_table_items = list_chunk_items
            fr_lay.addWidget(
                kv_table([(f"{i+1}. {labs_all[i]}", str(vals_all[i])) for i in range(min(first_table_items, total_items))]),
                0,
            )
            sections.append(fr)
            consumed = min(total_items, list_chunk_items)
            while consumed < total_items:
                fr_next, fr_next_lay = section_frame(f"{title} (продолжение)")
                fr_next_lay.addWidget(
                    kv_table(
                        [
                            (f"{consumed + i + 1}. {labs_all[consumed + i]}", str(vals_all[consumed + i]))
                            for i in range(min(list_chunk_items, total_items - consumed))
                        ]
                    ),
                    0,
                )
                sections.append(fr_next)
                consumed += list_chunk_items

    if "Топ продуктов" in blocks:
        add_top("Топ продуктов", "top_products", lambda r: r.product, horizontal=True)
    if "Топ сотрудников" in blocks:
        add_top("Топ сотрудников", "top_staff", lambda r: r.made_by, horizontal=True)
    if "Распределение по цехам" in blocks:
        add_top("Распределение по цехам", "workshops", lambda r: r.workshop, horizontal=True)

    # ---- Paginate by section chunks ----
    pages: list[QWidget] = []
    page, lay = mk_page()
    used = 0

    def section_h(w: QWidget) -> int:
        """Best-effort accurate height for a fixed-width section widget."""
        w.setFixedWidth(content_w)
        try:
            if w.layout() is not None:
                w.layout().activate()
        except Exception:
            pass
        # Prefer layout-driven size hints: in complex nested layouts Qt may under-report
        # widget.sizeHint()/height until layout hint is queried.
        try:
            lay = w.layout()
            if lay is not None:
                lh = int(lay.sizeHint().height())
            else:
                lh = 0
        except Exception:
            lh = 0
        # Секция с pie: не смешивать layout.sizeHint с max(adjustSize, …) — иначе need завышается.
        try:
            if lh > 0 and w.findChild(_MiniPieChart) is not None:
                return lh
        except Exception:
            pass
        try:
            if w.hasHeightForWidth():
                return int(w.heightForWidth(content_w))
        except Exception:
            pass
        w.adjustSize()
        sh = w.sizeHint().height()
        return max(int(sh), int(w.height()), int(w.minimumHeight()), int(lh))

    for s in sections:
        # Compact bar continuation blocks when they appear on the same page:
        # hide repeated "(продолжение)" headers so bars look like one continuous section.
        try:
            kind = str(s.property("FlowKind") or "")
            key = str(s.property("FlowKey") or "")
            part = int(s.property("FlowPart") or 0)
        except Exception:
            kind, key, part = "", "", 0

        # Enforce "one bar widget per page" for a given top section:
        # bar continuation blocks must always start on a new page.
        if kind == "top_bar" and part > 0 and used > 0:
            pages.append(page)
            page, lay = mk_page()
            used = 0

        if used > 0 and kind == "top_bar" and part > 0:
            # If previous placed block on this page is the same bar-flow, hide title+divider.
            # This reduces whitespace and prevents "new section" look.
            try:
                prev_kind = str(page.property("LastFlowKind") or "")
                prev_key = str(page.property("LastFlowKey") or "")
            except Exception:
                prev_kind, prev_key = "", ""
            hide_hdr = (prev_kind == "top_bar" and prev_key == key)
        else:
            hide_hdr = False

        if kind == "top_bar":
            # Apply header visibility (must happen before measuring heights).
            tl = s.findChild(QLabel, "SectionTitleLabel")
            dv = s.findChild(QFrame, "SectionDivider")
            if tl is not None:
                tl.setVisible(not hide_hdr)
            if dv is not None:
                dv.setVisible(not hide_hdr)

        # Pie сразу под гистограммой той же top-секции: заголовок «… (продолжение)» даёт лишнюю
        # высоту в section_h() и провоцирует перенос, хотя круг+легенда помещаются под bar.
        if kind == "top_pie" and used > 0:
            try:
                prev_kind = str(page.property("LastFlowKind") or "")
                prev_key = str(page.property("LastFlowKey") or "")
            except Exception:
                prev_kind, prev_key = "", ""
            hide_pie_hdr = prev_kind == "top_bar" and prev_key == key
            p_tl = s.findChild(QLabel, "SectionTitleLabel")
            p_dv = s.findChild(QFrame, "SectionDivider")
            if p_tl is not None:
                p_tl.setVisible(not hide_pie_hdr)
            if p_dv is not None:
                p_dv.setVisible(not hide_pie_hdr)

        # Time histogram continuation: hide "(продолжение)" if the block fits on the same sheet.
        ts_tl = s.findChild(QLabel, "SectionTitleLabel")
        ts_dv = s.findChild(QFrame, "SectionDivider")
        if kind == "time_series_table" and part > 0 and ts_tl is not None and ts_dv is not None:
            ts_tl.setVisible(used <= 0)
            ts_dv.setVisible(used <= 0)

        h = section_h(s)
        need = h + (section_gap if used > 0 else 0)
        body_h = int(_measured_body_h or (content_h - 64))
        remaining = body_h - used
        # Pie сразу под bar той же секции: сравниваем с небольшим запасом по remaining —
        # иначе погрешность int(sizeHint) у нескольких виджетов подряд даёт need чуть > remaining
        # при том, что в layout блоки помещаются (тот же сценарий, что и завышенный hh).
        if kind == "top_pie":
            try:
                prev_kind = str(page.property("LastFlowKind") or "")
                prev_key = str(page.property("LastFlowKey") or "")
            except Exception:
                prev_kind, prev_key = "", ""
            if prev_kind == "top_bar" and prev_key == key:
                remaining += 20
        # Только нехватка места — иначе блок уезжает на следующий лист и на текущем остаётся
        # большой пустой хвост (типично: pie сразу после короткой bar-секции «Топ сотрудников» / цехов).
        if used > 0 and need > remaining:
            pages.append(page)
            page, lay = mk_page()
            used = 0
            if kind == "time_series_table" and part > 0 and ts_tl is not None and ts_dv is not None:
                ts_tl.setVisible(True)
                ts_dv.setVisible(True)
                h = section_h(s)
                need = h + (section_gap if used > 0 else 0)
            if kind == "top_pie":
                p_tl2 = s.findChild(QLabel, "SectionTitleLabel")
                p_dv2 = s.findChild(QFrame, "SectionDivider")
                if p_tl2 is not None:
                    p_tl2.setVisible(True)
                if p_dv2 is not None:
                    p_dv2.setVisible(True)
                h = section_h(s)
                need = h + (section_gap if used > 0 else 0)
        lay.addWidget(s, 0, Qt.AlignTop)
        used += need

        # Track last placed flow on this page for header-compaction decisions.
        if kind:
            page.setProperty("LastFlowKind", kind)
            page.setProperty("LastFlowKey", key)
        else:
            page.setProperty("LastFlowKind", "")
            page.setProperty("LastFlowKey", "")

    pages.append(page)

    # Fill page numbers now that we know total pages.
    total = len(pages)
    for i, pg in enumerate(pages, start=1):
        lb = pg.findChild(QLabel, "PageNumberLabel")
        if lb is not None:
            lb.setText(f"Стр. {i}/{total}")
    return pages


def print_pages_to_printer(printer: QPrinter, pages: list[QWidget]) -> None:
    if not pages:
        return
    painter = QPainter(printer)
    try:
        if not painter.isActive():
            raise RuntimeError("QPainter не активен для принтера.")
        page_rect = printer.pageRect()
        if page_rect.width() <= 0 or page_rect.height() <= 0:
            page_rect = printer.paperRect()

        for i, pg in enumerate(pages):
            if i > 0:
                printer.newPage()
            w = max(1, pg.width())
            h = max(1, pg.height())
            sx = float(page_rect.width()) / float(w)
            sy = float(page_rect.height()) / float(h)
            scale = min(sx, sy)
            painter.save()
            painter.translate(page_rect.left(), page_rect.top())
            painter.scale(scale, scale)
            pg.render(painter)
            painter.restore()
    finally:
        painter.end()


def run_print_reports_flow(
    *,
    period_label: str,
    period_subtitle: str,
    records: list[PrintRecord],
    parent: QWidget | None = None,
) -> None:
    """
    High-level orchestration: settings -> preview -> back -> settings ...
    Returns only when user closes flow or successfully prints.
    """
    opts = PrintReportOptions(
        period_label=(period_label or "").strip() or "Период",
        period_subtitle=(period_subtitle or "").strip() or ((period_label or "").strip() or "Период"),
        selected_blocks=("Сводка", "Гистограмма по времени", "Топ продуктов", "Топ сотрудников", "Распределение по цехам"),
        format_label="Графики и значения",
        orientation_label="Книжная",
        chart_modes=(
            ("top_products", ("bar", "pie")),
            ("top_staff", ("bar", "pie")),
            ("workshops", ("bar", "pie")),
        ),
    )
    while True:
        dlg = PrintReportsDialog(opts, parent=parent)
        if dlg.exec_() != QDialog.Accepted:
            return
        opts = dlg.selected_options()

        prev = PrintReportsPreviewDialog(
            options=opts,
            period_label=opts.period_label,
            period_subtitle=opts.period_subtitle,
            format_label=opts.format_label,
            orientation_label=opts.orientation_label,
            blocks=list(opts.selected_blocks),
            records=list(records or []),
            parent=parent,
        )
        res = prev.exec_()
        if res == PrintReportsPreviewDialog.RESULT_BACK:
            continue
        if res == PrintReportsPreviewDialog.RESULT_PRINTED:
            return
        return


class PrintReportsDialog(QDialog):
    def __init__(self, options: PrintReportOptions, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setWindowTitle(" ")
        self.setModal(True)
        self.setMinimumSize(720, 520)
        self.resize(760, 560)

        self._options = options
        self._period_label = (options.period_label or "").strip() or "Период"

        self.setStyleSheet("QDialog { background: #f6f7f9; }")
        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(0)

        card = QFrame()
        card.setObjectName("PrintReportsCard")
        card.setStyleSheet(
            "#PrintReportsCard {"
            "background: #ffffff;"
            "border: 1px solid #e5e7eb;"
            "border-radius: 20px;"
            "}"
        )
        card_lay = QVBoxLayout(card)
        card_lay.setContentsMargins(18, 16, 18, 16)
        card_lay.setSpacing(14)

        title = QLabel("Печать отчётов")
        title.setStyleSheet(
            f"font-family: {_STATS_FONT_FAMILY}; font-size: 20px; font-weight: 750; "
            f"color: {_C_TITLE}; background: transparent; border: none;"
        )
        card_lay.addWidget(title, 0, Qt.AlignLeft)

        info = QLabel(f"Период: {self._period_label}")
        info.setWordWrap(True)
        info.setStyleSheet(
            f"font-family: {_STATS_FONT_FAMILY}; font-size: 13px; font-weight: 600; "
            f"color: {_C_SUB}; background: transparent; border: none;"
        )
        card_lay.addWidget(info, 0, Qt.AlignLeft)

        def _section_title(text: str) -> QLabel:
            lb = QLabel(text)
            lb.setStyleSheet(
                f"font-family: {_STATS_FONT_FAMILY}; font-size: 13px; font-weight: 800; "
                f"color: {_C_TEXT}; background: transparent; border: none;"
            )
            return lb

        def _checkbox(text: str, checked: bool = True) -> QCheckBox:
            cb = QCheckBox(text)
            cb.setChecked(bool(checked))
            cb.setStyleSheet(
                "QCheckBox {"
                f"font-family: {_STATS_FONT_FAMILY};"
                "font-size: 14px;"
                "font-weight: 600;"
                f"color: {_C_TEXT};"
                "padding: 6px 0;"
                "}"
                "QCheckBox::indicator { width: 18px; height: 18px; }"
            )
            return cb

        # ---- Content selection ----
        card_lay.addWidget(_section_title("Состав отчёта"), 0, Qt.AlignLeft)
        sel = set(options.selected_blocks or ())
        self.cb_summary = _checkbox("Сводка", ("Сводка" in sel) if sel else True)
        # UI label can differ from internal key stored in PrintReportOptions.selected_blocks.
        self.cb_time_hist = _checkbox("Этикеток за период", ("Гистограмма по времени" in sel) if sel else True)
        self.cb_top_products = _checkbox("Топ продуктов", ("Топ продуктов" in sel) if sel else True)
        self.cb_top_staff = _checkbox("Топ сотрудников", ("Топ сотрудников" in sel) if sel else True)
        self.cb_workshops = _checkbox("Распределение по цехам", ("Распределение по цехам" in sel) if sel else True)

        def _make_chart_mode_toggles(section_key: str) -> tuple[QWidget, QToolButton, QToolButton]:
            w = QWidget()
            w.setStyleSheet("background: transparent; border: none;")
            lay = QHBoxLayout(w)
            lay.setContentsMargins(0, 0, 0, 0)
            lay.setSpacing(6)

            def icon_bar() -> QIcon:
                pm = QPixmap(16, 16)
                pm.fill(Qt.transparent)
                p = QPainter(pm)
                p.setRenderHint(QPainter.Antialiasing, True)
                p.setPen(Qt.NoPen)
                p.setBrush(QColor(_C_SUB))
                p.drawRoundedRect(2, 9, 3, 5, 1.2, 1.2)
                p.drawRoundedRect(7, 6, 3, 8, 1.2, 1.2)
                p.drawRoundedRect(12, 3, 3, 11, 1.2, 1.2)
                p.end()
                return QIcon(pm)

            def icon_pie() -> QIcon:
                pm = QPixmap(16, 16)
                pm.fill(Qt.transparent)
                p = QPainter(pm)
                p.setRenderHint(QPainter.Antialiasing, True)
                p.setPen(Qt.NoPen)
                p.setBrush(QColor(_C_SUB))
                p.drawEllipse(2, 2, 12, 12)
                p.setBrush(QColor(255, 255, 255, 255))
                p.drawPie(2, 2, 12, 12, 90 * 16, -70 * 16)
                p.end()
                return QIcon(pm)

            btn_bar = QToolButton()
            btn_pie = QToolButton()
            for b in (btn_bar, btn_pie):
                b.setCursor(Qt.PointingHandCursor)
                b.setCheckable(True)
                b.setAutoRaise(True)
                b.setIconSize(QSize(18, 18))
                b.setFixedSize(44, 28)
                b.setStyleSheet(
                    "QToolButton { background: transparent; border: 1px solid rgba(148,163,184,0.35); border-radius: 8px; }"
                    "QToolButton:checked { background: rgba(224,231,255,0.8); border: 1px solid rgba(99,102,241,0.45); }"
                    "QToolButton:disabled { opacity: 0.55; }"
                )
            btn_bar.setToolTip("Гистограмма")
            btn_pie.setToolTip("Диаграмма")
            btn_bar.setIcon(icon_bar())
            btn_pie.setIcon(icon_pie())

            modes = _get_chart_modes(options, section_key)
            if not modes:
                modes = {"bar", "pie"}
            btn_bar.setChecked("bar" in modes)
            btn_pie.setChecked("pie" in modes)

            lay.addWidget(btn_bar, 0)
            lay.addWidget(btn_pie, 0)
            lay.addStretch(1)
            return w, btn_bar, btn_pie

        self._mode_top_products_wrap, self._mode_top_products_bar, self._mode_top_products_pie = _make_chart_mode_toggles("top_products")
        self._mode_top_staff_wrap, self._mode_top_staff_bar, self._mode_top_staff_pie = _make_chart_mode_toggles("top_staff")
        self._mode_workshops_wrap, self._mode_workshops_bar, self._mode_workshops_pie = _make_chart_mode_toggles("workshops")

        c_grid = QGridLayout()
        c_grid.setContentsMargins(0, 0, 0, 0)
        c_grid.setHorizontalSpacing(14)
        c_grid.setVerticalSpacing(0)
        # Visual order only (keys/logic preserved in selected_blocks()).
        # Left column:
        c_grid.addWidget(self.cb_top_products, 0, 0)
        c_grid.addWidget(self.cb_top_staff, 1, 0)
        c_grid.addWidget(self.cb_workshops, 2, 0)
        # Toggles column (for left items only):
        c_grid.addWidget(self._mode_top_products_wrap, 0, 1, Qt.AlignLeft | Qt.AlignVCenter)
        c_grid.addWidget(self._mode_top_staff_wrap, 1, 1, Qt.AlignLeft | Qt.AlignVCenter)
        c_grid.addWidget(self._mode_workshops_wrap, 2, 1, Qt.AlignLeft | Qt.AlignVCenter)
        # Right column:
        c_grid.addWidget(self.cb_time_hist, 0, 2)
        c_grid.addWidget(self.cb_summary, 1, 2)
        card_lay.addLayout(c_grid)

        def _sync_modes_enabled():
            self._mode_top_products_bar.setEnabled(self.cb_top_products.isChecked())
            self._mode_top_products_pie.setEnabled(self.cb_top_products.isChecked())
            self._mode_top_staff_bar.setEnabled(self.cb_top_staff.isChecked())
            self._mode_top_staff_pie.setEnabled(self.cb_top_staff.isChecked())
            self._mode_workshops_bar.setEnabled(self.cb_workshops.isChecked())
            self._mode_workshops_pie.setEnabled(self.cb_workshops.isChecked())

            # ensure at least one mode if enabled
            if self.cb_top_products.isChecked() and (not self._mode_top_products_bar.isChecked()) and (not self._mode_top_products_pie.isChecked()):
                self._mode_top_products_bar.setChecked(True)
            if self.cb_top_staff.isChecked() and (not self._mode_top_staff_bar.isChecked()) and (not self._mode_top_staff_pie.isChecked()):
                self._mode_top_staff_bar.setChecked(True)
            if self.cb_workshops.isChecked() and (not self._mode_workshops_bar.isChecked()) and (not self._mode_workshops_pie.isChecked()):
                self._mode_workshops_bar.setChecked(True)

        def _guard_modes(btn_bar: QToolButton, btn_pie: QToolButton):
            # Called on toggle to ensure at least one remains checked.
            if (not btn_bar.isChecked()) and (not btn_pie.isChecked()):
                btn_bar.setChecked(True)

        self.cb_top_products.toggled.connect(lambda _v: _sync_modes_enabled())
        self.cb_top_staff.toggled.connect(lambda _v: _sync_modes_enabled())
        self.cb_workshops.toggled.connect(lambda _v: _sync_modes_enabled())
        self._mode_top_products_bar.toggled.connect(lambda _v: _guard_modes(self._mode_top_products_bar, self._mode_top_products_pie))
        self._mode_top_products_pie.toggled.connect(lambda _v: _guard_modes(self._mode_top_products_bar, self._mode_top_products_pie))
        self._mode_top_staff_bar.toggled.connect(lambda _v: _guard_modes(self._mode_top_staff_bar, self._mode_top_staff_pie))
        self._mode_top_staff_pie.toggled.connect(lambda _v: _guard_modes(self._mode_top_staff_bar, self._mode_top_staff_pie))
        self._mode_workshops_bar.toggled.connect(lambda _v: _guard_modes(self._mode_workshops_bar, self._mode_workshops_pie))
        self._mode_workshops_pie.toggled.connect(lambda _v: _guard_modes(self._mode_workshops_bar, self._mode_workshops_pie))
        _sync_modes_enabled()

        # ---- Format ----
        card_lay.addSpacing(2)
        card_lay.addWidget(_section_title("Формат"), 0, Qt.AlignLeft)
        self.rb_format_tables = QRadioButton("Только значения")
        self.rb_format_both = QRadioButton("Графики и значения")
        fmt = (options.format_label or "").strip()
        if fmt == "Только значения":
            self.rb_format_tables.setChecked(True)
        else:
            self.rb_format_both.setChecked(True)
        self._format_group = QButtonGroup(self)
        self._format_group.addButton(self.rb_format_tables)
        self._format_group.addButton(self.rb_format_both)
        for rb in (self.rb_format_tables, self.rb_format_both):
            rb.setStyleSheet(
                "QRadioButton {"
                f"font-family: {_STATS_FONT_FAMILY};"
                "font-size: 14px;"
                "font-weight: 600;"
                f"color: {_C_TEXT};"
                "padding: 6px 0;"
                "}"
                "QRadioButton::indicator { width: 18px; height: 18px; }"
            )
        f_row = QHBoxLayout()
        f_row.setContentsMargins(0, 0, 0, 0)
        f_row.setSpacing(18)
        f_row.addWidget(self.rb_format_tables, 0)
        f_row.addWidget(self.rb_format_both, 0)
        f_row.addStretch(1)
        card_lay.addLayout(f_row)

        # ---- Orientation ----
        card_lay.addSpacing(2)
        card_lay.addWidget(_section_title("Ориентация"), 0, Qt.AlignLeft)
        self.rb_portrait = QRadioButton("Книжная")
        self.rb_landscape = QRadioButton("Альбомная")
        orient = (options.orientation_label or "").strip()
        if orient == "Альбомная":
            self.rb_landscape.setChecked(True)
        else:
            self.rb_portrait.setChecked(True)
        self._orient_group = QButtonGroup(self)
        self._orient_group.addButton(self.rb_portrait)
        self._orient_group.addButton(self.rb_landscape)
        for rb in (self.rb_portrait, self.rb_landscape):
            rb.setStyleSheet(
                "QRadioButton {"
                f"font-family: {_STATS_FONT_FAMILY};"
                "font-size: 14px;"
                "font-weight: 600;"
                f"color: {_C_TEXT};"
                "padding: 6px 0;"
                "}"
                "QRadioButton::indicator { width: 18px; height: 18px; }"
            )
        o_row = QHBoxLayout()
        o_row.setContentsMargins(0, 0, 0, 0)
        o_row.setSpacing(18)
        o_row.addWidget(self.rb_portrait, 0)
        o_row.addWidget(self.rb_landscape, 0)
        o_row.addStretch(1)
        card_lay.addLayout(o_row)

        card_lay.addStretch(1)

        # ---- Actions ----
        actions = QHBoxLayout()
        actions.setContentsMargins(0, 6, 0, 0)
        actions.setSpacing(12)
        actions.addStretch(1)

        cancel_btn = QPushButton("Отмена")
        cancel_btn.setCursor(Qt.PointingHandCursor)
        cancel_btn.setStyleSheet(
            "QPushButton {"
            "background: #ffffff;"
            "border: 1px solid #e5e7eb;"
            "border-radius: 14px;"
            "padding: 12px 22px;"
            "min-height: 44px;"
            f"font-family: {_STATS_FONT_FAMILY};"
            "font-size: 14px;"
            f"font-weight: 700; color: {_C_TEXT};"
            "}"
            "QPushButton:hover { background: #f8fafc; }"
            "QPushButton:pressed { background: #eef2f7; }"
        )

        preview_btn = QPushButton("Предпросмотр")
        preview_btn.setCursor(Qt.PointingHandCursor)
        preview_btn.setStyleSheet(
            "QPushButton {"
            "background: #facc15;"
            "border: 1px solid #eab308;"
            "border-radius: 14px;"
            "padding: 12px 22px;"
            "min-height: 44px;"
            f"font-family: {_STATS_FONT_FAMILY};"
            "font-size: 14px;"
            "font-weight: 900;"
            "color: #78350f;"
            "}"
            "QPushButton:hover { background: #fde047; }"
            "QPushButton:pressed { background: #f59e0b; color: #451a03; border-color: #d97706; }"
        )

        actions.addWidget(cancel_btn, 0)
        actions.addWidget(preview_btn, 0)
        card_lay.addLayout(actions)

        root.addWidget(card, 1)

        cancel_btn.clicked.connect(self.reject)

        def _on_preview_clicked():
            # Validate: if a section is enabled, at least one chart mode is selected.
            if self.cb_top_products.isChecked() and (not self._mode_top_products_bar.isChecked()) and (not self._mode_top_products_pie.isChecked()):
                QMessageBox.warning(self, "Печать отчётов", "Для «Топ продуктов» выберите хотя бы один формат: гистограмма или диаграмма.")
                return
            if self.cb_top_staff.isChecked() and (not self._mode_top_staff_bar.isChecked()) and (not self._mode_top_staff_pie.isChecked()):
                QMessageBox.warning(self, "Печать отчётов", "Для «Топ сотрудников» выберите хотя бы один формат: гистограмма или диаграмма.")
                return
            if self.cb_workshops.isChecked() and (not self._mode_workshops_bar.isChecked()) and (not self._mode_workshops_pie.isChecked()):
                QMessageBox.warning(self, "Печать отчётов", "Для «Распределение по цехам» выберите хотя бы один формат: гистограмма или диаграмма.")
                return
            self.accept()

        preview_btn.clicked.connect(_on_preview_clicked)

    def selected_blocks(self) -> list[str]:
        out: list[str] = []
        if self.cb_summary.isChecked():
            out.append("Сводка")
        if self.cb_time_hist.isChecked():
            out.append("Гистограмма по времени")
        if self.cb_top_products.isChecked():
            out.append("Топ продуктов")
        if self.cb_top_staff.isChecked():
            out.append("Топ сотрудников")
        if self.cb_workshops.isChecked():
            out.append("Распределение по цехам")
        return out

    def selected_format_label(self) -> str:
        if self.rb_format_tables.isChecked():
            return "Только значения"
        return "Графики и значения"

    def selected_orientation_label(self) -> str:
        return "Книжная" if self.rb_portrait.isChecked() else "Альбомная"

    def selected_options(self) -> PrintReportOptions:
        opts = PrintReportOptions(
            period_label=self._options.period_label,
            period_subtitle=self._options.period_subtitle,
            selected_blocks=tuple(self.selected_blocks()),
            format_label=self.selected_format_label(),
            orientation_label=self.selected_orientation_label(),
            chart_modes=self._options.chart_modes,
        )

        def modes_from(btn_bar: QToolButton, btn_pie: QToolButton) -> set[str]:
            out: set[str] = set()
            if btn_bar.isChecked():
                out.add("bar")
            if btn_pie.isChecked():
                out.add("pie")
            return out

        opts = _set_chart_modes(opts, "top_products", modes_from(self._mode_top_products_bar, self._mode_top_products_pie))
        opts = _set_chart_modes(opts, "top_staff", modes_from(self._mode_top_staff_bar, self._mode_top_staff_pie))
        opts = _set_chart_modes(opts, "workshops", modes_from(self._mode_workshops_bar, self._mode_workshops_pie))
        return opts


class PrintReportsPreviewDialog(QDialog):
    RESULT_BACK = 2
    RESULT_PRINTED = 3

    def __init__(
        self,
        *,
        options: PrintReportOptions | None = None,
        period_label: str,
        period_subtitle: str,
        format_label: str,
        orientation_label: str,
        blocks: list[str],
        records: list[PrintRecord],
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setWindowTitle(" ")
        self.setModal(True)
        # Larger, more comfortable preview window by default
        self.setMinimumSize(1100, 820)
        self.resize(1280, 940)

        if options is None:
            options = PrintReportOptions(
                period_label=(period_label or "").strip() or "Период",
                period_subtitle=(period_subtitle or "").strip() or ((period_label or "").strip() or "Период"),
                selected_blocks=tuple(blocks or ()),
                format_label=(format_label or "").strip() or "Графики и значения",
                orientation_label=(orientation_label or "").strip() or "Книжная",
            )
        self._options = options
        self._records = list(records or [])

        self._print_btn: QPushButton | None = None
        self._pages: list[QWidget] = []

        self.setStyleSheet("QDialog { background: #f1f5f9; }")
        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(0)

        card = QFrame()
        card.setObjectName("PrintReportsPreviewCard")
        card.setStyleSheet(
            "#PrintReportsPreviewCard {"
            "background: #ffffff;"
            "border: 1px solid #e5e7eb;"
            "border-radius: 20px;"
            "}"
        )
        card_lay = QVBoxLayout(card)
        card_lay.setContentsMargins(18, 16, 18, 16)
        card_lay.setSpacing(12)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("QScrollArea { background: transparent; border: none; } QScrollBar { background: transparent; }")
        card_lay.addWidget(scroll, 1)

        pages_host = QWidget()
        pages_host.setStyleSheet("background: transparent; border: none;")
        pages_lay = QVBoxLayout(pages_host)
        pages_lay.setContentsMargins(0, 0, 0, 0)
        pages_lay.setSpacing(14)
        pages_lay.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        scroll.setWidget(pages_host)

        self._pages = build_report_pages(self._options, self._records)
        for pg in self._pages:
            pages_lay.addWidget(pg, 0, Qt.AlignHCenter)
        pages_lay.addStretch(1)

        # ---- Actions ----
        actions = QHBoxLayout()
        actions.setContentsMargins(0, 8, 0, 0)
        actions.setSpacing(12)
        actions.addStretch(1)

        back_btn = QPushButton("Назад")
        back_btn.setCursor(Qt.PointingHandCursor)
        back_btn.setStyleSheet(
            "QPushButton {"
            "background: #ffffff;"
            "border: 1px solid #e5e7eb;"
            "border-radius: 14px;"
            "padding: 12px 18px;"
            "min-height: 44px;"
            f"font-family: {_STATS_FONT_FAMILY};"
            "font-size: 14px;"
            f"font-weight: 700; color: {_C_TEXT};"
            "}"
            "QPushButton:hover { background: #f8fafc; }"
            "QPushButton:pressed { background: #eef2f7; }"
        )

        close_btn = QPushButton("Закрыть")
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setStyleSheet(
            "QPushButton {"
            "background: #ffffff;"
            "border: 1px solid #e5e7eb;"
            "border-radius: 14px;"
            "padding: 12px 22px;"
            "min-height: 44px;"
            f"font-family: {_STATS_FONT_FAMILY};"
            "font-size: 14px;"
            f"font-weight: 700; color: {_C_TEXT};"
            "}"
            "QPushButton:hover { background: #f8fafc; }"
            "QPushButton:pressed { background: #eef2f7; }"
        )

        print_btn = QPushButton("Печать…")
        print_btn.setCursor(Qt.PointingHandCursor)
        print_btn.setStyleSheet(
            "QPushButton {"
            "background: #facc15;"
            "border: 1px solid #eab308;"
            "border-radius: 14px;"
            "padding: 12px 22px;"
            "min-height: 44px;"
            f"font-family: {_STATS_FONT_FAMILY};"
            "font-size: 14px;"
            "font-weight: 900;"
            "color: #78350f;"
            "}"
            "QPushButton:hover { background: #fde047; }"
            "QPushButton:pressed { background: #f59e0b; color: #451a03; border-color: #d97706; }"
        )

        actions.addWidget(back_btn, 0)
        actions.addWidget(close_btn, 0)
        actions.addWidget(print_btn, 0)
        card_lay.addLayout(actions)

        root.addWidget(card, 1)

        close_btn.clicked.connect(self.reject)
        back_btn.clicked.connect(lambda: self.done(self.RESULT_BACK))

        self._print_btn = print_btn
        print_btn.clicked.connect(self._on_print_clicked)

    def _on_print_clicked(self) -> None:
        printer = QPrinter(QPrinter.HighResolution)
        try:
            printer.setPageSize(QPrinter.A4)
        except Exception:
            pass
        try:
            if (self._orientation_label or "").strip() == "Альбомная":
                printer.setOrientation(QPrinter.Landscape)
            else:
                printer.setOrientation(QPrinter.Portrait)
        except Exception:
            pass
        try:
            printer.setDocName("mirlis_stats_report")
        except Exception:
            pass

        dlg = QPrintDialog(printer, self)
        if dlg.exec_() != QDialog.Accepted:
            return

        try:
            print_pages_to_printer(printer, self._pages)
            self.done(self.RESULT_PRINTED)
        except Exception as e:
            QMessageBox.warning(self, "Печать", f"Не удалось распечатать отчёт:\n{e}")

