from __future__ import annotations

import math

from collections import Counter
from datetime import datetime, timedelta

from resources import resource_path
from statistics_detail_dialog import StatisticsDetailView
from statistics_data import (
    filter_records_by_period,
    filter_records_by_datetime_range,
    load_print_records_from_archive,
    normalize_stat_key,
    PrintRecord,
    compute_shift_totals,
)

from PyQt5.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QFrame,
    QLabel,
    QSizePolicy,
    QStackedLayout,
    QToolButton,
    QButtonGroup,
    QDialog,
    QDateTimeEdit,
    QApplication,
    QMessageBox,
    QPushButton,
    QFileDialog,
    QGraphicsDropShadowEffect,
    QStyledItemDelegate,
    QStyleOptionViewItem,
    QTableView,
    QCalendarWidget,
    QStyle,
    QAbstractSpinBox,
)
from PyQt5 import sip
from PyQt5.QtCore import QObject, Qt, QSize, QRect, QRectF, QDateTime, QLocale, QEvent, QTimer, QPoint, QModelIndex, pyqtSignal
from PyQt5.QtGui import (
    QPixmap,
    QPainter,
    QColor,
    QCursor,
    QPalette,
    QBrush,
    QFontMetrics,
    QImage,
    QLinearGradient,
    QRadialGradient,
    QPen,
    QFont,
    QIcon,
    QPainterPath,
    QRegion,
    QMouseEvent,
)


SAD_HERO_PATH = resource_path("assets/sad_hero.png")
ROLL_PATH = resource_path("assets/roll.png")
PRINTER_PATH = resource_path("assets/printer.png")
PRINT_REPORTS_BTN_ICON_PATH = resource_path("assets/printing reports.png")
SUN_PATH = resource_path("assets/sun.png")
MOON_PATH = resource_path("assets/moon.png")

# ---- Typography system (statistics area only) ----
_STATS_FONT_FAMILY = '"Inter","Segoe UI","Manrope","Arial",sans-serif'
_C_TITLE = "#1E2F45"   # main titles / KPI numbers
_C_TEXT = "#24364D"    # card titles / primary text
_C_SUB = "#6B7C93"     # secondary text (axes / labels)
_C_MUTED = "#7C8CA3"   # muted
# Подписи значений над столбцами «Этикеток по часам» (графит / navy, не чистый чёрный)
_C_HOUR_BAR_VALUE = "#102A43"

# Кнопки-плитки «Импорт и экспорт» и «Печать отчётов» (один визуальный класс).
_DASHBOARD_IO_TILE_BTN_HEIGHT_PX = 58
_DASHBOARD_IO_TILE_ICON_PX = 48


def _dashboard_io_tile_button_stylesheet() -> str:
    return (
        "QPushButton {"
        "background: #ffffff;"
        "border: 1px solid #d1d5db;"
        "border-radius: 12px;"
        f"font-family: {_STATS_FONT_FAMILY};"
        "font-size: 14px;"
        "font-weight: 600;"
        f"color: {_C_TEXT};"
        "padding: 10px 14px;"
        "text-align: left;"
        "}"
        "QPushButton:hover { background: #f8fafc; border-color: #9ca3af; }"
        "QPushButton:pressed { background: #f1f5f9; border-color: #6b7280; }"
    )


def _dashboard_io_tile_button_disabled_stylesheet() -> str:
    # Disabled-like (grey) style, but we keep the button clickable.
    return (
        "QPushButton {"
        "background: #f8fafc;"
        "border: 1px solid #e5e7eb;"
        "border-radius: 12px;"
        f"font-family: {_STATS_FONT_FAMILY};"
        "font-size: 14px;"
        "font-weight: 600;"
        f"color: {_C_MUTED};"
        "padding: 10px 14px;"
        "text-align: left;"
        "}"
        "QPushButton:hover { background: #f8fafc; border-color: #e5e7eb; }"
        "QPushButton:pressed { background: #f1f5f9; border-color: #e5e7eb; }"
    )


# Внутренняя зона графика (топ продукты / сотрудники / цехи): мягкая рамка и hover.
_STATS_CHART_INNER_QSS = """
QFrame#StatsChartInner {
    background: rgba(255, 255, 255, 0);
    border: 1px solid #eef2f7;
    border-radius: 12px;
}
QFrame#StatsChartInner:hover {
    background: rgba(99, 102, 241, 0.06);
    border: 1px solid rgba(99, 102, 241, 0.55);
}
"""


class _StatsChartClickFilter(QObject):
    """Клик по chart-area без подкласса QFrame (избегаем регрессии PyQt5 при нескольких подклассах)."""

    def __init__(self, handler, parent: QWidget | None = None):
        super().__init__(parent)
        self._handler = handler

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:  # type: ignore[override]
        if event.type() == QEvent.MouseButtonPress and isinstance(event, QMouseEvent):
            if event.button() == Qt.LeftButton and self._handler is not None:
                self._handler()
        return False


# Popup QCalendarWidget for custom period dialog (QDateTimeEdit calendar popup)
_CUSTOM_PERIOD_CALENDAR_QSS = """
QCalendarWidget {{
    background: #ffffff;
    border: none;
    border-radius: 16px;
    font-family: {ff};
    color: {ctext};
    font-size: 17px;
    margin: 0px;
    min-width: 440px;
    min-height: 400px;
}}
QCalendarWidget QWidget#qt_calendar_navigationbar {{
    background: #ffffff;
    border: none;
    border-bottom: 1px solid #d1d5db;
    border-top-left-radius: 16px;
    border-top-right-radius: 16px;
    min-height: 56px;
}}
QCalendarWidget QToolButton {{
    background: transparent;
    color: {ctext};
    border: none;
    border-radius: 10px;
    padding: 10px 14px;
    font-weight: 600;
    font-size: 17px;
}}
QCalendarWidget QToolButton:hover {{
    background: #f1f5f9;
}}
QCalendarWidget QToolButton:pressed {{
    background: #e2e8f0;
}}
QCalendarWidget QToolButton#qt_calendar_prevmonth,
QCalendarWidget QToolButton#qt_calendar_nextmonth {{
    min-width: 44px;
    min-height: 44px;
    padding: 8px;
}}
QCalendarWidget QMenu {{
    background: #ffffff;
    border: 1px solid #d1d5db;
    border-radius: 10px;
    padding: 6px;
    font-size: 16px;
}}
QCalendarWidget QMenu::item {{
    padding: 12px 20px;
    color: {ctext};
    border-radius: 8px;
    font-size: 16px;
}}
QCalendarWidget QMenu::item:selected {{
    background: #f1f5f9;
}}
QCalendarWidget QAbstractItemView {{
    outline: none;
    border: none;
    background: #ffffff;
    alternate-background-color: #ffffff;
    selection-background-color: #FACC15;
    selection-color: #78350f;
    gridline-color: transparent;
    font-size: 18px;
}}
QCalendarWidget QTableView,
QCalendarWidget QTableView#qt_calendar_calendarview {{
    outline: none;
    border: none;
    background: #ffffff;
    alternate-background-color: #ffffff;
    selection-background-color: #FACC15;
    selection-color: #78350f;
    gridline-color: transparent;
    border-bottom-left-radius: 16px;
    border-bottom-right-radius: 16px;
    font-size: 18px;
}}
QCalendarWidget QTableView::item {{
    padding: 12px 8px;
    min-height: 44px;
    color: {ctext};
    background: transparent;
    font-size: 18px;
}}
QCalendarWidget QTableView::item:disabled {{
    color: #9CA3AF;
}}
QCalendarWidget QTableView::item:selected {{
    background: #FACC15;
    color: #78350f;
}}
QCalendarWidget QTableView::item:selected:hover {{
    background: #fde047;
    color: #78350f;
}}
QCalendarWidget QHeaderView::section {{
    background: #ffffff;
    color: {csub};
    border: none;
    border-bottom: 1px solid #f1f5f9;
    padding: 12px 6px;
    font-weight: 600;
    font-size: 15px;
}}
""".format(
    ff=_STATS_FONT_FAMILY,
    ctext=_C_TEXT,
    csub=_C_SUB,
)


class _CalendarDayHoverDelegate(QStyledItemDelegate):
    """
    Hover-подсветка ячеек дат в QCalendarWidget (QSS на Windows часто не даёт :hover у item).
    Selected и disabled не перекрываем; цвета текста (выходные и т.д.) остаются из модели.
    """

    _HOVER_BG = QColor("#fef3c7")
    _RADIUS = 8.0
    _CELL_PAD = 3

    def __init__(self, table: QTableView, cal: QCalendarWidget):
        super().__init__(cal)
        self._view = table
        self._hover = QModelIndex()

    def _table(self) -> QTableView | None:
        """QTableView из popup-календаря может быть уничтожен раньше делегата."""
        v = self._view
        if v is None:
            return None
        try:
            if sip.isdeleted(v):
                return None
        except Exception:
            return None
        return v

    def eventFilter(self, watched, event: QEvent) -> bool:  # type: ignore[override]
        view = self._table()
        if view is None:
            return False
        try:
            vp = view.viewport()
        except RuntimeError:
            return False
        if vp is None or watched is not vp:
            return False
        et = event.type()
        if et == QEvent.MouseMove:
            try:
                idx = view.indexAt(event.pos())
            except RuntimeError:
                return False
            if idx != self._hover:
                old = self._hover
                self._hover = idx
                try:
                    if old.isValid():
                        view.update(old)
                    if idx.isValid():
                        view.update(idx)
                    else:
                        vp.update()
                except RuntimeError:
                    return False
        elif et == QEvent.Leave:
            if self._hover.isValid():
                old = self._hover
                self._hover = QModelIndex()
                try:
                    view.update(old)
                except RuntimeError:
                    return False
        return False

    def paint(self, painter: QPainter, option, index) -> None:
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)

        selected = bool(opt.state & QStyle.State_Selected)
        enabled = bool(opt.state & QStyle.State_Enabled)
        hovered = self._hover.isValid() and index.isValid() and index == self._hover

        if hovered and enabled and not selected:
            r = opt.rect.adjusted(self._CELL_PAD, 2, -self._CELL_PAD, -2)
            painter.save()
            painter.setRenderHint(QPainter.Antialiasing, True)
            painter.setPen(Qt.NoPen)
            painter.setBrush(self._HOVER_BG)
            painter.drawRoundedRect(r, self._RADIUS, self._RADIUS)
            painter.restore()
            opt.backgroundBrush = QBrush(Qt.transparent)
            opt.palette.setBrush(QPalette.Base, QBrush(Qt.transparent))
            opt.palette.setBrush(QPalette.AlternateBase, QBrush(Qt.transparent))

        super().paint(painter, opt, index)


def _install_calendar_day_hover_delegate(cal: QCalendarWidget) -> None:
    tv = cal.findChild(QTableView, "qt_calendar_calendarview")
    if tv is None:
        tv = cal.findChild(QTableView)
    if tv is None:
        return
    tv.setMouseTracking(True)
    delegate = _CalendarDayHoverDelegate(tv, cal)
    tv.setItemDelegate(delegate)
    tv.viewport().installEventFilter(delegate)


# Unified series palette (single source of truth).
# Used for per-item coloring in horizontal/vertical bars, pie/donut, and detail table markers.
_BAR_PALETTE_HEX = [
    "#3B66C3", "#2FAE7A", "#C35A6B", "#2F97AE", "#B88A2D",
    "#7A4EC3", "#2FAE4B", "#C35A2F", "#2F6AAE", "#6FAE2F",
    "#C34EB6", "#2FAEA8", "#C37D2F", "#4A2FAE", "#3EAE2F",
    "#C33D7D", "#2F7EAE", "#9EAE2F", "#AE2FA6", "#2FAE90",
    "#C34A3B", "#2F56AE", "#57AE2F", "#C32F92", "#2F9E9E",
    "#C3A23B", "#8C2FAE", "#2FAE60", "#C32F4A", "#2F86AE",
    "#86AE2F", "#AE2F8C", "#2FAEAA", "#C3652F", "#2F3EAE",
    "#4AAE2F", "#C32F6A", "#2F9AAE", "#B3AE2F", "#6A2FAE",
    "#2FAE73", "#C33B3B", "#2F72AE", "#7EAE2F", "#C32FAE",
    "#2FAE9E", "#C38A3B", "#5A2FAE", "#2FAE3E", "#C33B86"
]


class _Card(QFrame):
    def __init__(self, title: str, parent=None):
        super().__init__(parent)
        self.setObjectName("StatsCard")
        self.setStyleSheet("background: #ffffff; border: 1px solid #e5e7eb; border-radius: 18px;")

        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 14, 16, 14)
        lay.setSpacing(10)

        header = QWidget()
        header.setStyleSheet("background: transparent; border: none;")
        header_lay = QHBoxLayout(header)
        header_lay.setContentsMargins(0, 0, 0, 0)
        header_lay.setSpacing(8)

        self.title_lbl = QLabel(title)
        self.title_lbl.setStyleSheet(
            f'font-family: {_STATS_FONT_FAMILY}; font-size: 17px; font-weight: 600; line-height: 20px; color: {_C_TITLE}; '
            "background: transparent; border: none; outline: none; padding: 0; margin: 0;"
        )
        header_lay.addWidget(self.title_lbl, 1, Qt.AlignLeft | Qt.AlignVCenter)

        self._header_right = QWidget()
        self._header_right.setStyleSheet("background: transparent; border: none;")
        self._header_right_lay = QHBoxLayout(self._header_right)
        self._header_right_lay.setContentsMargins(0, 0, 0, 0)
        self._header_right_lay.setSpacing(6)
        header_lay.addWidget(self._header_right, 0, Qt.AlignRight | Qt.AlignVCenter)

        lay.addWidget(header, 0)

        self.body = QWidget()
        lay.addWidget(self.body, 1)

        self.body_lay = QVBoxLayout(self.body)
        self.body_lay.setContentsMargins(0, 0, 0, 0)
        self.body_lay.setSpacing(10)

    def set_header_right_widget(self, w: QWidget | None):
        """Optional compact control on the right side of card header."""
        while self._header_right_lay.count():
            item = self._header_right_lay.takeAt(0)
            if item and item.widget():
                item.widget().setParent(None)
        if w is not None:
            self._header_right_lay.addWidget(w, 0, Qt.AlignRight | Qt.AlignVCenter)


class _KpiCard(QFrame):
    def __init__(self, title: str, icon_path: str, parent=None):
        super().__init__(parent)
        self.setObjectName("StatsCard")
        self.setStyleSheet("background: #ffffff; border: 1px solid #e5e7eb; border-radius: 18px;")
        self.setFixedHeight(102)

        self._icon_path = icon_path
        self._icon_box = 188

        root = QHBoxLayout(self)
        root.setContentsMargins(16, 10, 16, 8)
        root.setSpacing(12)

        left = QVBoxLayout()
        left.setContentsMargins(0, 0, 0, 0)
        left.setSpacing(4)

        self.title_lbl = QLabel(title)
        self.title_lbl.setStyleSheet(
            f'font-family: {_STATS_FONT_FAMILY}; font-size: 17px; font-weight: 600; line-height: 18px; color: {_C_TEXT}; '
            "background: transparent; border: none; outline: none; padding: 0; margin: 0;"
        )
        left.addWidget(self.title_lbl, 0, Qt.AlignLeft | Qt.AlignTop)

        self.value_lbl = QLabel("0")
        self.value_lbl.setStyleSheet(
            f'font-family: {_STATS_FONT_FAMILY}; font-size: 28px; font-weight: 700; line-height: 32px; color: {_C_TITLE}; '
            "background: transparent; border: none; outline: none;"
        )
        left.addWidget(self.value_lbl, 0, Qt.AlignLeft | Qt.AlignTop)
        left.addStretch(1)
        root.addLayout(left, 1)

        self.icon_wrap = QWidget()
        self.icon_wrap.setStyleSheet("background: transparent; border: none; outline: none;")
        self.icon_wrap.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding)
        icon_wrap_lay = QVBoxLayout(self.icon_wrap)
        if icon_path == ROLL_PATH:
            top_icon_margin = 14
        elif icon_path == PRINTER_PATH:
            top_icon_margin = 8
        else:
            top_icon_margin = 4
        icon_wrap_lay.setContentsMargins(0, top_icon_margin, 0, 0)
        icon_wrap_lay.setSpacing(0)

        self.icon_lbl = QLabel()
        self.icon_lbl.setAlignment(Qt.AlignRight | Qt.AlignTop)
        self.icon_lbl.setStyleSheet("background: transparent; border: none; outline: none;")
        self.icon_lbl.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        icon_wrap_lay.addWidget(self.icon_lbl, 0, Qt.AlignRight | Qt.AlignTop)
        icon_wrap_lay.addStretch(1)
        root.addWidget(self.icon_wrap, 0, Qt.AlignRight | Qt.AlignTop)

        self._refresh_icon_pixmap()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._refresh_icon_pixmap()

    def _trim_transparent(self, pix: QPixmap) -> QPixmap:
        img = pix.toImage().convertToFormat(QImage.Format_ARGB32)
        w, h = img.width(), img.height()
        min_x, min_y = w, h
        max_x, max_y = -1, -1
        for y in range(h):
            for x in range(w):
                if (img.pixel(x, y) >> 24) & 0xFF:
                    min_x = min(min_x, x)
                    min_y = min(min_y, y)
                    max_x = max(max_x, x)
                    max_y = max(max_y, y)
        if max_x >= min_x and max_y >= min_y:
            img = img.copy(min_x, min_y, max_x - min_x + 1, max_y - min_y + 1)
            return QPixmap.fromImage(img)
        return pix

    def _set_icon(self, path: str):
        pix = QPixmap(path)
        if pix.isNull():
            return
        pix = self._trim_transparent(pix)
        target = min(int(self._icon_box), max(40, int(self.height()) - 14))
        pix = pix.scaled(target, target, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.icon_lbl.setPixmap(pix)
        self.icon_lbl.setFixedSize(pix.size())
        self.icon_wrap.setFixedWidth(max(pix.width(), 64))

    def _refresh_icon_pixmap(self):
        if self._icon_path:
            self._set_icon(self._icon_path)

    def set_value(self, v: int):
        self.value_lbl.setText(str(int(v)))


class _CompactKpiCard(QFrame):
    def __init__(self, title: str, icon_path: str | None = None, parent=None):
        super().__init__(parent)
        self.setObjectName("StatsCard")
        self.setStyleSheet("background: #ffffff; border: 1px solid #e5e7eb; border-radius: 18px;")
        self.setFixedHeight(102)

        self._icon_path = icon_path

        root = QHBoxLayout(self)
        root.setContentsMargins(14, 10, 14, 8)
        root.setSpacing(10)

        left = QVBoxLayout()
        left.setContentsMargins(0, 0, 0, 0)
        left.setSpacing(4)

        self.title_lbl = QLabel(title)
        self.title_lbl.setStyleSheet(
            f'font-family: {_STATS_FONT_FAMILY}; font-size: 17px; font-weight: 600; line-height: 18px; color: {_C_TEXT}; '
            "background: transparent; border: none; outline: none; padding: 0; margin: 0;"
        )
        left.addWidget(self.title_lbl, 0, Qt.AlignLeft | Qt.AlignTop)

        self.value_lbl = QLabel("0")
        self.value_lbl.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.value_lbl.setStyleSheet(
            f'font-family: {_STATS_FONT_FAMILY}; font-size: 28px; font-weight: 700; line-height: 32px; color: {_C_TITLE}; '
            "background: transparent; border: none; outline: none;"
        )
        left.addWidget(self.value_lbl, 0, Qt.AlignLeft | Qt.AlignTop)
        left.addStretch(1)
        root.addLayout(left, 1)

        self.icon_wrap = QWidget()
        self.icon_wrap.setStyleSheet("background: transparent; border: none; outline: none;")
        self.icon_wrap.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding)
        icon_lay = QVBoxLayout(self.icon_wrap)
        icon_lay.setContentsMargins(0, 10, 0, 0)
        icon_lay.setSpacing(0)

        self.icon_lbl = QLabel()
        self.icon_lbl.setAlignment(Qt.AlignRight | Qt.AlignTop)
        self.icon_lbl.setStyleSheet("background: transparent; border: none; outline: none;")
        self.icon_lbl.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        icon_lay.addWidget(self.icon_lbl, 0, Qt.AlignRight | Qt.AlignTop)
        icon_lay.addStretch(1)
        root.addWidget(self.icon_wrap, 0, Qt.AlignRight | Qt.AlignTop)

        self._refresh_icon_pixmap()

    def set_value(self, v: int):
        self.value_lbl.setText(str(int(v)))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._refresh_icon_pixmap()

    def _refresh_icon_pixmap(self):
        path = getattr(self, "_icon_path", None)
        if not path:
            self.icon_lbl.clear()
            return
        pix = QPixmap(path)
        if pix.isNull():
            return
        # Делаем иконки смен по размеру сопоставимыми с обычными KPI-иконками
        desired = 188
        max_target = max(40, int(self.height()) - 14)
        target = min(desired, max_target)
        pix = pix.scaled(target, target, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.icon_lbl.setPixmap(pix)
        self.icon_lbl.setFixedSize(pix.size())
        self.icon_wrap.setFixedWidth(max(pix.width(), 40))


def _hour_bucket_is_night(hour_index: int) -> bool:
    """Бакет i = час 0..23. Ночь 20:00–08:00: hour >= 20 или hour < 8; день 8 <= hour < 20."""
    h = int(hour_index) % 24
    return h >= 20 or h < 8


def _vbar_draw_hour_chart_value_label(painter: QPainter, rect: QRect, text: str, *, night_zone: bool) -> None:
    """
    Подписи значений над столбцами «Этикеток по часам»: UI-style, без грубой обводки.

    Ночной фон: едва заметный светлый микро-ореол (читаемость на тёмной зоне).
    Дневной фон: лёгкая «воздушная» тень (глубина без грязи).
    Основной цвет — графитово-синий _C_HOUR_BAR_VALUE.
    """
    painter.save()
    painter.setRenderHint(QPainter.TextAntialiasing, True)
    flags = int(Qt.AlignHCenter | Qt.AlignVCenter)
    main = QColor(_C_HOUR_BAR_VALUE)
    if night_zone:
        halo = QColor(255, 255, 255, 20)
        painter.setPen(halo)
        for dx, dy in ((-1, 0), (1, 0), (0, -1), (0, 1)):
            painter.drawText(rect.translated(dx, dy), flags, text)
    else:
        painter.setPen(QColor(16, 39, 67, 34))
        painter.drawText(rect.translated(0, 1), flags, text)
    painter.setPen(main)
    painter.drawText(rect, flags, text)
    painter.restore()


class _VBarChart(QWidget):
    hoverLabelChanged = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._labels: list[str] = []
        self._values: list[int] = []
        self._bar_color = QColor("#93c5fd")  # light blue
        self._grid_color = QColor("#eef2f7")
        self._axis_color = QColor("#cbd5e1")
        self._text_color = QColor(_C_SUB)
        self._show_all_x_labels = False
        self._hour_zones_bg = False
        self._warm_vertical_gradient = False
        self._detail_mode = False
        self._hover_label: str = ""
        self._hover_index: int = -1
        self._last_plot: QRectF | None = None
        self._last_slot: float = 0.0
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setMouseTracking(True)

    def set_detail_mode(self, enabled: bool) -> None:
        """Detail mode: enable hover sync with the right table."""
        self._detail_mode = bool(enabled)
        self.update()

    def set_hover_label(self, label: str) -> None:
        lab = (label or "").strip()
        if lab == self._hover_label:
            return
        self._hover_label = lab
        try:
            self._hover_index = self._labels.index(lab) if lab else -1
        except ValueError:
            self._hover_index = -1
        self.update()

    def _set_hover_index(self, idx: int) -> None:
        idx = int(idx)
        if idx < 0 or idx >= len(self._labels):
            idx = -1
        if idx == self._hover_index:
            return
        self._hover_index = idx
        self._hover_label = self._labels[idx] if idx >= 0 else ""
        self.hoverLabelChanged.emit(self._hover_label)
        self.update()

    def set_warm_vertical_gradient(self, enabled: bool) -> None:
        """Тёплый вертикальный градиент в столбцах (только time_chart: 7/30 дней, custom не по часам)."""
        self._warm_vertical_gradient = bool(enabled)
        self.update()

    def set_data(self, labels: list[str], values: list[int], bar_color: QColor | None = None):
        self._labels = list(labels or [])
        self._values = [int(v) for v in (values or [])]
        if bar_color is not None:
            self._bar_color = bar_color
        self.update()

    def set_show_all_x_labels(self, enabled: bool):
        self._show_all_x_labels = bool(enabled)
        self.update()

    def set_hour_zones_background(self, enabled: bool):
        """Рисовать адаптивный фон день/ночь (только для графика по часам)."""
        self._hour_zones_bg = bool(enabled)
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)

        r = self.rect().adjusted(2, 2, -2, -2)
        painter.fillRect(r, Qt.transparent)

        if not self._values:
            painter.end()
            return

        left_pad = 40
        bottom_pad = 44
        top_pad = 22
        right_pad = 6

        n = len(self._values)
        if n == 0:
            painter.end()
            return

        plot = r.adjusted(left_pad, top_pad, -right_pad, -bottom_pad)
        slot = plot.width() / n
        self._last_plot = QRectF(plot)
        self._last_slot = float(slot)

        # background zones for 24-hour charts (night/day/night)
        if self._hour_zones_bg and n == 24:
            painter.save()
            painter.setPen(Qt.NoPen)

            # Draw background a bit higher than plot-area so value labels
            # don't look like they're going above the colored ceiling.
            bg_extra_top = 16.0
            bg_top = max(float(r.top()), float(plot.top()) - bg_extra_top)
            bg_rect = QRectF(float(plot.left()), bg_top, float(plot.width()), float(plot.bottom()) - bg_top)

            def zone_rect(h_from: int, h_to: int) -> QRectF:
                x0 = plot.left() + h_from * slot
                x1 = plot.left() + (h_to + 1) * slot
                return QRectF(float(x0), bg_top, float(x1 - x0), float(plot.bottom()) - bg_top)

            # Night: 00-07
            night1 = zone_rect(0, 7)
            g1 = QLinearGradient(night1.left(), night1.top(), night1.left(), night1.bottom())
            g1.setColorAt(0.0, QColor(49, 46, 129, 104))  # indigo (deeper)
            g1.setColorAt(1.0, QColor(15, 23, 42, 34))    # slate
            painter.fillRect(night1, g1)

            # Day: 08-19
            day = zone_rect(8, 19)
            gd = QLinearGradient(day.left(), day.top(), day.left(), day.bottom())
            gd.setColorAt(0.0, QColor(224, 242, 254, 104))  # sky (cleaner)
            gd.setColorAt(1.0, QColor(224, 242, 254, 26))
            painter.fillRect(day, gd)

            # Night: 20-23
            night2 = zone_rect(20, 23)
            g2 = QLinearGradient(night2.left(), night2.top(), night2.left(), night2.bottom())
            g2.setColorAt(0.0, QColor(49, 46, 129, 112))
            g2.setColorAt(1.0, QColor(15, 23, 42, 38))
            painter.fillRect(night2, g2)

            # shift dividers at 08:00 and 20:00
            painter.setPen(QColor(71, 85, 105, 48))
            x_8 = plot.left() + 8 * slot
            x_20 = plot.left() + 20 * slot
            painter.drawLine(int(x_8), int(bg_top), int(x_8), int(plot.bottom()))
            painter.drawLine(int(x_20), int(bg_top), int(x_20), int(plot.bottom()))

            # subtle global vertical lighting to avoid flatness
            vg = QLinearGradient(plot.left(), bg_top, plot.left(), plot.bottom())
            vg.setColorAt(0.0, QColor(255, 255, 255, 26))
            vg.setColorAt(1.0, QColor(255, 255, 255, 0))
            painter.fillRect(bg_rect, vg)

            # decorative sun/moon (simple primitives)
            y_icon = plot.top() + 16
            r_icon = 11  # ~+20-30% vs previous for better readability

            # sun (circle + rays + soft glow) over day zone
            sun_c = day.center().x()
            painter.setPen(Qt.NoPen)

            # soft glow (radial)
            sun_glow_r = r_icon + 10
            sg = QRadialGradient(sun_c, y_icon, sun_glow_r)
            sg.setColorAt(0.0, QColor(253, 224, 71, 90))   # warm yellow
            sg.setColorAt(0.55, QColor(56, 189, 248, 30))  # sky tint
            sg.setColorAt(1.0, QColor(56, 189, 248, 0))
            painter.setBrush(sg)
            painter.drawEllipse(QRectF(sun_c - sun_glow_r, y_icon - sun_glow_r, sun_glow_r * 2, sun_glow_r * 2))

            # core
            painter.setBrush(QColor(250, 204, 21, 190))  # slightly richer center
            painter.drawEllipse(QRectF(sun_c - r_icon, y_icon - r_icon, r_icon * 2, r_icon * 2))
            painter.setBrush(QColor(255, 240, 170, 95))  # subtle highlight
            painter.drawEllipse(QRectF(sun_c - (r_icon * 0.55), y_icon - (r_icon * 0.55), (r_icon * 1.10), (r_icon * 1.10)))

            # rays: a bit longer + thicker, still subtle
            ray_len = r_icon + 4
            ray_inner = r_icon - 1
            sun_pen = QPen(QColor(250, 204, 21, 165))
            sun_pen.setWidthF(1.6)
            sun_pen.setCapStyle(Qt.RoundCap)
            painter.setPen(sun_pen)
            for dx, dy in [(0, -1), (0.7, -0.7), (1, 0), (0.7, 0.7), (0, 1), (-0.7, 0.7), (-1, 0), (-0.7, -0.7)]:
                x0 = sun_c + dx * ray_inner
                y0 = y_icon + dy * ray_inner
                x1 = sun_c + dx * ray_len
                y1 = y_icon + dy * ray_len
                painter.drawLine(int(x0), int(y0), int(x1), int(y1))
            painter.setPen(Qt.NoPen)

            # moon (crescent + soft glow) over night zone
            moon_c = night1.center().x()
            painter.setPen(Qt.NoPen)
            moon_r = r_icon + 2

            # soft glow (radial)
            moon_glow_r = moon_r + 9
            mg = QRadialGradient(moon_c, y_icon, moon_glow_r)
            mg.setColorAt(0.0, QColor(255, 248, 220, 70))  # warm, light glow
            mg.setColorAt(0.65, QColor(219, 234, 254, 22))
            mg.setColorAt(1.0, QColor(219, 234, 254, 0))
            painter.setBrush(mg)
            painter.drawEllipse(QRectF(moon_c - moon_glow_r, y_icon - moon_glow_r, moon_glow_r * 2, moon_glow_r * 2))

            # crescent body (slightly warmer/lighter)
            painter.setBrush(QColor(255, 248, 220, 190))
            painter.drawEllipse(QRectF(moon_c - moon_r, y_icon - moon_r, moon_r * 2, moon_r * 2))

            # cut-out (controls thickness)
            cut_r = moon_r * 0.92
            cut_dx = 4.0
            painter.setBrush(QColor(15, 23, 42, 52))
            painter.drawEllipse(QRectF(moon_c - cut_r + cut_dx, y_icon - cut_r, cut_r * 2, cut_r * 2))

            painter.restore()

        max_v = max(self._values) if max(self._values) > 0 else 1
        is_hour_chart = bool(self._hour_zones_bg and n == 24)

        # Y-axis upper bound:
        # Keep 1 unit headroom on all vertical bar charts so the tallest bar
        # doesn't touch the plot ceiling and its value label can stay above the bar.
        scale_max = max_v + 1

        # Steps / ticks:
        # - For small ranges show integer шкала 0..scale_max (no fractional rounding => no duplicates)
        # - For larger ranges keep compact grid (3 steps)
        if scale_max <= 4:
            steps = int(scale_max)
        elif (not is_hour_chart) and max_v <= 3:
            steps = int(max_v)
        else:
            steps = 3
        fm = QFontMetrics(painter.font())
        for i in range(steps + 1):
            y = plot.top() + (plot.height() * i) / steps
            painter.setPen(self._grid_color)
            painter.drawLine(plot.left(), int(y), plot.right(), int(y))
            if scale_max <= 4:
                tick_value = (steps - i)
            else:
                tick_value = int(round(scale_max * (steps - i) / steps))
            tick_font = painter.font()
            tick_font.setFamily("Segoe UI")
            tick_font.setPointSize(11)
            tick_font.setWeight(QFont.Medium)  # ~500
            painter.setFont(tick_font)
            painter.setPen(QColor(_C_SUB))
            painter.drawText(r.left(), int(y) - 8, left_pad - 8, 16, Qt.AlignRight | Qt.AlignVCenter, str(tick_value))

        painter.setPen(self._axis_color)
        painter.drawLine(plot.left(), plot.bottom(), plot.right(), plot.bottom())

        bar_w = max(6.0, slot * 0.55)

        painter.setPen(Qt.NoPen)
        for i, v in enumerate(self._values):
            # Суточная гистограмма: нулевой час — без столбца, контура и подписи (никакой полоски у оси).
            if is_hour_chart and int(v) <= 0:
                continue
            x = plot.left() + i * slot + (slot - bar_w) / 2
            h = (v / scale_max) * plot.height()
            y = plot.bottom() - h
            is_active = (self._hover_index == i) and (self._hover_index >= 0)
            has_hover = self._hover_index >= 0
            # Bar colors:
            # - hour chart (день, 24 бакета): ночь — голубая заливка + белый контур; день — синяя + чёрный контур
            # - time_chart 7/30 дней и custom (не по часам): тёплый вертикальный градиент внутри столбца
            # - прочие вертикальные (цехи): палитра по индексу
            if is_hour_chart:
                if _hour_bucket_is_night(i):
                    fill = QColor("#4dabf7")  # спокойный голубой для ночных столбцов
                    outline = QColor("#ffffff")
                else:
                    fill = QColor("#2563eb")  # blue-600 / синяя заливка
                    outline = QColor("#000000")
                if has_hover and not is_active and getattr(self, "_detail_mode", False):
                    fill.setAlpha(170)
                bar_pen = QPen(outline)
                bar_pen.setWidthF(1.45)
                bar_pen.setJoinStyle(Qt.RoundJoin)
                bar_pen.setCapStyle(Qt.RoundCap)
                painter.setPen(bar_pen)
                painter.setBrush(fill)
            elif self._warm_vertical_gradient:
                # Низ столбца — жёлтый, верх — тёплый красно-оранжевый (выше столбец — больше «горячего»)
                grad = QLinearGradient(float(x), float(y), float(x), float(y + h))
                grad.setColorAt(0.0, QColor("#E85D4A"))
                grad.setColorAt(0.5, QColor("#F29E4C"))
                grad.setColorAt(1.0, QColor("#F4C20D"))
                outline = QColor("#c2410c")
                outline.setAlpha(180)
                pen = QPen(outline)
                pen.setWidthF(1.0)
                painter.setPen(pen)
                painter.setBrush(grad)
            else:
                base = QColor(_BAR_PALETTE_HEX[i % len(_BAR_PALETTE_HEX)])

                def hover_fill_hsv(c: QColor) -> QColor:
                    hh, ss, vv0, aa = c.getHsv()
                    if hh < 0:
                        return c
                    ss2 = min(255, int(ss * 1.32) + 8)   # ~+25–35%
                    vv2 = min(255, int(vv0 * 1.07) + 2)  # ~+5–10%
                    out = QColor()
                    out.setHsv(hh, ss2, vv2, aa)
                    return out

                fill = QColor(base)
                if has_hover and not is_active and getattr(self, "_detail_mode", False):
                    fill.setAlpha(190)
                if is_active and getattr(self, "_detail_mode", False):
                    fill = hover_fill_hsv(fill)

                if is_active and getattr(self, "_detail_mode", False):
                    pen = QPen(QColor(fill).darker(160))
                    pen.setWidthF(2.0)
                    pen.setJoinStyle(Qt.RoundJoin)
                    painter.setPen(pen)
                else:
                    painter.setPen(Qt.NoPen)
                painter.setBrush(fill)

            rect = QRectF(x, y, bar_w, h)
            if is_active and getattr(self, "_detail_mode", False):
                rect = rect.adjusted(-1.0, -1.0, 1.0, 1.0)
            painter.drawRoundedRect(rect, 4, 4)
            value_font = painter.font()
            if is_hour_chart:
                value_font = QFont()
                if hasattr(value_font, "setFamilies"):
                    value_font.setFamilies(["Inter", "Segoe UI", "Roboto", "Arial"])
                else:
                    value_font.setFamily("Segoe UI")
                value_font.setPointSize(13)
                value_font.setWeight(QFont.DemiBold)
                value_font.setStyleStrategy(QFont.PreferAntialias | QFont.PreferQuality)
            else:
                value_font.setFamily("Segoe UI")
                value_font.setPointSize(12)
                value_font.setWeight(QFont.DemiBold)
            painter.setFont(value_font)
            # Value label above bar: give it enough height and safe vertical alignment
            # to avoid clipping (tops of digits) on small and large values.
            text_h = 24 if is_hour_chart else 20
            fixed_gap = 10  # ~8–12 px над верхом столбца
            # place the label above the bar top with a fixed pixel gap (does not depend on Y-scale)
            # clamp only to the chart's inner rect (not plot-area), so the label can use top padding
            # when the tallest bar is very close to plot.top.
            text_y = int(y) - text_h - fixed_gap
            text_y = max(int(r.top()) + 2, text_y)
            val_rect = QRect(int(x - 10), int(text_y), int(bar_w + 20), int(text_h))
            if is_hour_chart:
                _vbar_draw_hour_chart_value_label(
                    painter,
                    val_rect,
                    str(int(v)),
                    night_zone=_hour_bucket_is_night(i),
                )
            else:
                painter.setPen(QColor(_C_TEXT))
                painter.drawText(
                    val_rect,
                    Qt.AlignHCenter | Qt.AlignVCenter,
                    str(int(v)),
                )
            painter.setPen(Qt.NoPen)
            # brush will be set per-bar on next iteration

        label_font = painter.font()
        label_font.setFamily("Segoe UI")
        label_font.setPointSize(11)
        label_font.setWeight(QFont.Medium)  # ~500
        painter.setFont(label_font)
        painter.setPen(self._text_color)
        show_every = 1
        if not self._show_all_x_labels and n > 12:
            show_every = 2
        if not self._show_all_x_labels and n > 18:
            show_every = 3
        if self._show_all_x_labels:
            # keep small but readable for 24h labels
            label_font.setPointSize(11)
            painter.setFont(label_font)
            fm = QFontMetrics(painter.font())
        for i, lab in enumerate(self._labels):
            if i % show_every != 0:
                continue
            x = plot.left() + i * slot
            text = str(lab)
            text_w = max(10, int(slot) - 6)
            text_x = int(x + 3)
            text_y = int(plot.bottom() + 8)
            elided = fm.elidedText(text, Qt.ElideRight, text_w)
            painter.drawText(text_x, text_y, text_w, bottom_pad - 8, Qt.AlignHCenter | Qt.AlignTop, elided)

        painter.end()

    def mouseMoveEvent(self, event):  # type: ignore[override]
        if not getattr(self, "_detail_mode", False):
            return super().mouseMoveEvent(event)
        plot = self._last_plot
        if plot is None or self._last_slot <= 0:
            return super().mouseMoveEvent(event)
        x = float(event.pos().x())
        if x < float(plot.left()) or x > float(plot.right()):
            self._set_hover_index(-1)
            return super().mouseMoveEvent(event)
        idx = int((x - float(plot.left())) // float(self._last_slot))
        self._set_hover_index(idx)
        return super().mouseMoveEvent(event)

    def leaveEvent(self, event):  # type: ignore[override]
        if getattr(self, "_detail_mode", False):
            self._set_hover_index(-1)
        return super().leaveEvent(event)


class _HBarChart(QWidget):
    hoverLabelChanged = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._labels: list[str] = []
        self._values: list[int] = []
        self._bar_color = QColor("#facc15")  # yellow
        self._grid_color = QColor("#f3f4f6")
        self._text_color = QColor(_C_TEXT)
        self._muted = QColor(_C_SUB)
        self._bar_width_factor = 1.0
        self._detail_mode = False
        self._hover_label: str = ""
        self._hover_index: int = -1
        self._last_plot: QRectF | None = None
        self._last_row_h: float = 0.0
        self._last_n: int = 0
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setMouseTracking(True)

    def set_detail_mode(self, enabled: bool) -> None:
        """Detail mode: reduce inner top air and avoid vertically inflated rows."""
        self._detail_mode = bool(enabled)
        self.update()

    def set_hover_label(self, label: str) -> None:
        lab = (label or "").strip()
        if lab == self._hover_label:
            return
        self._hover_label = lab
        try:
            self._hover_index = self._labels.index(lab) if lab else -1
        except ValueError:
            self._hover_index = -1
        self.update()

    def _set_hover_index(self, idx: int) -> None:
        idx = int(idx)
        if idx < 0 or idx >= len(self._labels):
            idx = -1
        if idx == self._hover_index:
            return
        self._hover_index = idx
        self._hover_label = self._labels[idx] if idx >= 0 else ""
        self.hoverLabelChanged.emit(self._hover_label)
        self.update()

    def set_bar_width_factor(self, factor: float):
        try:
            f = float(factor)
        except Exception:
            f = 1.0
        self._bar_width_factor = max(0.5, min(1.0, f))
        self.update()

    def set_data(self, labels: list[str], values: list[int], bar_color: QColor | None = None):
        self._labels = list(labels or [])
        self._values = [int(v) for v in (values or [])]
        if bar_color is not None:
            self._bar_color = bar_color
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)

        if getattr(self, "_detail_mode", False):
            r = self.rect()
        else:
            r = self.rect().adjusted(2, 2, -2, -2)
        painter.fillRect(r, Qt.transparent)

        n = len(self._values)
        if n == 0:
            painter.end()
            return

        is_detail = bool(getattr(self, "_detail_mode", False))

        # Fonts & metrics (used for both geometry and drawing)
        label_font = painter.font()
        label_font.setFamily("Segoe UI")
        label_font.setPointSize(11)
        label_font.setWeight(QFont.Medium)  # ~500
        fm = QFontMetrics(label_font)

        value_font = painter.font()
        value_font.setFamily("Segoe UI")
        value_font.setPointSize(12)
        value_font.setWeight(QFont.DemiBold)  # ~600
        fm_v = QFontMetrics(value_font)

        # Value area: keep predictable width for numbers
        max_v = max(self._values) if max(self._values) > 0 else 1
        value_txt = str(int(max_v))
        value_w = max(52, fm_v.horizontalAdvance(value_txt) + 16)

        # Horizontal geometry
        if is_detail:
            # Detail horizontal geometry (explicit, top-level rect anchored):
            # total = full available rect
            # label_w = clamp(longest + padding, min, max)
            # value_w = compact width for number
            # track fills the remaining space, value is near right edge
            total = r

            right_margin = 8.0
            gap = 8.0
            min_label_w = 120.0
            max_label_w = 230.0

            # value_w based on the widest number we need to show
            value_w = max(44.0, float(fm_v.horizontalAdvance(value_txt)) + 8.0)

            longest = 0.0
            for lab in self._labels:
                try:
                    longest = max(longest, float(fm.horizontalAdvance(str(lab))))
                except Exception:
                    continue
            label_w = max(min_label_w, min(max_label_w, longest + 18.0))

            left_x = float(total.left())
            right_x = float(total.right())

            track_left = left_x + label_w
            value_left = right_x - right_margin - value_w
            track_right = value_left - gap

            track_w = max(10.0, track_right - track_left)

            # Backward-compatible names used below for drawing bars.
            bar_left = track_left
            bar_right = track_right
            bar_area_w = track_w
        else:
            left_pad = 140
            right_pad = 26
        top_pad = 0 if getattr(self, "_detail_mode", False) else 8
        bottom_pad = 0 if getattr(self, "_detail_mode", False) else 8
        if is_detail:
            # Vertical plot rect uses the full width; horizontal geometry is explicit (bar_left/bar_right).
            plot = r.adjusted(0, top_pad, 0, -bottom_pad)
            bar_right = bar_left + bar_area_w
        else:
            plot = r.adjusted(left_pad, top_pad, -right_pad, -bottom_pad)
            bar_area_w = plot.width() * float(getattr(self, "_bar_width_factor", 1.0))
            bar_right = plot.left() + bar_area_w

        if getattr(self, "_detail_mode", False):
            desired_row_h = max(20, int(fm.height()) + 8)
            # If all rows fit with compact height — keep bars near the top,
            # leave any extra free space at the bottom (instead of inflating each row).
            if desired_row_h * n <= plot.height():
                row_h = desired_row_h
            else:
                row_h = max(22, int(plot.height() / max(n, 1)))
        else:
            row_h = max(22, int(plot.height() / max(n, 1)))

        bar_h = max(10, int(row_h * (0.82 if getattr(self, "_detail_mode", False) else 0.55)))
        self._last_plot = plot
        self._last_row_h = float(row_h)
        self._last_n = int(n)

        for i, (lab, v) in enumerate(zip(self._labels, self._values)):
            y_top = plot.top() + i * row_h
            y_center = y_top + row_h / 2

            # label
            painter.setFont(label_font)
            painter.setPen(self._muted)
            text = str(lab)
            if is_detail:
                label_rect_w = max(10, int(label_w) - 12)
                elided = fm.elidedText(text, Qt.ElideRight, label_rect_w)
                painter.drawText(int(r.left() + 6), int(y_top), label_rect_w, row_h, Qt.AlignVCenter | Qt.AlignLeft, elided)
            else:
                elided = fm.elidedText(text, Qt.ElideRight, left_pad - 10)
                painter.drawText(r.left() + 6, int(y_top), left_pad - 12, row_h, Qt.AlignVCenter | Qt.AlignLeft, elided)

            # background bar
            painter.setPen(Qt.NoPen)
            is_active = (self._hover_index == i) and (self._hover_index >= 0)
            has_hover = self._hover_index >= 0
            bg = QColor(self._grid_color)
            if has_hover and not is_active:
                bg.setAlpha(140)
            painter.setBrush(bg)
            if getattr(self, "_detail_mode", False):
                # In detail mode keep bars visually anchored to the top of each row
                # so the first bar starts much higher (less top air).
                y_bar = float(y_top) + 1.0
            else:
                y_bar = float(y_center - bar_h / 2)
            if is_detail:
                painter.drawRoundedRect(QRectF(bar_left, y_bar, bar_area_w, bar_h), 6, 6)
            else:
                painter.drawRoundedRect(QRectF(plot.left(), y_bar, bar_area_w, bar_h), 6, 6)

            # value bar
            w = (v / max_v) * bar_area_w
            base = QColor(_BAR_PALETTE_HEX[i % len(_BAR_PALETTE_HEX)])

            def hover_fill_hsv(c: QColor) -> QColor:
                h, s, v0, a = c.getHsv()
                if h < 0:
                    return c
                # Keep hue, make it "juicier": more saturation + slightly brighter.
                s2 = min(255, int(s * 1.32) + 8)   # ~+25–35%
                v2 = min(255, int(v0 * 1.07) + 2)  # ~+5–10%
                out = QColor()
                out.setHsv(h, s2, v2, a)
                return out

            fill = QColor(base)
            if has_hover and not is_active:
                fill.setAlpha(190)
            if is_active:
                fill = hover_fill_hsv(fill)
            painter.setBrush(fill)

            if is_active:
                # Stronger contour (no white as the main effect).
                outer_pen = QPen(QColor(fill).darker(160))
                outer_pen.setWidthF(2.0)
                outer_pen.setJoinStyle(Qt.RoundJoin)
                painter.setPen(outer_pen)
            else:
                painter.setPen(Qt.NoPen)
            if is_detail:
                rect = QRectF(bar_left, y_bar, max(2.0, w), bar_h)
            else:
                rect = QRectF(plot.left(), y_bar, max(2.0, w), bar_h)
            if is_active:
                rect = rect.adjusted(-1.0, -1.0, 1.0, 1.0)
            painter.drawRoundedRect(rect, 6, 6)
            painter.setPen(Qt.NoPen)

            # value text
            painter.setFont(value_font)
            painter.setPen(self._text_color)
            if is_detail:
                value_rect = QRectF(float(value_left), float(y_top), float(value_w), float(row_h))
                painter.drawText(value_rect, Qt.AlignVCenter | Qt.AlignLeft, str(int(v)))
            else:
                value_x = int(bar_right + 6)
                painter.drawText(value_x, int(y_center - row_h / 2), right_pad, row_h, Qt.AlignVCenter | Qt.AlignLeft, str(int(v)))

        painter.end()

    def mouseMoveEvent(self, event):  # type: ignore[override]
        if not getattr(self, "_detail_mode", False):
            return super().mouseMoveEvent(event)
        plot = self._last_plot
        if plot is None or self._last_row_h <= 0:
            return super().mouseMoveEvent(event)
        y = float(event.pos().y())
        if y < float(plot.top()) or y > float(plot.bottom()):
            self._set_hover_index(-1)
            return super().mouseMoveEvent(event)
        idx = int((y - float(plot.top())) // float(self._last_row_h))
        self._set_hover_index(idx)
        return super().mouseMoveEvent(event)

    def leaveEvent(self, event):  # type: ignore[override]
        if getattr(self, "_detail_mode", False):
            self._set_hover_index(-1)
        return super().leaveEvent(event)


class _PieChart(QWidget):
    """Compact pie/donut chart with a small legend (statistics cards)."""

    hoverLabelChanged = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._labels: list[str] = []
        self._values: list[int] = []
        self._legend_max_items = 6
        self._detail_mode = False
        self._hover_label: str = ""
        self._hover_index: int = -1
        self._last_cx: float = 0.0
        self._last_cy: float = 0.0
        self._last_outer: QRectF | None = None
        self._last_angles: list[tuple[float, float]] = []  # (start_deg, span_deg) in clockwise-negative logic
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setMouseTracking(True)

    def set_detail_mode(self, enabled: bool) -> None:
        """Detail mode: anchor donut higher, reduce top air."""
        self._detail_mode = bool(enabled)
        self.update()

    def set_hover_label(self, label: str) -> None:
        lab = (label or "").strip()
        if lab == self._hover_label:
            return
        self._hover_label = lab
        try:
            self._hover_index = self._labels.index(lab) if lab else -1
        except ValueError:
            self._hover_index = -1
        self.update()

    def _set_hover_index(self, idx: int) -> None:
        idx = int(idx)
        if idx < 0 or idx >= len(self._labels):
            idx = -1
        if idx == self._hover_index:
            return
        self._hover_index = idx
        self._hover_label = self._labels[idx] if idx >= 0 else ""
        self.hoverLabelChanged.emit(self._hover_label)
        self.update()

    def set_legend_max_items(self, max_items: int) -> None:
        """Сколько строк легенды рисовать (по умолчанию 6 — карточки на главной)."""
        self._legend_max_items = max(1, int(max_items))
        self.update()

    def set_data(self, labels: list[str], values: list[int]):
        self._labels = list(labels or [])
        self._values = [int(v) for v in (values or [])]
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)

        if getattr(self, "_detail_mode", False):
            r = self.rect()
        else:
            r = self.rect().adjusted(2, 2, -2, -2)
        painter.fillRect(r, Qt.transparent)

        if not self._values:
            painter.end()
            return

        total = sum(max(0, int(v)) for v in self._values)
        if total <= 0:
            painter.end()
            return

        # Layout inside widget: donut on the left, legend on the right.
        # Important: legend must be anchored to the *actual* donut outer edge (outer.right()),
        # not to a fixed right column from pie_area.
        is_detail = bool(getattr(self, "_detail_mode", False))
        if is_detail:
            # Detail layout: reserve real legend width, keep donut smaller.
            legend_w = max(220.0, min(300.0, float(r.width()) * 0.34))
            gap = 24.0
            pad = 8.0
            inset = 4.0
        else:
            gap = 6.0
            inset = 6.0
            pad = 8.0
            legend_w = 0.0  # computed from outer edge below

        avail_h = max(10.0, r.height() - pad * 2)

        if is_detail:
            available_for_donut = float(r.width()) - (pad * 2) - float(legend_w) - gap
            donut_size = min(avail_h * 0.82, available_for_donut)
            donut_size = max(160.0, donut_size)
            donut_size *= 0.90
        else:
            # Dashboard: keep a minimum width for legend text, choose donut size from remaining space.
            legend_min_w = 96.0
            donut_size = min(avail_h, max(64.0, r.width() - (pad * 2) - legend_min_w - gap))

        cx = r.left() + pad + inset + donut_size / 2
        if is_detail:
            # Anchor donut higher for detail view.
            cy = r.top() + pad + donut_size / 2
        else:
            cy = r.top() + pad + (avail_h - donut_size) / 2 + donut_size / 2
        size = donut_size
        outer = QRectF(cx - size / 2, cy - size / 2, size, size)
        inner = QRectF(cx - size * 0.33, cy - size * 0.33, size * 0.66, size * 0.66)
        self._last_cx = float(cx)
        self._last_cy = float(cy)
        self._last_outer = outer
        self._last_angles = []

        if is_detail:
            legend_left = float(r.right()) - pad - float(legend_w)
            legend_area = QRectF(legend_left, r.top() + pad, float(legend_w), avail_h)
        else:
            legend_left = float(outer.right()) + gap
            legend_w2 = max(60.0, float(r.right()) - pad - legend_left)
            legend_area = QRectF(legend_left, r.top() + pad, legend_w2, avail_h)

        start_angle = 90.0
        painter.setPen(Qt.NoPen)

        # Percent labels near segments (not in legend)
        pct_font = painter.font()
        pct_font.setFamily("Segoe UI")
        pct_font.setPointSize(11)
        pct_font.setWeight(QFont.DemiBold)
        fm_pct = QFontMetrics(pct_font)

        for i, v in enumerate(self._values):
            vv = max(0, int(v))
            if vv <= 0:
                continue
            span = 360.0 * (vv / total)
            color = QColor(_BAR_PALETTE_HEX[i % len(_BAR_PALETTE_HEX)])
            self._last_angles.append((start_angle, span))

            path = QPainterPath()
            path.moveTo(cx, cy)
            path.arcTo(outer, start_angle, -span)
            path.closeSubpath()

            is_active = (self._hover_index == i) and (self._hover_index >= 0)
            has_hover = self._hover_index >= 0

            def hover_fill_hsv(c: QColor) -> QColor:
                h, s, v0, a = c.getHsv()
                if h < 0:
                    return c
                # Keep hue, make it "juicier": more saturation + slightly brighter.
                s2 = min(255, int(s * 1.32) + 8)   # ~+25–35%
                v2 = min(255, int(v0 * 1.07) + 2)  # ~+5–10%
                out = QColor()
                out.setHsv(h, s2, v2, a)
                return out

            fill = QColor(color)
            if has_hover and not is_active:
                fill.setAlpha(190)
            if is_active:
                fill = hover_fill_hsv(fill)
            painter.setBrush(fill)

            # Optional tiny lift (detail-mode only): 4px radial offset for active sector
            draw_path = path
            if is_active and getattr(self, "_detail_mode", False):
                mid_deg = start_angle - span / 2.0
                rad = math.radians(mid_deg)
                dx = math.cos(rad) * 3.0
                dy = -math.sin(rad) * 3.0
                outer2 = outer.translated(dx, dy)
                cx2 = cx + dx
                cy2 = cy + dy
                draw_path = QPainterPath()
                draw_path.moveTo(cx2, cy2)
                draw_path.arcTo(outer2, start_angle, -span)
                draw_path.closeSubpath()

            if is_active:
                outer_pen = QPen(QColor(fill).darker(160))
                outer_pen.setWidthF(2.0)
                outer_pen.setJoinStyle(Qt.RoundJoin)
                painter.setPen(outer_pen)
            else:
                painter.setPen(Qt.NoPen)
            painter.drawPath(draw_path)
            painter.setPen(Qt.NoPen)

            # label at mid-angle; hide very tiny slices to avoid clutter
            if span >= 10.0:
                pct = int(round((vv / total) * 100)) if total > 0 else 0
                mid_deg = start_angle - span / 2.0
                rad = math.radians(mid_deg)

                outer_r = size / 2.0
                # push outward for smaller slices
                label_r = outer_r * (0.92 if span < 18.0 else 0.80)
                lx = cx + math.cos(rad) * label_r
                ly = cy - math.sin(rad) * label_r

                txt = f"{pct}%"
                tw = fm_pct.horizontalAdvance(txt)
                th = fm_pct.height()
                rect = QRectF(lx - tw / 2 - 4, ly - th / 2 - 2, tw + 8, th + 4)

                painter.save()
                painter.setFont(pct_font)
                # subtle pill to keep readability on any segment color
                painter.setPen(Qt.NoPen)
                painter.setBrush(QColor(255, 255, 255, 210))
                painter.drawRoundedRect(rect, 6, 6)
                painter.setPen(QColor(_C_TEXT))
                painter.drawText(rect, Qt.AlignCenter, txt)
                painter.restore()

            start_angle -= span

        # donut hole
        painter.setBrush(QColor(255, 255, 255, 255))
        painter.drawEllipse(inner)

        # legend
        label_font = painter.font()
        label_font.setFamily("Segoe UI")
        label_font.setPointSize(10 if is_detail else 11)
        label_font.setWeight(QFont.Medium)
        painter.setFont(label_font)
        fm = QFontMetrics(painter.font())

        value_font = painter.font()
        value_font.setFamily("Segoe UI")
        value_font.setPointSize(12)
        value_font.setWeight(QFont.DemiBold)

        leg_n = min(len(self._labels), self._legend_max_items)
        if is_detail:
            # Detail: make legend rows compact; keep extra space below.
            row_h = max(16, int(fm.height()) + 4)
        else:
            row_h = max(18, int(legend_area.height() / max(1, leg_n)))
        y0 = int(legend_area.top())
        dot_r = 5

        for i, (lab, v) in enumerate(zip(self._labels, self._values)):
            if i >= self._legend_max_items:
                break
            vv = int(v)
            y = y0 + i * row_h
            color = QColor(_BAR_PALETTE_HEX[i % len(_BAR_PALETTE_HEX)])

            painter.setPen(Qt.NoPen)
            painter.setBrush(color)
            painter.drawEllipse(int(legend_area.left()), int(y + row_h / 2 - dot_r), dot_r * 2, dot_r * 2)

            # legend: dot -> label (left) + value+percent (right)
            base_x = int(legend_area.left()) + dot_r * 2 + 8
            right_pad = 8
            if is_detail:
                value_txt = f"{vv}"
            else:
                pct = int(round((max(0, vv) / total) * 100)) if total > 0 else 0
                value_txt = f"{vv} · {pct}%"

            painter.setFont(value_font)
            fm_v = QFontMetrics(painter.font())
            value_w = max(72, fm_v.horizontalAdvance(value_txt) + 6)
            value_x = int(legend_area.right()) - right_pad - value_w

            painter.setPen(QColor(_C_TEXT))
            painter.drawText(
                value_x,
                y,
                value_w,
                row_h,
                (Qt.AlignVCenter | Qt.AlignLeft) if is_detail else (Qt.AlignVCenter | Qt.AlignRight),
                value_txt,
            )

            painter.setPen(QColor(_C_SUB))
            painter.setFont(label_font)
            max_w = max(10, value_x - 10 - base_x)
            elided = fm.elidedText(str(lab), Qt.ElideRight, int(max_w))
            painter.drawText(base_x, y, int(max_w), row_h, Qt.AlignVCenter | Qt.AlignLeft, elided)

        painter.end()

    def mouseMoveEvent(self, event):  # type: ignore[override]
        if not getattr(self, "_detail_mode", False):
            return super().mouseMoveEvent(event)
        outer = self._last_outer
        if outer is None or not self._last_angles:
            return super().mouseMoveEvent(event)
        x = float(event.pos().x())
        y = float(event.pos().y())
        cx = float(self._last_cx)
        cy = float(self._last_cy)
        dx = x - cx
        dy = cy - y  # screen y down
        r2 = dx * dx + dy * dy
        outer_r = float(outer.width()) / 2.0
        inner_r = outer_r * 0.33
        if r2 < inner_r * inner_r or r2 > outer_r * outer_r:
            self._set_hover_index(-1)
            return super().mouseMoveEvent(event)
        # angle in degrees: 0 at +x, counter-clockwise positive
        ang = math.degrees(math.atan2(dy, dx))
        if ang < 0:
            ang += 360.0
        # Our chart uses start_angle decreasing by span (clockwise drawing).
        # Convert to "from top" logic:
        # start_angle is in Qt degrees with 0 at 3 o'clock, positive CCW; we used arcTo with -span.
        # Here we match by checking if ang lies within [start-span, start] in circular sense.
        for i, (start, span) in enumerate(self._last_angles):
            a0 = (start - span) % 360.0
            a1 = start % 360.0
            if a0 <= a1:
                hit = (a0 <= ang <= a1)
            else:
                hit = (ang >= a0 or ang <= a1)
            if hit:
                self._set_hover_index(i)
                break
        else:
            self._set_hover_index(-1)
        return super().mouseMoveEvent(event)

    def leaveEvent(self, event):  # type: ignore[override]
        if getattr(self, "_detail_mode", False):
            self._set_hover_index(-1)
        return super().leaveEvent(event)


class _PeriodDateTimeEdit(QDateTimeEdit):
    """
    QDateTimeEdit для диалога периода: выпадающий QCalendarWidget центрируется под полем,
    внешний popup-контейнер делается прозрачным и обрезается по скруглению (без серых углов).
    """

    _POPUP_RADIUS = 16
    _POPUP_GAP_Y = 4

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setCalendarPopup(True)
        self.setButtonSymbols(QAbstractSpinBox.NoButtons)
        self._calendar_popup_shell_hooked = False
        self.calendarWidget().installEventFilter(self)

    def eventFilter(self, watched: QWidget, event: QEvent) -> bool:  # type: ignore[override]
        cal = self.calendarWidget()
        if cal is not None and watched is cal and event.type() == QEvent.Show:
            QTimer.singleShot(0, self._sync_calendar_popup)
        pop = cal.parentWidget() if cal is not None else None
        if (
            pop is not None
            and pop is not self
            and watched is pop
            and event.type() == QEvent.Resize
        ):
            self._apply_popup_round_mask(pop)
        return False

    def _sync_calendar_popup(self) -> None:
        cal = self.calendarWidget()
        pop = cal.parentWidget()
        if pop is None or pop is self:
            return
        if not self._calendar_popup_shell_hooked:
            pop.installEventFilter(self)
            self._calendar_popup_shell_hooked = True
        self._polish_popup_shell(pop)
        self._position_popup_under_field(pop)
        self._apply_popup_round_mask(pop)
        self._apply_calendar_soft_shadow(cal)

    def _apply_calendar_soft_shadow(self, cal: QWidget) -> None:
        """Мягкая равномерная тень (blur, offset 0), без смещения только вправо-вниз."""
        if cal.graphicsEffect() is not None:
            return
        eff = QGraphicsDropShadowEffect(cal)
        eff.setBlurRadius(48)
        eff.setOffset(0, 0)
        eff.setColor(QColor(15, 23, 42, 46))
        cal.setGraphicsEffect(eff)

    def _polish_popup_shell(self, pop: QWidget) -> None:
        # Непрозрачный shell: рамка и фон рисуются самим контейнером по всему периметру (как у QComboBox popup в main.py).
        pop.setObjectName("PeriodCalendarPopupShell")
        pop.setAttribute(Qt.WA_TranslucentBackground, False)
        pop.setAutoFillBackground(True)
        r = int(self._POPUP_RADIUS)
        pop.setStyleSheet(
            f"#PeriodCalendarPopupShell {{"
            f"background-color: #ffffff;"
            f"border: 2px solid #cbd5e1;"
            f"border-radius: {r}px;"
            f"}}"
        )
        lay = pop.layout()
        if lay is not None:
            lay.setContentsMargins(0, 0, 0, 0)

    def _position_popup_under_field(self, pop: QWidget) -> None:
        top_left = self.mapToGlobal(QPoint(0, 0))
        pw, ph = pop.width(), pop.height()
        x = top_left.x() + max(0, (self.width() - pw) // 2)
        y = top_left.y() + self.height() + self._POPUP_GAP_Y

        screen = QApplication.screenAt(QPoint(x + min(pw, self.width()) // 2, y + 4))
        if screen is not None:
            ag = screen.availableGeometry()
            if x + pw > ag.right():
                x = max(ag.left(), ag.right() - pw + 1)
            if x < ag.left():
                x = ag.left()
            if y + ph > ag.bottom():
                y = max(ag.top(), top_left.y() - ph - self._POPUP_GAP_Y)
            if y < ag.top():
                y = ag.top()

        pop.move(x, y)

    def _apply_popup_round_mask(self, pop: QWidget) -> None:
        r = float(self._POPUP_RADIUS)
        path = QPainterPath()
        path.addRoundedRect(QRectF(pop.rect()), r, r)
        pop.setMask(QRegion(path.toFillPolygon().toPolygon()))


class StatisticsPage(QWidget):
    """Сигнал: True — открыт встроенный detail drill-down, False — главный dashboard статистики."""
    detail_mode_changed = pyqtSignal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("StatisticsPage")
        self.setStyleSheet(f"background: #f8fafc; font-family: {_STATS_FONT_FAMILY};")

        self._period = "day"
        self._custom_start_dt: datetime | None = None
        self._custom_end_dt: datetime | None = None
        self._all_records: list[PrintRecord] = []
        self._archive_root: str | None = None
        self._dashboard_records: list[PrintRecord] = []
        self._detail_mode = False
        self._current_detail_type = "products"

        root = QVBoxLayout(self)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(14)

        self.header = QLabel("Статистика — День")
        self.header.setStyleSheet(
            f"font-family: {_STATS_FONT_FAMILY}; font-size: 21px; font-weight: 650; line-height: 24px; "
            f"color: {_C_TITLE}; background: transparent;"
        )

        # content stack:
        # - page 0: empty-state
        # - page 1: main dashboard
        # - page 2: встроенный detail (drill-down)
        self.content_stack = QStackedLayout()
        self.content_stack.setContentsMargins(0, 0, 0, 0)
        root.addLayout(self.content_stack, 1)

        self.empty_state_page = self._build_empty_state_page()
        self.dashboard_widget = self._build_dashboard()
        self.detail_view = StatisticsDetailView(self)

        self.content_stack.addWidget(self.empty_state_page)
        self.content_stack.addWidget(self.dashboard_widget)
        self.content_stack.addWidget(self.detail_view)
        # Dashboard-структура должна быть видна всегда; empty-state показываем внутри dashboard ниже основных блоков.
        self.content_stack.setCurrentWidget(self.dashboard_widget)

        self.refresh_from_archive()

    # ----- Public API -----
    def _get_period_subtitle(self) -> str:
        now = datetime.now()
        if self._period == "day":
            return f"День · {now.strftime('%d.%m.%Y')}"
        if self._period == "week":
            start = (now - timedelta(days=6)).replace(hour=0, minute=0, second=0, microsecond=0)
            return f"Последние 7 дней · {start.strftime('%d.%m.%Y')} — {now.strftime('%d.%m.%Y')}"
        if self._period == "month":
            d0 = now.date() - timedelta(days=29)
            start = datetime.combine(d0, datetime.min.time())
            return f"Последние 30 дней · {start.strftime('%d.%m.%Y')} — {now.strftime('%d.%m.%Y')}"
        if self._period == "custom" and self._custom_start_dt and self._custom_end_dt:
            return (
                f"{self._custom_start_dt.strftime('%d.%m.%Y %H:%M')} — "
                f"{self._custom_end_dt.strftime('%d.%m.%Y %H:%M')}"
            )
        return "Период"

    def _open_statistics_detail(self, detail_type: str) -> None:
        if not self._dashboard_records:
            return
        self._current_detail_type = (detail_type or "").strip().lower()
        self._detail_mode = True
        self.detail_view.populate(
            self._current_detail_type,
            self._dashboard_records,
            self._get_period_subtitle(),
        )
        self.header.hide()
        try:
            if getattr(self, "_stats_header_row", None) is not None:
                self._stats_header_row.hide()
        except Exception:
            pass
        try:
            if getattr(self, "print_reports_btn", None) is not None:
                self.print_reports_btn.hide()
        except Exception:
            pass
        self.content_stack.setCurrentWidget(self.detail_view)
        self.detail_mode_changed.emit(True)

    def leave_statistics_detail(self) -> None:
        if not self._detail_mode:
            return
        self._detail_mode = False
        self.header.show()
        try:
            if getattr(self, "_stats_header_row", None) is not None:
                self._stats_header_row.show()
        except Exception:
            pass
        try:
            if getattr(self, "print_reports_btn", None) is not None:
                self.print_reports_btn.show()
        except Exception:
            pass
        self.content_stack.setCurrentWidget(self.dashboard_widget)
        self.detail_mode_changed.emit(False)

    def _on_print_reports_clicked(self) -> None:
        from statistics_reports_printing import run_print_reports_flow

        # При нулевой статистике: кнопка выглядит disabled, но остаётся кликабельной и показывает подсказку.
        if not (self._dashboard_records or []):
            QMessageBox.information(
                self,
                "Печать отчётов",
                "Для печати отчётов необходимо сначала напечатать этикетки или добавить источники данных.",
            )
            return

        try:
            p = (self._period or "day").strip().lower()
        except Exception:
            p = "day"
        if p == "day":
            period_label = "День"
        elif p == "week":
            period_label = "7 дней"
        elif p == "month":
            period_label = "30 дней"
        else:
            period_label = "Период"

        run_print_reports_flow(
            period_label=period_label,
            period_subtitle=self._get_period_subtitle(),
            records=list(self._dashboard_records or []),
            parent=self,
        )

    def set_period(self, period: str):
        period = (period or "day").strip().lower()
        if period not in ("day", "week", "month", "custom"):
            return
        self._period = period
        if period == "custom":
            self.header.setText("Статистика — Период")
        else:
            mapping = {"day": "День", "week": "7 дней", "month": "30 дней"}
            self.header.setText(f"Статистика — {mapping.get(period, period)}")
        self.refresh_dashboard()

    def open_custom_period_dialog(self, parent: QWidget | None = None) -> bool:
        dlg = QDialog(parent)
        dlg.setWindowFlags(dlg.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        dlg.setWindowTitle(" ")
        dlg.setModal(True)
        dlg.setMinimumSize(720, 340)
        dlg.resize(800, 365)

        dlg.setStyleSheet(
            "QDialog {"
            "background: #f6f7f9;"
            "}"
        )

        root = QVBoxLayout(dlg)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(0)

        card = QFrame()
        card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        card.setObjectName("CustomPeriodCard")
        card.setStyleSheet(
            "#CustomPeriodCard {"
            "background: #ffffff;"
            "border: 1px solid #e5e7eb;"
            "border-radius: 16px;"
            "}"
        )
        shadow = QGraphicsDropShadowEffect(card)
        shadow.setBlurRadius(28)
        shadow.setOffset(0, 10)
        shadow.setColor(QColor(15, 23, 42, 28))
        card.setGraphicsEffect(shadow)

        lay = QVBoxLayout(card)
        lay.setContentsMargins(24, 18, 24, 18)
        lay.setSpacing(0)

        _label_sheet = "background: transparent; border: none; margin: 0; padding: 0;"
        title = QLabel("Период статистики")
        title.setAutoFillBackground(False)
        title.setAttribute(Qt.WA_TranslucentBackground, True)
        title.setStyleSheet(
            _label_sheet
            + f'font-family: {_STATS_FONT_FAMILY}; font-size: 27px; font-weight: 650; color: {_C_TITLE};'
        )
        subtitle = QLabel("Выберите диапазон дат для анализа")
        subtitle.setAutoFillBackground(False)
        subtitle.setAttribute(Qt.WA_TranslucentBackground, True)
        subtitle.setStyleSheet(
            _label_sheet
            + f'font-family: {_STATS_FONT_FAMILY}; font-size: 17px; font-weight: 500; color: {_C_SUB};'
        )
        subtitle.setWordWrap(True)
        lay.addWidget(title)
        lay.addSpacing(4)
        lay.addWidget(subtitle)
        lay.addSpacing(14)

        start_edit = _PeriodDateTimeEdit()
        end_edit = _PeriodDateTimeEdit()
        _period_dt_combo_btn = resource_path("assets/combo-btn.svg").replace("\\", "/")
        for w in (start_edit, end_edit):
            w.setLocale(QLocale(QLocale.Russian, QLocale.Russia))
            w.setDisplayFormat("d MMM yyyy HH:mm")
            w.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            w.setMinimumHeight(54)
            w.setStyleSheet(
                "QDateTimeEdit {"
                "background: #ffffff;"
                "border: 1px solid #e5e7eb;"
                "border-radius: 14px;"
                "padding: 12px 48px 12px 18px;"
                f"color: {_C_TEXT};"
                f"font-family: {_STATS_FONT_FAMILY};"
                "font-size: 18px;"
                "font-weight: 550;"
                "}"
                "QDateTimeEdit:focus {"
                "border: 1px solid #2563eb;"
                "}"
                "QDateTimeEdit::drop-down {"
                "subcontrol-origin: padding;"
                "subcontrol-position: center right;"
                "width: 36px;"
                "border: none;"
                "background: transparent;"
                "margin-right: 4px;"
                f"image: url({_period_dt_combo_btn});"
                "}"
                "QDateTimeEdit::down-arrow {"
                "image: none;"
                "width: 0px;"
                "height: 0px;"
                "}"
                "QDateTimeEdit::up-button, QDateTimeEdit::down-button {"
                "width: 0px;"
                "height: 0px;"
                "border: none;"
                "padding: 0px;"
                "}"
            )
            cal = w.calendarWidget()
            cal.setStyleSheet(_CUSTOM_PERIOD_CALENDAR_QSS)
            cal.setLocale(QLocale(QLocale.Russian, QLocale.Russia))
            cal.setMinimumSize(440, 400)
            _install_calendar_day_hover_delegate(cal)

        now = datetime.now()
        if self._custom_start_dt and self._custom_end_dt:
            start_edit.setDateTime(
                QDateTime.fromString(self._custom_start_dt.strftime("%d.%m.%Y %H:%M"), "dd.MM.yyyy HH:mm")
            )
            end_edit.setDateTime(
                QDateTime.fromString(self._custom_end_dt.strftime("%d.%m.%Y %H:%M"), "dd.MM.yyyy HH:mm")
            )
        else:
            start_edit.setDateTime(
                QDateTime.fromString(
                    now.replace(hour=0, minute=0, second=0, microsecond=0).strftime("%d.%m.%Y %H:%M"),
                    "dd.MM.yyyy HH:mm",
                )
            )
            end_edit.setDateTime(QDateTime.fromString(now.strftime("%d.%m.%Y %H:%M"), "dd.MM.yyyy HH:mm"))

        range_row = QHBoxLayout()
        range_row.setSpacing(12)
        range_row.addWidget(start_edit, 1)

        dash = QLabel("—")
        dash.setAutoFillBackground(False)
        dash.setAttribute(Qt.WA_TranslucentBackground, True)
        dash.setStyleSheet(
            "background: transparent; border: none; margin: 0;"
            + f'font-family: {_STATS_FONT_FAMILY}; font-size: 23px; font-weight: 600; color: {_C_SUB}; padding: 0 6px;'
        )
        dash.setAlignment(Qt.AlignCenter)
        range_row.addWidget(dash, 0, Qt.AlignVCenter)

        range_row.addWidget(end_edit, 1)
        lay.addLayout(range_row)

        lay.addSpacing(12)

        preview = QLabel("")
        preview.setWordWrap(True)
        preview.setAutoFillBackground(False)
        preview.setAttribute(Qt.WA_TranslucentBackground, True)
        preview.setStyleSheet(
            _label_sheet
            + f'font-family: {_STATS_FONT_FAMILY}; font-size: 16px; font-weight: 500; color: {_C_SUB};'
        )
        lay.addWidget(preview)

        lay.addSpacing(12)

        def fmt_preview_line(s: datetime, e: datetime) -> str:
            return f"{s.strftime('%d.%m.%Y %H:%M')} — {e.strftime('%d.%m.%Y %H:%M')}"

        def update_preview():
            s = start_edit.dateTime().toPyDateTime()
            e = end_edit.dateTime().toPyDateTime()
            if s > e:
                preview.setText("Диапазон некорректен: начало позже конца")
                return
            delta = e - s
            hours = int(delta.total_seconds() // 3600)
            days = delta.days
            if hours < 48:
                hint = f"{hours} ч"
            else:
                hint = f"{days} дн."
            preview.setText(f"Выбрано: {hint} · {fmt_preview_line(s, e)}")

        start_edit.dateTimeChanged.connect(lambda _dt: update_preview())
        end_edit.dateTimeChanged.connect(lambda _dt: update_preview())
        update_preview()

        actions = QHBoxLayout()
        actions.setContentsMargins(0, 2, 0, 0)
        actions.setSpacing(12)
        actions.addStretch(1)

        cancel_btn = QPushButton("Отмена")
        cancel_btn.setCursor(Qt.PointingHandCursor)
        cancel_btn.setStyleSheet(
            "QPushButton {"
            "background: #ffffff;"
            "border: 1px solid #e5e7eb;"
            "border-radius: 14px;"
            "padding: 13px 24px;"
            "min-height: 50px;"
            f"font-family: {_STATS_FONT_FAMILY};"
            "font-size: 18px;"
            f"font-weight: 600; color: {_C_TEXT};"
            "}"
            "QPushButton:hover { background: #f8fafc; }"
        )

        apply_btn = QPushButton("Применить")
        apply_btn.setCursor(Qt.PointingHandCursor)
        apply_btn.setStyleSheet(
            "QPushButton {"
            "background: #facc15;"
            "border: 1px solid #eab308;"
            "border-radius: 14px;"
            "padding: 13px 26px;"
            "min-height: 50px;"
            f"font-family: {_STATS_FONT_FAMILY};"
            "font-size: 18px;"
            "font-weight: 800;"
            "color: #78350f;"
            "}"
            "QPushButton:hover { background: #fde047; }"
            "QPushButton:pressed { background: #f59e0b; color: #451a03; border-color: #d97706; }"
        )

        actions.addWidget(cancel_btn, 0)
        actions.addWidget(apply_btn, 0)
        lay.addLayout(actions)

        root.addWidget(card, 1)

        accepted = {"ok": False}

        def on_cancel():
            dlg.reject()

        def on_apply():
            s = start_edit.dateTime().toPyDateTime()
            e = end_edit.dateTime().toPyDateTime()
            if s > e:
                QMessageBox.warning(parent or self, "Период", "Начало периода не может быть позже конца.")
                return
            accepted["ok"] = True
            dlg.accept()

        cancel_btn.clicked.connect(on_cancel)
        apply_btn.clicked.connect(on_apply)

        if dlg.exec_() != QDialog.Accepted or not accepted["ok"]:
            return False

        s = start_edit.dateTime().toPyDateTime()
        e = end_edit.dateTime().toPyDateTime()
        if s > e:
            QMessageBox.warning(parent or self, "Период", "Начало периода не может быть позже конца.")
            return False

        self._custom_start_dt = s
        self._custom_end_dt = e
        self._period = "custom"
        self.header.setText("Статистика — Период")
        self.refresh_dashboard()
        return True

    def set_archive_root(self, archive_root: str | None):
        self._archive_root = archive_root or None

    def refresh_from_archive(self):
        # Stage 2: prefer local JSONL journal; fallback to legacy txt-archive.
        records: list[PrintRecord] = []
        try:
            from stats_store import read_local_entries as _read_local_entries  # local-only source
            from stats_exchange import read_imported_entries as _read_imported_entries

            entries = list(_read_local_entries() or [])
            try:
                entries.extend(_read_imported_entries() or [])
            except Exception:
                pass

            # Deduplicate by record_id: local entries win, imported duplicates are ignored.
            uniq: list[dict] = []
            seen_ids: set[str] = set()
            for e in entries:
                if not isinstance(e, dict):
                    continue
                rid = str(e.get("record_id") or "").strip()
                if rid:
                    if rid in seen_ids:
                        continue
                    seen_ids.add(rid)
                uniq.append(e)
            entries = uniq

            if entries:
                for e in entries:
                    if not isinstance(e, dict):
                        continue
                    try:
                        ts = float(e.get("ts") or 0.0)
                    except Exception:
                        ts = 0.0
                    if ts <= 0:
                        continue
                    try:
                        dt = datetime.fromtimestamp(ts)
                    except Exception:
                        continue

                    product = str(e.get("product") or "").strip() or "—"
                    made_by = str(e.get("made_by") or "").strip() or "—"
                    workshop = str(e.get("workshop") or "").strip() or "—"
                    try:
                        copies = int(e.get("copies") or 1)
                    except Exception:
                        copies = 1
                    if copies <= 0:
                        copies = 1

                    records.append(PrintRecord(dt=dt, product=product, made_by=made_by, workshop=workshop, copies=copies))
        except Exception:
            records = []

        if not records:
            if not self._archive_root:
                self._all_records = []
                self.refresh_dashboard()
                return
            records = load_print_records_from_archive(self._archive_root)

        self._all_records = records
        self.refresh_dashboard()

    def _period_ts_range(self) -> tuple[float, float]:
        """Current UI period as [from_ts, to_ts]."""
        now = datetime.now()
        p = (self._period or "day").strip().lower()
        if p == "day":
            start = now.replace(hour=0, minute=0, second=0, microsecond=0)
            end = start + timedelta(days=1)
        elif p == "week":
            end = now
            start = (end - timedelta(days=6)).replace(hour=0, minute=0, second=0, microsecond=0)
        elif p == "month":
            end = now
            d0 = now.date() - timedelta(days=29)
            start = datetime.combine(d0, datetime.min.time())
        elif p == "custom" and self._custom_start_dt and self._custom_end_dt:
            start = self._custom_start_dt
            end = self._custom_end_dt
        else:
            start = now.replace(hour=0, minute=0, second=0, microsecond=0)
            end = now
        return float(start.timestamp()), float(end.timestamp())

    def _export_stats_package(self) -> None:
        """Stage 3: export local JSONL entries for current period to a portable zip package."""
        try:
            from stats_store import read_local_entries as _read_local_entries, ensure_station_identity as _ensure_station_identity
            from stats_exchange import export_package as _export_package
        except Exception as e:
            QMessageBox.warning(self, "Экспорт", f"Не удалось подготовить экспорт:\n{e}")
            return

        entries = []
        try:
            entries = _read_local_entries() or []
        except Exception:
            entries = []

        from_ts, to_ts = self._period_ts_range()
        station_uuid, station_label = ("", "")
        try:
            station_uuid, station_label = _ensure_station_identity()
        except Exception:
            station_uuid, station_label = ("", "")

        if not entries:
            QMessageBox.information(self, "Экспорт", "Нет локальных записей для экспорта (stats_journal.jsonl пуст).")
            return

        default_name = f"mirlis_stats_{int(from_ts)}_{int(to_ts)}.zip"
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Экспорт статистики",
            default_name,
            "ZIP (*.zip)",
        )
        if not save_path:
            return
        if not save_path.lower().endswith(".zip"):
            save_path = save_path + ".zip"

        try:
            n = _export_package(
                entries,
                from_ts,
                to_ts,
                save_path,
                station_uuid,
                station_label,
                app_version="",
            )
        except Exception as e:
            QMessageBox.warning(self, "Экспорт", f"Не удалось сохранить файл:\n{e}")
            return

        if n <= 0:
            QMessageBox.information(self, "Экспорт", "За выбранный период нет записей. Файл не создан.")
            return

        QMessageBox.information(self, "Экспорт", f"Экспортировано записей: {n}\nФайл: {save_path}")

    def _open_sources_dialog(self) -> None:
        try:
            from stats_sources_dialog import StatsSourcesDialog
        except Exception as e:
            QMessageBox.warning(self, "Источники", f"Не удалось открыть список источников:\n{e}")
            return

        dlg = StatsSourcesDialog(self, on_sources_changed=self.refresh_from_archive)
        dlg.exec_()

    def _export_stats_excel(self) -> None:
        """Stage 6: export current period stats to .xlsx (human-friendly)."""
        try:
            from stats_store import read_local_entries as _read_local_entries
            from stats_exchange import read_imported_entries as _read_imported_entries, export_to_excel as _export_to_excel
        except Exception as e:
            QMessageBox.warning(self, "Excel", f"Не удалось подготовить экспорт:\n{e}")
            return

        entries = list(_read_local_entries() or [])
        try:
            entries.extend(_read_imported_entries() or [])
        except Exception:
            pass

        # Deduplicate by record_id (same logic as refresh_from_archive)
        uniq: list[dict] = []
        seen_ids: set[str] = set()
        for e in entries:
            if not isinstance(e, dict):
                continue
            rid = str(e.get("record_id") or "").strip()
            if rid:
                if rid in seen_ids:
                    continue
                seen_ids.add(rid)
            uniq.append(e)
        entries = uniq

        from_ts, to_ts = self._period_ts_range()

        default_name = f"mirlis_stats_{int(from_ts)}_{int(to_ts)}.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Экспорт статистики в Excel",
            default_name,
            "Excel Files (*.xlsx)",
        )
        if not save_path:
            return
        if not save_path.lower().endswith(".xlsx"):
            save_path = save_path + ".xlsx"

        try:
            n = _export_to_excel(entries, from_ts, to_ts, save_path)
        except Exception as e:
            QMessageBox.warning(self, "Excel", f"Не удалось сохранить файл:\n{e}")
            return

        if n <= 0:
            QMessageBox.information(self, "Excel", "За выбранный период нет записей. Файл не создан.")
            return

        QMessageBox.information(self, "Excel", f"Экспортировано записей: {n}\nФайл: {save_path}")

    def _import_stats_package(self) -> None:
        """Stage 4: import a previously exported stats package zip."""
        try:
            from stats_exchange import (
                inspect_package as _inspect_package,
                import_package as _import_package,
                list_imported_sources as _list_imported_sources,
            )
            from stats_store import ensure_station_identity as _ensure_station_identity
        except Exception as e:
            QMessageBox.warning(self, "Импорт", f"Не удалось подготовить импорт:\n{e}")
            return

        zip_path, _ = QFileDialog.getOpenFileName(
            self,
            "Импорт пакета статистики",
            "",
            "ZIP (*.zip)",
        )
        if not zip_path:
            return

        try:
            m = _inspect_package(zip_path)
        except Exception as e:
            QMessageBox.warning(self, "Импорт", f"Не удалось прочитать пакет:\n{e}")
            return

        # Precheck BEFORE confirmation:
        # 1) already imported by package_id
        pkg_id = str(m.get("package_id") or "").strip()
        if pkg_id:
            try:
                for it in (_list_imported_sources() or []):
                    if str(it.get("package_id") or "").strip() == pkg_id:
                        QMessageBox.information(self, "Импорт", "Этот пакет уже импортирован (package_id совпадает).")
                        return
            except Exception:
                pass

        # 2) self-import by station_uuid
        try:
            local_station_uuid, _local_station_label = _ensure_station_identity()
            src_station_uuid = str(m.get("source_station_uuid") or "").strip()
            if src_station_uuid and str(local_station_uuid or "").strip() == src_station_uuid:
                QMessageBox.information(
                    self,
                    "Импорт",
                    "Этот пакет был создан на текущей станции/этом ПК. Импорт приведёт к дублям и не требуется.",
                )
                return
        except Exception:
            # If we cannot determine identity, don't block confirmation/import here.
            pass

        src_label = str(m.get("source_station_label") or "—")
        src_uuid = str(m.get("source_station_uuid") or "—")
        try:
            pf = float(m.get("period_from") or 0.0)
            pt = float(m.get("period_to") or 0.0)
        except Exception:
            pf, pt = 0.0, 0.0
        try:
            rc = int(m.get("record_count") or 0)
        except Exception:
            rc = 0

        msg = (
            f"Источник: {src_label}\n"
            f"UUID: {src_uuid}\n"
            f"Период: {datetime.fromtimestamp(pf).strftime('%d.%m.%Y %H:%M')} — {datetime.fromtimestamp(pt).strftime('%d.%m.%Y %H:%M')}\n"
            f"Записей: {rc}\n\n"
            "Импортировать этот пакет?"
        )
        if QMessageBox.question(self, "Импорт", msg, QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return

        try:
            rec = _import_package(zip_path)
        except Exception as e:
            QMessageBox.information(self, "Импорт", str(e))
            return

        QMessageBox.information(
            self,
            "Импорт",
            f"Пакет импортирован.\nЗаписей: {rec.get('record_count', 0)}",
        )
        self.refresh_from_archive()

    def set_has_data(self, has_data: bool):
        # kept for backwards-compat calls from main.py; actual state is driven by records
        # Не переключаем экран на отдельную empty-state страницу:
        # dashboard остаётся видимым, а empty-state показывается ниже основных карточек.
        try:
            self.refresh_dashboard()
        except Exception:
            pass

    # ----- UI builders -----
    def _build_empty_state(self) -> QWidget:
        w = QWidget()
        w.setStyleSheet("background: transparent; border: none;")
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(10)
        lay.setAlignment(Qt.AlignHCenter)

        img = QLabel()
        img.setAlignment(Qt.AlignHCenter)
        img.setStyleSheet("background: transparent; border: none;")
        pix = QPixmap(SAD_HERO_PATH)
        if not pix.isNull():
            pix = pix.scaled(300, 300, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            img.setPixmap(pix)
        lay.addWidget(img, 0, Qt.AlignHCenter)

        title = QLabel("Пока нет данных для статистики")
        title.setAlignment(Qt.AlignHCenter)
        title.setStyleSheet(
            f"font-family: {_STATS_FONT_FAMILY}; font-size: 15px; font-weight: 600; line-height: 20px; "
            f"color: {_C_TITLE}; background: transparent; border: none;"
        )
        lay.addWidget(title, 0, Qt.AlignHCenter)

        sub = QLabel("Напечатайте первую этикетку, чтобы здесь появились диаграммы и сводка")
        sub.setAlignment(Qt.AlignHCenter)
        sub.setWordWrap(True)
        sub.setStyleSheet(
            f"font-family: {_STATS_FONT_FAMILY}; font-size: 11px; font-weight: 500; line-height: 14px; "
            f"color: {_C_SUB}; background: transparent; border: none;"
        )
        lay.addWidget(sub, 0, Qt.AlignHCenter)

        outer = QVBoxLayout()
        outer.setContentsMargins(0, 0, 0, 0)
        # Поднимаем пустое состояние выше внутри контейнера:
        # верхний stretch меньше, нижний — больше, при этом весь блок остаётся единым и центрируется по горизонтали.
        outer.addStretch(0)
        outer.addWidget(w, 0, Qt.AlignHCenter)
        outer.addStretch(2)
        container = QWidget()
        container.setLayout(outer)
        container.setStyleSheet("background: transparent; border: none;")
        return container

    def _build_empty_state_page(self) -> QWidget:
        surface = QFrame()
        surface.setStyleSheet("background: #ffffff; border: 1px solid #e5e7eb; border-radius: 20px;")
        # Пустое состояние — крупный нижний блок. Делаем его выше примерно на 20%,
        # не масштабируя контент: центрирование обеспечивается layout'ом внутри.
        try:
            surface.setMinimumHeight(550)  # было 500 → стало ещё выше (~+10%)
        except Exception:
            pass
        lay = QVBoxLayout(surface)
        lay.setContentsMargins(16, 16, 16, 16)
        lay.setSpacing(0)

        self.empty_state_widget = self._build_empty_state()
        lay.addWidget(self.empty_state_widget, 1)
        return surface

    def _build_dashboard(self) -> QWidget:
        w = QWidget()
        w.setStyleSheet("background: transparent; border: none;")
        grid = QGridLayout(w)
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(14)
        grid.setVerticalSpacing(14)
        # Контроль пропорций по вертикали: средний ряд получает больше свободного места,
        # нижний ряд не должен схлопываться из-за жёстких minHeight сверху.
        grid.setRowStretch(0, 0)
        grid.setRowStretch(1, 0)
        grid.setRowStretch(2, 1)
        grid.setRowStretch(3, 1)
        grid.setRowStretch(4, 0)

        # 12-column grid:
        # top row -> 4 / 4 / 2 / 2
        # middle row -> 12
        # bottom row -> 4 / 4 / 4
        for c in range(12):
            grid.setColumnStretch(c, 1)

        # Row 0: header row (title centered, print button right)
        self._stats_header_row = QWidget()
        self._stats_header_row.setStyleSheet("background: transparent; border: none;")
        header_row_lay = QGridLayout(self._stats_header_row)
        header_row_lay.setContentsMargins(0, 0, 0, 0)
        header_row_lay.setHorizontalSpacing(0)
        header_row_lay.setVerticalSpacing(0)
        header_row_lay.setColumnStretch(0, 1)
        header_row_lay.setColumnStretch(1, 0)
        header_row_lay.setColumnStretch(2, 1)

        self.print_reports_btn = QPushButton("Печать отчётов")
        self.print_reports_btn.setCursor(Qt.PointingHandCursor)
        self.print_reports_btn.setFixedHeight(_DASHBOARD_IO_TILE_BTN_HEIGHT_PX)
        self.print_reports_btn.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.print_reports_btn.setStyleSheet(_dashboard_io_tile_button_stylesheet())
        try:
            pm = QPixmap(PRINT_REPORTS_BTN_ICON_PATH)
            if not pm.isNull():
                _isz = _DASHBOARD_IO_TILE_ICON_PX
                pm = pm.scaled(_isz, _isz, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                self.print_reports_btn.setIcon(QIcon(pm))
                self.print_reports_btn.setIconSize(QSize(_isz, _isz))
        except Exception:
            pass
        try:
            self.print_reports_btn.clicked.connect(self._on_print_reports_clicked)
        except Exception:
            pass

        header_row_lay.addWidget(self.header, 0, 1, Qt.AlignHCenter | Qt.AlignVCenter)
        header_row_lay.addWidget(self.print_reports_btn, 0, 2, Qt.AlignRight | Qt.AlignVCenter)
        grid.addWidget(self._stats_header_row, 0, 0, 1, 12, Qt.AlignBottom)

        # KPI row (row 1)
        self.kpi_labels = _KpiCard("Этикеток", icon_path=ROLL_PATH)
        self.kpi_ops = _KpiCard("Операций", icon_path=PRINTER_PATH)

        # Import/Export module (Stage 6 UI polish)
        self.io_card = _Card("Импорт и экспорт")
        # Prevent any inner "boxed" border around the card body (only keep outer card border).
        try:
            self.io_card.body.setStyleSheet("background: transparent; border: none;")
        except Exception:
            pass
        # Make this card more compact than a regular _Card (only for io_card).
        try:
            if self.io_card.layout() is not None:
                self.io_card.layout().setContentsMargins(14, 10, 14, 10)
                self.io_card.layout().setSpacing(6)
        except Exception:
            pass
        try:
            self.io_card.body_lay.setSpacing(6)
        except Exception:
            pass

        io_sub = QLabel("Работа с файлами и источниками данных")
        io_sub.setStyleSheet(
            f'font-family: {_STATS_FONT_FAMILY}; font-size: 12px; font-weight: 600; line-height: 14px; color: {_C_TEXT}; '
            "background: transparent; border: none; padding: 0; margin: 0;"
        )

        tiles = QGridLayout()
        tiles.setContentsMargins(0, 0, 0, 0)
        tiles.setHorizontalSpacing(8)
        tiles.setVerticalSpacing(8)

        def _make_tile(text: str, icon_rel_path: str, handler) -> QPushButton:
            btn = QPushButton(text)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setFixedHeight(_DASHBOARD_IO_TILE_BTN_HEIGHT_PX)
            btn.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
            btn.setStyleSheet(_dashboard_io_tile_button_stylesheet())
            try:
                pm = QPixmap(resource_path(icon_rel_path))
                if not pm.isNull():
                    _isz = _DASHBOARD_IO_TILE_ICON_PX
                    pm = pm.scaled(_isz, _isz, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    btn.setIcon(QIcon(pm))
                    btn.setIconSize(QSize(_isz, _isz))
            except Exception:
                pass
            try:
                btn.clicked.connect(handler)
            except Exception:
                pass
            return btn

        t_import = _make_tile("Импорт данных", "assets/data_import.png", self._import_stats_package)
        t_excel = _make_tile("Экспорт в Excel", "assets/excel_export.png", self._export_stats_excel)
        t_pkg = _make_tile("Экспорт пакета", "assets/package_export.png", self._export_stats_package)
        t_sources = _make_tile("Источники данных", "assets/data_sources.png", self._open_sources_dialog)

        tiles.addWidget(t_import, 0, 0)
        tiles.addWidget(t_excel, 0, 1)
        tiles.addWidget(t_pkg, 1, 0)
        tiles.addWidget(t_sources, 1, 1)

        self.io_card.body_lay.addWidget(io_sub, 0, Qt.AlignLeft)
        self.io_card.body_lay.addLayout(tiles, 0)
        self.kpi_day_shift = _CompactKpiCard("Дневная смена", icon_path=SUN_PATH)
        self.kpi_night_shift = _CompactKpiCard("Ночная смена", icon_path=MOON_PATH)

        grid.addWidget(self.kpi_labels, 1, 0, 1, 2, Qt.AlignBottom)
        grid.addWidget(self.kpi_ops, 1, 2, 1, 2, Qt.AlignBottom)
        grid.addWidget(self.io_card, 1, 4, 1, 4)
        grid.addWidget(self.kpi_day_shift, 1, 8, 1, 2, Qt.AlignBottom)
        grid.addWidget(self.kpi_night_shift, 1, 10, 1, 2, Qt.AlignBottom)

        # Middle wide chart
        self.time_card = _Card("Этикеток по часам")
        self.time_chart = _VBarChart()
        self.time_chart.set_show_all_x_labels(True)
        self.time_chart.set_hour_zones_background(True)
        self.time_card.body_lay.addWidget(self.time_chart, 1)
        grid.addWidget(self.time_card, 2, 0, 1, 12)

        # Bottom row (3 cards, each spans 2 columns)
        self.top_products_card = _Card("Топ продуктов")
        self.top_products_chart = _HBarChart()
        self.top_products_chart.set_bar_width_factor(0.88)
        self.top_products_pie = _PieChart()
        tp_wrap = QFrame()
        tp_wrap.setObjectName("StatsChartInner")
        tp_wrap.setAttribute(Qt.WA_Hover, True)
        tp_wrap.setStyleSheet(_STATS_CHART_INNER_QSS)
        tp_wrap.setCursor(Qt.PointingHandCursor)
        self.top_products_stack = QStackedLayout(tp_wrap)
        self.top_products_stack.setContentsMargins(0, 0, 0, 0)
        self.top_products_stack.setSpacing(0)
        self.top_products_stack.addWidget(self.top_products_chart)
        self.top_products_stack.addWidget(self.top_products_pie)
        self.top_products_stack.setCurrentIndex(0)
        _tp_click = _StatsChartClickFilter(lambda: self._open_statistics_detail("products"), tp_wrap)
        for _w in (tp_wrap, self.top_products_chart, self.top_products_pie):
            _w.installEventFilter(_tp_click)
        self.top_products_card.body_lay.addWidget(tp_wrap, 1)
        grid.addWidget(self.top_products_card, 3, 0, 1, 4)

        self.top_staff_card = _Card("Топ сотрудников")
        self.top_staff_chart = _HBarChart()
        self.top_staff_chart.set_bar_width_factor(0.88)
        self.top_staff_pie = _PieChart()
        ts_wrap = QFrame()
        ts_wrap.setObjectName("StatsChartInner")
        ts_wrap.setAttribute(Qt.WA_Hover, True)
        ts_wrap.setStyleSheet(_STATS_CHART_INNER_QSS)
        ts_wrap.setCursor(Qt.PointingHandCursor)
        self.top_staff_stack = QStackedLayout(ts_wrap)
        self.top_staff_stack.setContentsMargins(0, 0, 0, 0)
        self.top_staff_stack.setSpacing(0)
        self.top_staff_stack.addWidget(self.top_staff_chart)
        self.top_staff_stack.addWidget(self.top_staff_pie)
        self.top_staff_stack.setCurrentIndex(0)
        _ts_click = _StatsChartClickFilter(lambda: self._open_statistics_detail("staff"), ts_wrap)
        for _w in (ts_wrap, self.top_staff_chart, self.top_staff_pie):
            _w.installEventFilter(_ts_click)
        self.top_staff_card.body_lay.addWidget(ts_wrap, 1)
        grid.addWidget(self.top_staff_card, 3, 4, 1, 4)

        self.workshops_card = _Card("Распределение по цехам")
        self.workshops_chart = _VBarChart()
        self.workshops_pie = _PieChart()
        ws_wrap = QFrame()
        ws_wrap.setObjectName("StatsChartInner")
        ws_wrap.setAttribute(Qt.WA_Hover, True)
        ws_wrap.setStyleSheet(_STATS_CHART_INNER_QSS)
        ws_wrap.setCursor(Qt.PointingHandCursor)
        self.workshops_stack = QStackedLayout(ws_wrap)
        self.workshops_stack.setContentsMargins(0, 0, 0, 0)
        self.workshops_stack.setSpacing(0)
        self.workshops_stack.addWidget(self.workshops_chart)
        self.workshops_stack.addWidget(self.workshops_pie)
        self.workshops_stack.setCurrentIndex(0)
        _ws_click = _StatsChartClickFilter(lambda: self._open_statistics_detail("workshops"), ws_wrap)
        for _w in (ws_wrap, self.workshops_chart, self.workshops_pie):
            _w.installEventFilter(_ws_click)
        self.workshops_card.body_lay.addWidget(ws_wrap, 1)
        grid.addWidget(self.workshops_card, 3, 8, 1, 4)

        def _make_view_toggle(on_bar, on_pie) -> QWidget:
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
                b.setFixedSize(52, 28)
                b.setStyleSheet(
                    "QToolButton { background: transparent; border: 1px solid rgba(148,163,184,0.35); border-radius: 8px; }"
                    "QToolButton:!checked:hover { background: rgba(248,250,252,0.98); border: 1px solid rgba(148,163,184,0.55); }"
                    "QToolButton:checked { background: rgba(224,231,255,0.8); border: 1px solid rgba(99,102,241,0.45); }"
                    "QToolButton:checked:hover { background: rgba(219,225,254,0.95); border: 1px solid rgba(99,102,241,0.58); }"
                )

            btn_bar.setIcon(icon_bar())
            btn_pie.setIcon(icon_pie())

            grp = QButtonGroup(w)
            grp.setExclusive(True)
            grp.addButton(btn_bar, 0)
            grp.addButton(btn_pie, 1)
            btn_bar.setChecked(True)

            btn_bar.clicked.connect(on_bar)
            btn_pie.clicked.connect(on_pie)

            lay.addWidget(btn_bar)
            lay.addWidget(btn_pie)
            return w

        self.top_products_card.set_header_right_widget(
            _make_view_toggle(lambda: self.top_products_stack.setCurrentIndex(0), lambda: self.top_products_stack.setCurrentIndex(1))
        )
        self.top_staff_card.set_header_right_widget(
            _make_view_toggle(lambda: self.top_staff_stack.setCurrentIndex(0), lambda: self.top_staff_stack.setCurrentIndex(1))
        )
        self.workshops_card.set_header_right_widget(
            _make_view_toggle(lambda: self.workshops_stack.setCurrentIndex(0), lambda: self.workshops_stack.setCurrentIndex(1))
        )

        # Empty-state как дополнительный блок НИЖЕ основных карточек (не вместо них).
        self.dashboard_empty_state_surface = self._build_empty_state_page()
        try:
            self.dashboard_empty_state_surface.setVisible(False)
        except Exception:
            pass
        grid.addWidget(self.dashboard_empty_state_surface, 4, 0, 1, 12)

        return w

    # ----- Data & refresh -----
    def refresh_dashboard(self):
        now = datetime.now()
        if self._period == "custom":
            if not self._custom_start_dt or not self._custom_end_dt:
                records = []
            else:
                records = filter_records_by_datetime_range(self._all_records, self._custom_start_dt, self._custom_end_dt)
        else:
            records = filter_records_by_period(self._all_records, self._period, now=now)

        if not records:
            self._dashboard_records = []
            if self._detail_mode:
                self.leave_statistics_detail()
            else:
                self.content_stack.setCurrentWidget(self.dashboard_widget)

            # Кнопка «Печать отчётов» на месте, но выглядит disabled (и даёт подсказку по клику).
            try:
                if getattr(self, "print_reports_btn", None) is not None:
                    self.print_reports_btn.setStyleSheet(_dashboard_io_tile_button_disabled_stylesheet())
            except Exception:
                pass

            # KPI всегда видны; при отсутствии данных показываем нули.
            try:
                self.kpi_ops.set_value(0)
                self.kpi_labels.set_value(0)
                self.kpi_day_shift.set_value(0)
                self.kpi_night_shift.set_value(0)
            except Exception:
                pass

            # Нижние аналитические блоки скрываем полностью (без пустых рамок и отступов).
            try:
                if getattr(self, "time_card", None) is not None:
                    self.time_card.setVisible(False)
                if getattr(self, "top_products_card", None) is not None:
                    self.top_products_card.setVisible(False)
                if getattr(self, "top_staff_card", None) is not None:
                    self.top_staff_card.setVisible(False)
                if getattr(self, "workshops_card", None) is not None:
                    self.workshops_card.setVisible(False)
            except Exception:
                pass

            # Empty-state ниже основных карточек.
            try:
                if getattr(self, "dashboard_empty_state_surface", None) is not None:
                    self.dashboard_empty_state_surface.setVisible(True)
            except Exception:
                pass

            # На всякий случай очищаем данные графиков (виджетов может не быть при скрытии).
            try:
                self.time_chart.set_data([], [], bar_color=QColor("#93c5fd"))
            except Exception:
                pass
            try:
                self.top_products_chart.set_data([], [], bar_color=QColor("#facc15"))
                self.top_products_pie.set_data([], [])
                self.top_staff_chart.set_data([], [], bar_color=QColor("#facc15"))
                self.top_staff_pie.set_data([], [])
                self.workshops_chart.set_data([], [], bar_color=QColor("#93c5fd"))
                self.workshops_pie.set_data([], [])
            except Exception:
                pass
            return
        self._dashboard_records = records
        if self._detail_mode:
            self.detail_view.populate(
                self._current_detail_type,
                self._dashboard_records,
                self._get_period_subtitle(),
            )
            self.content_stack.setCurrentWidget(self.detail_view)
        else:
            self.content_stack.setCurrentWidget(self.dashboard_widget)

        # Кнопка «Печать отчётов» в обычном стиле при наличии данных.
        try:
            if getattr(self, "print_reports_btn", None) is not None:
                self.print_reports_btn.setStyleSheet(_dashboard_io_tile_button_stylesheet())
        except Exception:
            pass

        # При наличии данных empty-state внутри dashboard не нужен.
        try:
            if getattr(self, "dashboard_empty_state_surface", None) is not None:
                self.dashboard_empty_state_surface.setVisible(False)
        except Exception:
            pass

        # При наличии данных нижние аналитические блоки должны быть видны.
        try:
            if getattr(self, "time_card", None) is not None:
                self.time_card.setVisible(True)
            if getattr(self, "top_products_card", None) is not None:
                self.top_products_card.setVisible(True)
            if getattr(self, "top_staff_card", None) is not None:
                self.top_staff_card.setVisible(True)
            if getattr(self, "workshops_card", None) is not None:
                self.workshops_card.setVisible(True)
        except Exception:
            pass

        ops = len(records)
        labels_total = sum(int(r.copies) for r in records)
        self.kpi_ops.set_value(ops)
        self.kpi_labels.set_value(labels_total)

        # Сменные счётчики: по всем записям внутри выбранного периода.
        day_shift_total, night_shift_total = compute_shift_totals(records)
        self.kpi_day_shift.set_value(day_shift_total)
        self.kpi_night_shift.set_value(night_shift_total)

        # time chart title + buckets
        if self._period == "custom":
            start_dt = self._custom_start_dt or records[0].dt
            end_dt = self._custom_end_dt or records[-1].dt
            span_hours = max(0.0, (end_dt - start_dt).total_seconds() / 3600.0)

            if span_hours <= 24.0:
                self.time_card.title_lbl.setText("Этикеток по часам")
                bucket_starts: list[datetime] = []
                cur = start_dt.replace(minute=0, second=0, microsecond=0)
                if cur > start_dt:
                    cur -= timedelta(hours=1)
                while cur <= end_dt and len(bucket_starts) < 200:
                    bucket_starts.append(cur)
                    cur += timedelta(hours=1)
                if not bucket_starts:
                    bucket_starts = [start_dt.replace(minute=0, second=0, microsecond=0)]

                buckets = [0] * len(bucket_starts)
                for r in records:
                    placed = False
                    for i, bs in enumerate(bucket_starts):
                        be = bs + timedelta(hours=1)
                        if bs <= r.dt < be:
                            buckets[i] += int(r.copies)
                            placed = True
                            break
                    if not placed:
                        # edge case: exactly on last bucket end
                        if bucket_starts and r.dt == bucket_starts[-1] + timedelta(hours=1):
                            buckets[-1] += int(r.copies)

                dates = {bs.date() for bs in bucket_starts}
                multi_day = len(dates) > 1
                labs = []
                for bs in bucket_starts:
                    if multi_day:
                        labs.append(bs.strftime("%d.%m %H"))
                    else:
                        labs.append(bs.strftime("%H"))
                self.time_chart.set_hour_zones_background((not multi_day) and len(bucket_starts) == 24)
            else:
                span_days = max(1, (end_dt.date() - start_dt.date()).days + 1)
                if span_days <= 31:
                    self.time_card.title_lbl.setText("Этикеток по дням")
                    days = []
                    d = start_dt.date()
                    end_d = end_dt.date()
                    while d <= end_d:
                        days.append(d)
                        d += timedelta(days=1)
                    idx = {dd: i for i, dd in enumerate(days)}
                    buckets = [0] * len(days)
                    for r in records:
                        rd = r.dt.date()
                        if rd in idx:
                            buckets[idx[rd]] += int(r.copies)
                    labs = [dd.strftime("%d.%m") for dd in days]
                else:
                    self.time_card.title_lbl.setText("Этикеток по неделям")
                    weeks: list[tuple[datetime, datetime]] = []
                    d0 = start_dt.date()
                    d1 = end_dt.date()
                    cur_d = d0
                    safety = 0
                    while cur_d <= d1 and safety < 800:
                        safety += 1
                        # ISO week: Monday..Sunday
                        monday = cur_d - timedelta(days=cur_d.weekday())
                        sunday = monday + timedelta(days=6)
                        week_end = min(sunday, d1)
                        weeks.append(
                            (
                                datetime.combine(monday, datetime.min.time()),
                                datetime.combine(week_end, datetime.max.time().replace(microsecond=0)),
                            )
                        )
                        cur_d = week_end + timedelta(days=1)

                    buckets = [0] * len(weeks)
                    for r in records:
                        for i, (ws, we) in enumerate(weeks):
                            if ws <= r.dt <= we:
                                buckets[i] += int(r.copies)
                                break
                    labs = []
                    for ws, we in weeks:
                        labs.append(f"{ws.strftime('%d.%m')}–{we.strftime('%d.%m')}")

                self.time_chart.set_hour_zones_background(False)
        elif self._period == "day":
            self.time_card.title_lbl.setText("Этикеток по часам")
            buckets = [0] * 24
            for r in records:
                buckets[int(r.dt.hour)] += int(r.copies)
            labs = [f"{h:02d}" for h in range(24)]
            self.time_chart.set_hour_zones_background(True)
        elif self._period == "week":
            self.time_card.title_lbl.setText("Этикеток за последние 7 дней")
            days = [ (now.date()) - timedelta(days=i) for i in range(6, -1, -1) ]
            idx = {d: i for i, d in enumerate(days)}
            buckets = [0] * 7
            for r in records:
                d = r.dt.date()
                if d in idx:
                    buckets[idx[d]] += int(r.copies)
            labs = [d.strftime("%d.%m") for d in days]
            self.time_chart.set_hour_zones_background(False)
        else:
            # month: последние 30 календарных дней (согласовано с filter_records_by_period)
            self.time_card.title_lbl.setText("Этикеток за последние 30 дней")
            days = [(now.date()) - timedelta(days=i) for i in range(29, -1, -1)]
            idx = {d: i for i, d in enumerate(days)}
            buckets = [0] * 30
            for r in records:
                d = r.dt.date()
                if d in idx:
                    buckets[idx[d]] += int(r.copies)
            labs = [d.strftime("%d.%m") for d in days]
            self.time_chart.set_hour_zones_background(False)

        use_warm_time_gradient = False
        if self._period == "week" or self._period == "month":
            use_warm_time_gradient = True
        elif self._period == "custom" and self._custom_start_dt and self._custom_end_dt:
            span_h = (self._custom_end_dt - self._custom_start_dt).total_seconds() / 3600.0
            use_warm_time_gradient = span_h > 24.0
        self.time_chart.set_warm_vertical_gradient(use_warm_time_gradient)
        self.time_chart.set_data(labs, buckets, bar_color=QColor("#93c5fd"))

        # top products/staff
        prod_counter = Counter()
        staff_counter = Counter()
        ws_counter = Counter()
        for r in records:
            p = normalize_stat_key(r.product)
            s = normalize_stat_key(r.made_by)
            w = normalize_stat_key(r.workshop)
            if p:
                prod_counter[p] += int(r.copies)
            if s:
                staff_counter[s] += int(r.copies)
            if w:
                ws_counter[w] += int(r.copies)

        top_p = prod_counter.most_common(6)
        top_s = staff_counter.most_common(5)
        top_ws = ws_counter.most_common(3)

        tp_labels = [k for k, _ in top_p]
        tp_values = [v for _, v in top_p]
        ts_labels = [k for k, _ in top_s]
        ts_values = [v for _, v in top_s]

        self.top_products_chart.set_data(tp_labels, tp_values, bar_color=QColor("#facc15"))
        self.top_products_pie.set_data(tp_labels, tp_values)
        self.top_staff_chart.set_data(ts_labels, ts_values, bar_color=QColor("#facc15"))
        self.top_staff_pie.set_data(ts_labels, ts_values)

        ws_labels = [k for k, _ in top_ws]
        ws_values = [v for _, v in top_ws]
        self.workshops_chart.set_data(ws_labels, ws_values, bar_color=QColor("#93c5fd"))
        self.workshops_pie.set_data(ws_labels, ws_values)

