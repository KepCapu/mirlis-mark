from __future__ import annotations

from collections import Counter
from typing import Callable

from statistics_data import PrintRecord, normalize_stat_key

from PyQt5.QtWidgets import (
    QAbstractItemView,
    QButtonGroup,
    QComboBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QScrollArea,
    QSizePolicy,
    QStackedLayout,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QToolButton,
    QVBoxLayout,
    QWidget,
    QStyledItemDelegate,
)
from PyQt5.QtCore import Qt, QSize, QEvent
from PyQt5.QtGui import QColor, QIcon, QPainter, QPixmap, QPen


_STATS_FONT_FAMILY = '"Inter","Segoe UI","Manrope","Arial",sans-serif'
_C_TITLE = "#1E2F45"
_C_TEXT = "#24364D"
_C_SUB = "#6B7C93"

_DETAIL_TITLES = {
    "products": "Топ продуктов",
    "staff": "Топ сотрудников",
    "workshops": "Распределение по цехам",
}
_NAME_HEADERS = {
    "products": "Продукт",
    "staff": "Сотрудник",
    "workshops": "Цех",
}


def _import_chart_widgets():
    from statistics_page import _HBarChart, _PieChart, _VBarChart

    return _HBarChart, _PieChart, _VBarChart


def _aggregate_rows(
    detail_type: str, records: list[PrintRecord]
) -> list[tuple[str, int, int, float]]:
    """Строки: имя, этикеток (сумма copies), операций (число записей), доля % по этикеткам."""
    if not records:
        return []
    if detail_type == "products":
        key = lambda r: r.product
    elif detail_type == "staff":
        key = lambda r: r.made_by
    elif detail_type == "workshops":
        key = lambda r: r.workshop
    else:
        return []

    labels_c = Counter()
    ops_c = Counter()
    for r in records:
        k0 = key(r)
        k = normalize_stat_key(k0)
        if not k:
            continue
        labels_c[k] += int(r.copies)
        ops_c[k] += 1

    total_l = sum(labels_c.values())
    rows: list[tuple[str, int, int, float]] = []
    for name, lc in labels_c.most_common():
        oc = int(ops_c[name])
        pct = (100.0 * float(lc) / float(total_l)) if total_l else 0.0
        rows.append((str(name), int(lc), oc, pct))
    return rows


class _SortTableItem(QTableWidgetItem):
    """Текст в ячейке задаётся конструктором; сортировка — по Qt.UserRole (число или строка)."""

    def __lt__(self, other: QTableWidgetItem) -> bool:
        v1 = self.data(Qt.UserRole)
        v2 = other.data(Qt.UserRole)
        if v1 is not None and v2 is not None:
            if isinstance(v1, (int, float)) and isinstance(v2, (int, float)):
                return float(v1) < float(v2)
            if isinstance(v1, str) and isinstance(v2, str):
                return v1 < v2
            try:
                return float(v1) < float(v2)
            except (TypeError, ValueError):
                pass
        return super().__lt__(other)


def _make_bar_pie_toggle(on_bar: Callable[[], None], on_pie: Callable[[], None]) -> QWidget:
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
            "QToolButton:checked { background: rgba(224,231,255,0.8); border: 1px solid rgba(99,102,241,0.45); }"
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


class StatisticsDetailView(QWidget):
    """
    Встроенный drill-down: график (bar/pie), поиск, таблица.
    detail_type: 'products' | 'staff' | 'workshops'
    """

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("StatisticsDetailView")
        self.setStyleSheet(f"background: transparent; font-family: {_STATS_FONT_FAMILY};")
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        self._detail_type = ""
        self._records: list[PrintRecord] = []
        self._rows: list[tuple[str, int, int, float]] = []
        self._all_rows: list[tuple[str, int, int, float]] = []
        self._leader_name: str = ""
        self._marker_by_name: dict[str, QColor] = {}
        self._hover_label: str = ""
        self._hover_row: int = -1
        self._chart_stack: QStackedLayout | None = None
        self._chart_holder: QWidget | None = None
        self._scroll: QScrollArea | None = None
        self._chart_block: QFrame | None = None
        self._tools_lay: QHBoxLayout | None = None
        self._sort_combo: QComboBox | None = None

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        card = QFrame()
        card.setObjectName("DetailCard")
        card.setStyleSheet(
            "#DetailCard { background: #ffffff; border: 1px solid #e5e7eb; border-radius: 18px; }"
        )
        card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self._card_lay = QVBoxLayout(card)
        self._card_lay.setContentsMargins(20, 16, 20, 16)
        self._card_lay.setSpacing(12)

        self._title_lbl = QLabel("")
        self._title_lbl.setStyleSheet(
            f"font-family: {_STATS_FONT_FAMILY}; font-size: 22px; font-weight: 650; color: {_C_TITLE}; "
            "background: transparent; border: none;"
        )
        self._title_lbl.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self._title_lbl.setContentsMargins(0, 0, 0, 0)
        self._sub_lbl = QLabel("")
        self._sub_lbl.setWordWrap(True)
        self._sub_lbl.setStyleSheet(
            f"font-family: {_STATS_FONT_FAMILY}; font-size: 15px; font-weight: 500; color: {_C_SUB}; "
            "background: transparent; border: none;"
        )
        self._sub_lbl.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self._sub_lbl.setContentsMargins(0, 0, 0, 0)

        # Mini KPI summary (right column; filled in populate)
        self._kpi_wrap = QFrame()
        self._kpi_wrap.setObjectName("DetailKPIWrap")
        self._kpi_wrap.setStyleSheet(
            "#DetailKPIWrap { background: transparent; border: none; }"
        )
        kpi_lay = QGridLayout(self._kpi_wrap)
        kpi_lay.setContentsMargins(0, 0, 0, 0)
        kpi_lay.setHorizontalSpacing(10)
        kpi_lay.setVerticalSpacing(10)

        def _make_kpi(title: str) -> tuple[QFrame, QLabel]:
            box = QFrame()
            box.setObjectName("DetailKPIBox")
            box.setStyleSheet(
                "#DetailKPIBox { background: #f8fafc; border: 1px solid #eef2f7; border-radius: 14px; }"
            )
            box.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            box_lay = QVBoxLayout(box)
            box_lay.setContentsMargins(12, 10, 12, 10)
            box_lay.setSpacing(2)

            t = QLabel(title)
            t.setStyleSheet(
                f"font-family: {_STATS_FONT_FAMILY}; font-size: 12px; font-weight: 650; color: {_C_SUB};"
                "background: transparent; border: none;"
            )
            v = QLabel("—")
            v.setStyleSheet(
                f"font-family: {_STATS_FONT_FAMILY}; font-size: 16px; font-weight: 700; color: {_C_TITLE};"
                "background: transparent; border: none;"
            )
            v.setMinimumHeight(20)
            box_lay.addWidget(t)
            box_lay.addWidget(v)
            return box, v

        self._kpi_items_box, self._kpi_items_val = _make_kpi("Элементов")
        self._kpi_labels_box, self._kpi_labels_val = _make_kpi("Всего этикеток")
        self._kpi_ops_box, self._kpi_ops_val = _make_kpi("Всего операций")
        self._kpi_leader_box, self._kpi_leader_val = _make_kpi("Лидер")
        kpi_lay.addWidget(self._kpi_items_box, 0, 0)
        kpi_lay.addWidget(self._kpi_labels_box, 0, 1)
        kpi_lay.addWidget(self._kpi_ops_box, 1, 0)
        kpi_lay.addWidget(self._kpi_leader_box, 1, 1)
        kpi_lay.setColumnStretch(0, 1)
        kpi_lay.setColumnStretch(1, 1)

        # =========================
        # Two main columns layout:
        # Left: title + period + toggle + chart
        # Right: KPI + legend + search + table
        # =========================

        main_row = QHBoxLayout()
        main_row.setSpacing(14)

        # Left column: visualization
        left_col = QFrame()
        left_col.setObjectName("DetailLeftCol")
        left_col.setStyleSheet(
            "#DetailLeftCol { background: #ffffff; border: 1px solid #eef2f7; border-radius: 16px; }"
        )
        left_col.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        left_lay = QVBoxLayout(left_col)
        # Tighter vertical rhythm: title -> period -> toggle -> chart
        left_lay.setContentsMargins(14, 3, 14, 10)
        left_lay.setSpacing(1)

        title_wrap = QWidget()
        title_wrap.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        title_col = QVBoxLayout()
        title_col.setContentsMargins(0, 0, 0, 0)
        title_col.setSpacing(2)
        title_col.addWidget(self._title_lbl)
        title_col.addWidget(self._sub_lbl)
        title_wrap.setLayout(title_col)
        # Header row: title/subtitle (left) + toggle (right)
        header_row = QWidget()
        header_row.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        header_lay = QHBoxLayout(header_row)
        header_lay.setContentsMargins(0, 0, 0, 0)
        header_lay.setSpacing(10)
        header_lay.addWidget(title_wrap, 1)

        toggle_host = QWidget()
        toggle_host.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self._tools_lay = QHBoxLayout(toggle_host)
        self._tools_lay.setContentsMargins(0, 0, 0, 0)
        self._tools_lay.setSpacing(8)
        header_lay.addWidget(toggle_host, 0, Qt.AlignTop | Qt.AlignRight)

        # Predictable header height: keeps chart-area starting consistently below.
        header_row.setFixedHeight(76)
        left_lay.addWidget(header_row, 0, Qt.AlignTop)

        self._chart_holder = QWidget()
        self._chart_holder.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self._chart_stack = QStackedLayout(self._chart_holder)
        self._chart_stack.setContentsMargins(0, 0, 0, 0)

        # Dedicated chart block for detail-view (fills remaining space under header)
        self._chart_block = QFrame()
        self._chart_block.setObjectName("DetailChartBlock")
        self._chart_block.setStyleSheet(
            "#DetailChartBlock { background: transparent; border: none; }"
        )
        self._chart_block.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        chart_block_lay = QVBoxLayout(self._chart_block)
        chart_block_lay.setContentsMargins(0, 0, 0, 0)
        chart_block_lay.setSpacing(0)

        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(False)
        self._scroll.setFrameShape(QFrame.NoFrame)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self._scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        # If the chart-holder is smaller than the viewport, keep it top-anchored (never centered).
        self._scroll.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self._scroll.setMinimumHeight(320)
        self._scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self._scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")
        self._scroll.setWidget(self._chart_holder)
        chart_block_lay.addWidget(self._scroll, 1)
        left_lay.addWidget(self._chart_block, 1)

        # Right column: analytics (KPI + legend + search + table)
        right_col = QFrame()
        right_col.setObjectName("DetailRightCol")
        right_col.setStyleSheet(
            "#DetailRightCol { background: #ffffff; border: 1px solid #eef2f7; border-radius: 16px; }"
        )
        right_col.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        right_lay = QVBoxLayout(right_col)
        right_lay.setContentsMargins(14, 12, 14, 12)
        right_lay.setSpacing(10)

        right_lay.addWidget(self._kpi_wrap, 0)

        search_row = QHBoxLayout()
        search_row.setSpacing(10)

        self._search = QLineEdit()
        self._search.setPlaceholderText("Поиск по названию…")
        self._search.setClearButtonEnabled(True)
        self._search.setMinimumHeight(42)
        self._search.setStyleSheet(
            f"QLineEdit {{ background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 12px; "
            f"padding: 10px 14px; font-family: {_STATS_FONT_FAMILY}; font-size: 15px; color: {_C_TEXT}; }}"
            f"QLineEdit:focus {{ border: 1px solid #6366f1; }}"
        )
        # Make room for a wider sort dropdown on the right.
        search_row.addWidget(self._search, 3)

        self._sort_combo = QComboBox()
        self._sort_combo.setMinimumHeight(42)
        self._sort_combo.setMinimumWidth(220)
        self._sort_combo.setCursor(Qt.PointingHandCursor)
        self._sort_combo.setStyleSheet(
            f"QComboBox {{ background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 12px; "
            f"padding: 8px 12px; font-family: {_STATS_FONT_FAMILY}; font-size: 14px; color: {_C_TEXT}; }}"
            "QComboBox::drop-down { border: none; width: 26px; }"
            "QComboBox QAbstractItemView { background: #ffffff; border: 1px solid #e5e7eb; selection-background-color: rgba(99,102,241,0.12); }"
        )
        self._sort_combo.addItem("Этикетки: по убыванию", ("col", 1, Qt.DescendingOrder))
        self._sort_combo.addItem("Этикетки: по возрастанию", ("col", 1, Qt.AscendingOrder))
        self._sort_combo.addItem("Операции: по убыванию", ("col", 2, Qt.DescendingOrder))
        self._sort_combo.addItem("Операции: по возрастанию", ("col", 2, Qt.AscendingOrder))
        self._sort_combo.addItem("Доля: по убыванию", ("col", 3, Qt.DescendingOrder))
        self._sort_combo.addItem("Доля: по возрастанию", ("col", 3, Qt.AscendingOrder))
        self._sort_combo.addItem("Название: А–Я", ("col", 0, Qt.AscendingOrder))
        self._sort_combo.addItem("Название: Я–А", ("col", 0, Qt.DescendingOrder))
        search_row.addWidget(self._sort_combo, 2)

        right_lay.addLayout(search_row, 0)

        self._table = QTableWidget()
        self._table.setColumnCount(4)
        self._table.setAlternatingRowColors(True)
        self._table.setSelectionBehavior(QAbstractItemView.SelectRows)
        # Detail table should never keep selection/current/focus visuals.
        self._table.setSelectionMode(QAbstractItemView.NoSelection)
        self._table.setFocusPolicy(Qt.NoFocus)
        self._table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._table.verticalHeader().setVisible(False)
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self._table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self._table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self._table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self._table.setSortingEnabled(True)
        # Disable header-driven sorting UI (keep sorting only via dropdown).
        hh = self._table.horizontalHeader()
        hh.setSectionsClickable(False)
        hh.setSortIndicatorShown(False)
        hh.setSortIndicator(-1, Qt.AscendingOrder)
        self._table.setIconSize(QSize(12, 12))
        self._table.setStyleSheet(
            f"QTableWidget {{ background: #ffffff; border: 1px solid #eef2f7; border-radius: 12px; "
            f"font-family: {_STATS_FONT_FAMILY}; font-size: 14px; color: {_C_TEXT}; gridline-color: #f1f5f9; }}"
            f"\nQTableWidget::item {{ padding: 8px 10px; }}"
            f"\nQHeaderView::section {{ background: #f8fafc; color: {_C_SUB}; border: none; "
            "\nborder-bottom: 1px solid #eef2f7; padding: 10px 8px; font-weight: 600; font-size: 13px; }}"
        )
        right_lay.addWidget(self._table, 5)

        # Hover sync from table -> chart
        self._table.setMouseTracking(True)
        self._table.viewport().setMouseTracking(True)
        self._table.viewport().installEventFilter(self)
        self._table.setItemDelegate(_DetailHoverRowDelegate(self))

        main_row.addWidget(left_col, 1)   # 50%
        main_row.addWidget(right_col, 1)  # 50%
        self._card_lay.addLayout(main_row, 1)
        root.addWidget(card, 1)

        self._search.textChanged.connect(self._on_search)
        if self._sort_combo is not None:
            self._sort_combo.currentIndexChanged.connect(self._apply_sort_ui)

    def _clear_tools_row(self) -> None:
        assert self._tools_lay is not None
        while self._tools_lay.count():
            item = self._tools_lay.takeAt(0)
            w = item.widget()
            if w is not None:
                w.setParent(None)
                w.deleteLater()

    def _clear_chart_stack(self) -> None:
        assert self._chart_stack is not None
        while self._chart_stack.count():
            w = self._chart_stack.widget(0)
            self._chart_stack.removeWidget(w)
            w.deleteLater()

    def populate(
        self,
        detail_type: str,
        records: list[PrintRecord],
        period_label: str,
    ) -> None:
        self._detail_type = (detail_type or "").strip().lower()
        self._records = list(records or [])
        self._rows = _aggregate_rows(self._detail_type, self._records)
        self._all_rows = list(self._rows)
        self._leader_name = ""
        if self._rows:
            try:
                self._leader_name = max(self._rows, key=lambda r: int(r[1]))[0]
            except Exception:
                self._leader_name = self._rows[0][0]

        self._title_lbl.setText(_DETAIL_TITLES.get(self._detail_type, "Статистика"))
        self._sub_lbl.setText(period_label or "")
        nh = _NAME_HEADERS.get(self._detail_type, "Название")
        self._table.setHorizontalHeaderLabels([nh, "Этикеток", "Операций", "Доля"])

        self._clear_tools_row()
        self._clear_chart_stack()

        _HBarChart, _PieChart, _VBarChart = _import_chart_widgets()

        labs = [r[0] for r in self._rows]
        vals = [r[1] for r in self._rows]
        n = max(1, len(labs))

        # Build color markers for the table:
        # keep the same per-index palette as the charts (bar/pie).
        try:
            from statistics_page import _BAR_PALETTE_HEX
        except Exception:
            _BAR_PALETTE_HEX = []
        self._marker_by_name = {}
        if _BAR_PALETTE_HEX:
            for i, lab in enumerate(labs):
                try:
                    self._marker_by_name[str(lab)] = QColor(_BAR_PALETTE_HEX[i % len(_BAR_PALETTE_HEX)])
                except Exception:
                    pass

        # KPI summary (always based on full aggregated set)
        total_labels = sum(int(r[1]) for r in self._rows) if self._rows else 0
        total_ops = sum(int(r[2]) for r in self._rows) if self._rows else 0
        self._kpi_items_val.setText(str(len(self._rows)))
        self._kpi_labels_val.setText(str(int(total_labels)))
        self._kpi_ops_val.setText(str(int(total_ops)))
        self._kpi_leader_val.setText(self._leader_name or "—")

        if self._detail_type in ("products", "staff"):
            bar_chart = _HBarChart()
            bar_chart.set_bar_width_factor(1.0)
            if hasattr(bar_chart, "set_detail_mode"):
                bar_chart.set_detail_mode(True)
            if hasattr(bar_chart, "hoverLabelChanged"):
                bar_chart.hoverLabelChanged.connect(self._on_chart_hover_label)  # type: ignore[attr-defined]
            bar_chart.set_data(labs, vals, bar_color=QColor("#facc15"))
        else:
            bar_chart = _VBarChart()
            if hasattr(bar_chart, "set_detail_mode"):
                bar_chart.set_detail_mode(True)
            if hasattr(bar_chart, "hoverLabelChanged"):
                bar_chart.hoverLabelChanged.connect(self._on_chart_hover_label)  # type: ignore[attr-defined]
            bar_chart.set_data(labs, vals, bar_color=QColor("#93c5fd"))

        pie_chart = _PieChart()
        if hasattr(pie_chart, "set_detail_mode"):
            pie_chart.set_detail_mode(True)
        if hasattr(pie_chart, "hoverLabelChanged"):
            pie_chart.hoverLabelChanged.connect(self._on_chart_hover_label)  # type: ignore[attr-defined]
        pie_chart.set_data(labs, vals)
        pie_chart.set_legend_max_items(max(12, min(n, 80)))

        assert self._chart_stack is not None
        self._chart_stack.addWidget(bar_chart)
        self._chart_stack.addWidget(pie_chart)
        self._chart_stack.setCurrentIndex(0)

        # Container behavior:
        # For horizontal detail bars we need the chart-holder to stretch to the full viewport width,
        # otherwise _HBarChart gets a narrow rect and bars look short.
        if self._scroll is not None:
            if self._detail_type in ("products", "staff"):
                self._scroll.setWidgetResizable(True)
                self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            else:
                # Workshops: keep horizontal scroll when needed, but allow the holder/widget
                # to expand vertically with the viewport so the donut chart isn't constrained
                # to a small minimum height.
                self._scroll.setWidgetResizable(True)
                self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        # Chart widgets should size to content; chart-area grows downwards.
        # We keep only a modest minimum so the block doesn't collapse.
        min_h = 240
        bar_chart.setMinimumHeight(min_h)
        pie_chart.setMinimumHeight(min_h)
        pie_chart.setMinimumWidth(max(620, 520 if self._detail_type == "workshops" else 620))
        if self._detail_type == "workshops":
            bar_chart.setMinimumWidth(max(520, min(1600, 32 * n + 120)))

        # No fixed heights here: chart area expands (free space remains at the bottom).

        # QScrollArea с widgetResizable(False) не раздувает widget сам.
        # Если не задать геометрию holder'у, он может остаться 0x0 и графики не будут видны.
        if self._chart_holder is not None and (self._scroll is None or not self._scroll.widgetResizable()):
            min_w = max(
                520 if self._detail_type == "workshops" else 360,
                int(getattr(bar_chart, "minimumWidth", lambda: 0)() or 0),
                int(getattr(pie_chart, "minimumWidth", lambda: 0)() or 0),
            )
            self._chart_holder.setMinimumSize(min_w, min_h)
            self._chart_holder.resize(min_w, min_h)
            self._chart_holder.update()

        assert self._tools_lay is not None
        cs = self._chart_stack
        def _on_bar() -> None:
            cs.setCurrentIndex(0)

        def _on_pie() -> None:
            cs.setCurrentIndex(1)

        toggle = _make_bar_pie_toggle(_on_bar, _on_pie)
        # Right-align toggle inside toggle_host
        self._tools_lay.addStretch(1)
        self._tools_lay.addWidget(toggle, 0, Qt.AlignRight)

        self._fill_table(self._all_rows)
        if self._sort_combo is not None:
            self._sort_combo.setCurrentIndex(0)
        self._apply_sort_ui()
        self._reset_table_interaction_state()

        self._search.blockSignals(True)
        self._search.clear()
        self._search.blockSignals(False)
        # Default view: bar

    def _fill_table(self, rows: list[tuple[str, int, int, float]]) -> None:
        self._table.setRowCount(0)
        self._table.setSortingEnabled(False)
        for name, lc, oc, pct in rows:
            r = self._table.rowCount()
            self._table.insertRow(r)

            it0 = QTableWidgetItem(name)
            it0.setData(Qt.UserRole, name.lower())
            # Colored marker matching the chart's palette color for this item.
            c = self._marker_by_name.get(str(name))
            if isinstance(c, QColor) and c.isValid():
                s = 10
                pm = QPixmap(s, s)
                pm.fill(Qt.transparent)
                p = QPainter(pm)
                p.setRenderHint(QPainter.Antialiasing, True)
                p.setPen(Qt.NoPen)
                p.setBrush(c)
                p.drawEllipse(0, 0, s - 1, s - 1)
                p.end()
                it0.setIcon(QIcon(pm))

            it1 = _SortTableItem(str(lc))
            it1.setData(Qt.UserRole, float(lc))
            it2 = _SortTableItem(str(oc))
            it2.setData(Qt.UserRole, float(oc))
            pct_txt = f"{pct:.1f} %".replace(".", ",")
            it3 = _SortTableItem(pct_txt)
            it3.setData(Qt.UserRole, float(pct))

            self._table.setItem(r, 0, it0)
            self._table.setItem(r, 1, it1)
            self._table.setItem(r, 2, it2)
            self._table.setItem(r, 3, it3)

        self._table.setSortingEnabled(True)
        hh = self._table.horizontalHeader()
        hh.setSortIndicatorShown(False)
        hh.setSortIndicator(-1, Qt.AscendingOrder)
        self._apply_table_hover_visuals()
        self._reset_table_interaction_state()

    def _reset_table_interaction_state(self) -> None:
        """Ensure the table has no persistent selection/current/focus state."""
        try:
            self._table.clearSelection()
        except Exception:
            pass
        try:
            # Some Qt builds don't like (-1, -1); None is safest.
            self._table.setCurrentItem(None)
        except Exception:
            pass
        try:
            self._table.setCurrentCell(-1, -1)
        except Exception:
            pass
        try:
            self._table.clearFocus()
        except Exception:
            pass

    def _apply_table_hover_visuals(self) -> None:
        """Highlight hovered item and slightly mute others."""
        t = (self._hover_label or "").strip()
        self._hover_row = -1
        for r in range(self._table.rowCount()):
            it0 = self._table.item(r, 0)
            name = it0.text() if it0 is not None else ""
            is_active = bool(t and name == t)
            if is_active:
                self._hover_row = r
            # Update marker icon to look active/inactive without changing data.
            c = self._marker_by_name.get(str(name))
            if it0 is not None and isinstance(c, QColor) and c.isValid():
                s = 10
                pm = QPixmap(s, s)
                pm.fill(Qt.transparent)
                p = QPainter(pm)
                p.setRenderHint(QPainter.Antialiasing, True)
                if is_active:
                    pen = QPen(QColor(255, 255, 255, 235))
                    pen.setWidthF(1.6)
                    p.setPen(pen)
                    p.setBrush(QColor(c).lighter(125))
                else:
                    p.setPen(Qt.NoPen)
                    p.setBrush(c)
                p.drawEllipse(0, 0, s - 1, s - 1)
                p.end()
                it0.setIcon(QIcon(pm))
            for c in range(self._table.columnCount()):
                it = self._table.item(r, c)
                if it is None:
                    continue
                if not t:
                    it.setForeground(QColor(_C_TEXT))
                else:
                    it.setForeground(QColor(_C_TITLE) if is_active else QColor(_C_SUB))
        self._table.viewport().update()

    def _on_chart_hover_label(self, label: str) -> None:
        self._hover_label = (label or "").strip()
        self._apply_table_hover_visuals()

    def _set_chart_hover_label(self, label: str) -> None:
        lab = (label or "").strip()
        self._hover_label = lab
        self._apply_table_hover_visuals()
        if self._chart_stack is None:
            return
        w = self._chart_stack.currentWidget()
        if w is None:
            return
        if hasattr(w, "set_hover_label"):
            try:
                w.set_hover_label(lab)  # type: ignore[attr-defined]
            except Exception:
                pass

    def eventFilter(self, watched: QWidget, event) -> bool:  # type: ignore[override]
        if watched is self._table.viewport():
            et = event.type()
            if et == QEvent.MouseMove:
                pos = event.pos()
                idx = self._table.indexAt(pos)
                if idx.isValid():
                    it = self._table.item(idx.row(), 0)
                    self._set_chart_hover_label(it.text() if it is not None else "")
                else:
                    self._set_chart_hover_label("")
            elif et == QEvent.Leave:
                self._set_chart_hover_label("")
        return super().eventFilter(watched, event)

    def _apply_sort_ui(self) -> None:
        if self._sort_combo is None:
            return
        data = self._sort_combo.currentData()
        if isinstance(data, tuple) and len(data) == 3 and data[0] == "col":
            _, col, order = data
            try:
                self._table.sortByColumn(int(col), order)
            except Exception:
                pass
            # Never show header sort indicator (sorting is controlled only by dropdown).
            hh = self._table.horizontalHeader()
            hh.setSortIndicatorShown(False)
            hh.setSortIndicator(-1, Qt.AscendingOrder)

    def _on_search(self, text: str) -> None:
        t = (text or "").strip().lower()
        if not t:
            self._fill_table(self._all_rows)
            self._apply_sort_ui()
            return
        filtered = [row for row in self._all_rows if t in str(row[0]).lower()]
        self._fill_table(filtered)
        self._apply_sort_ui()


class _DetailHoverRowDelegate(QStyledItemDelegate):
    """Subtle border around the active hover row (detail table)."""

    def __init__(self, view: StatisticsDetailView):
        super().__init__(view)
        self._view = view

    def paint(self, painter: QPainter, option, index):  # type: ignore[override]
        # Force clean zebra and ignore any selection/current/focus state.
        # Row 0 must be gray, row 1 white, etc.
        try:
            opt = option
            opt.state &= ~QStyle.State_Selected
            opt.state &= ~QStyle.State_HasFocus
            opt.state &= ~QStyle.State_Active
        except Exception:
            opt = option

        zebra_gray = QColor("#f3f4f6")
        zebra_white = QColor("#ffffff")
        base_bg = zebra_gray if (index.row() % 2 == 0) else zebra_white

        painter.save()
        painter.setPen(Qt.NoPen)
        painter.setBrush(base_bg)
        painter.drawRect(opt.rect)
        painter.restore()

        super().paint(painter, opt, index)

        r_active = getattr(self._view, "_hover_row", -1)
        if r_active < 0 or index.row() != r_active:
            return
        # Draw a thin, soft border around the row (without a heavy system focus ring).
        rect = opt.rect.adjusted(0, 0, -1, -1)
        pen = QPen(QColor(99, 102, 241, 170))
        pen.setWidthF(1.7)
        pen.setJoinStyle(Qt.RoundJoin)
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)
        col = index.column()
        last_col = index.model().columnCount() - 1
        # top/bottom on every cell; left on first; right on last
        painter.drawLine(rect.topLeft(), rect.topRight())
        painter.drawLine(rect.bottomLeft(), rect.bottomRight())
        if col == 0:
            painter.drawLine(rect.topLeft(), rect.bottomLeft())
        if col == last_col:
            painter.drawLine(rect.topRight(), rect.bottomRight())
        painter.restore()
