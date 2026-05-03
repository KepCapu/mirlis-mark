from __future__ import annotations

import os
import subprocess
from datetime import datetime
from typing import Callable

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
    QHeaderView,
)

from stats_exchange import (
    list_imported_sources,
    read_package_manifest_dict,
    remove_imported_source,
    resolve_imported_zip_path,
)

# Ширина колонки «Импортирован» (дата/время); без растягивания на всю оставшуюся ширину таблицы.
_SOURCES_COL_IMPORT_WIDTH = 120
# Минимальная ширина колонок с датой «Период (с/по)» после resizeToContents
_SOURCES_COL_PERIOD_MIN = 108
_SOURCES_COL_RECORDS_MIN = 72



def _fmt_ts(ts: object) -> str:
    try:
        t = float(ts or 0.0)
    except Exception:
        return "—"
    if t <= 0:
        return "—"
    try:
        return datetime.fromtimestamp(t).strftime("%d.%m.%Y %H:%M")
    except Exception:
        return "—"


def _display_archive_basename(it: dict, zip_path: str) -> str:
    """Имя zip при импорте; иначе имя локального файла. UUID и логика импорта не затрагиваются."""
    s = str((it or {}).get("source_archive_basename") or "").strip()
    if s:
        return s
    if zip_path:
        try:
            return os.path.basename(zip_path)
        except Exception:
            pass
    return "—"


def _manifest_str(m: dict | None, *keys: str) -> str:
    if not m:
        return ""
    for k in keys:
        v = m.get(k)
        if v is None:
            continue
        s = str(v).strip()
        if s:
            return s
    return ""


def _display_computer_user(it: dict, zip_path: str) -> str:
    """
    Понятная подпись «компьютер / пользователь» для таблицы.
    Данные: реестр импорта → при необходимости manifest.json из локального zip.
    Для старых пакетов без user в манифесте — только «—» справа (или «— / —»).
    """
    it = it or {}
    pc = str(it.get("source_computer_name") or "").strip()
    usr = str(it.get("source_user_name") or "").strip()
    m: dict | None = None
    if ((not pc) or (not usr)) and zip_path:
        try:
            if os.path.isfile(zip_path):
                m = read_package_manifest_dict(zip_path)
        except Exception:
            m = None
    if m:
        if not pc:
            pc = _manifest_str(m, "source_computer_name", "source_station_label")
        if not usr:
            usr = _manifest_str(m, "source_user_name")
    if not pc:
        pc = str(it.get("source_station_label") or "").strip()
    if not pc and m:
        pc = _manifest_str(m, "source_station_label")
    if not pc:
        pc = "—"
    if not usr:
        usr = "—"
    return f"{pc} / {usr}"


class StatsSourcesDialog(QDialog):
    def __init__(self, parent=None, on_sources_changed: Callable[[], None] | None = None):
        super().__init__(parent)
        self.setWindowTitle("Источники данных")
        self.setModal(True)
        self.resize(1140, 460)
        self.setMinimumSize(1020, 400)
        self._on_sources_changed = on_sources_changed

        root = QVBoxLayout(self)
        root.setContentsMargins(14, 12, 14, 12)
        root.setSpacing(10)

        title = QLabel("Импортированные источники")
        title.setStyleSheet("font-size: 16px; font-weight: 650; color: #111827;")
        root.addWidget(title, 0, Qt.AlignLeft)

        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(
            ["Файл", "Компьютер / пользователь", "Период (с)", "Период (по)", "Записей", "Импортирован"]
        )
        self.table.setSelectionBehavior(self.table.SelectRows)
        self.table.setSelectionMode(self.table.SingleSelection)
        self.table.setEditTriggers(self.table.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        hdr = self.table.horizontalHeader()
        hdr.setStretchLastSection(False)
        # Файл, компьютер/пользователь — основная ширина; периоды/записи — по содержимому;
        # «Импортирован» — компактная фиксированная ширина.
        hdr.setSectionResizeMode(0, QHeaderView.Stretch)
        hdr.setSectionResizeMode(1, QHeaderView.Stretch)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(5, QHeaderView.Fixed)
        self.table.setSortingEnabled(False)
        root.addWidget(self.table, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        self.delete_btn = QPushButton("Удалить")
        self.close_btn = QPushButton("Закрыть")
        btn_row.addWidget(self.delete_btn)
        btn_row.addWidget(self.close_btn)
        root.addLayout(btn_row, 0)

        self.delete_btn.clicked.connect(self._on_delete_clicked)
        self.close_btn.clicked.connect(self.accept)
        self.table.cellDoubleClicked.connect(self._on_cell_double_clicked)

        self.refresh()

    def refresh(self) -> None:
        items = list_imported_sources() or []
        self.table.setRowCount(0)
        for it in items:
            r = self.table.rowCount()
            self.table.insertRow(r)

            pkg_id = str(it.get("package_id") or "").strip()
            zip_path = resolve_imported_zip_path(it)
            file_label = _display_archive_basename(it, zip_path)
            host_label = _display_computer_user(it, zip_path)
            pf = _fmt_ts(it.get("period_from"))
            pt = _fmt_ts(it.get("period_to"))
            try:
                rc = str(int(it.get("record_count") or 0))
            except Exception:
                rc = "0"
            ia = _fmt_ts(it.get("imported_at"))

            def _cell(text: str) -> QTableWidgetItem:
                c = QTableWidgetItem(text)
                c.setData(Qt.UserRole, pkg_id)
                c.setData(Qt.UserRole + 1, zip_path)
                return c

            self.table.setItem(r, 0, _cell(file_label))
            self.table.setItem(r, 1, _cell(host_label))
            self.table.setItem(r, 2, _cell(pf))
            self.table.setItem(r, 3, _cell(pt))
            self.table.setItem(r, 4, _cell(rc))
            self.table.setItem(r, 5, _cell(ia))

        self._apply_sources_table_column_widths()
        self.delete_btn.setEnabled(bool(items))

    def _apply_sources_table_column_widths(self) -> None:
        """Узкая колонка «Импортирован»; периоды/записи по тексту; 0–1 тянутся через Stretch."""
        for col in (2, 3, 4):
            self.table.resizeColumnToContents(col)
            w = self.table.columnWidth(col)
            if col in (2, 3):
                w = max(w, _SOURCES_COL_PERIOD_MIN)
            elif col == 4:
                w = max(w, _SOURCES_COL_RECORDS_MIN)
            self.table.setColumnWidth(col, w)
        self.table.setColumnWidth(5, _SOURCES_COL_IMPORT_WIDTH)

    def _on_cell_double_clicked(self, row: int, _column: int) -> None:
        if row < 0:
            return
        it0 = self.table.item(row, 0)
        if it0 is None:
            QMessageBox.warning(self, "Источники данных", "Файл архива не найден")
            return
        zip_path = str(it0.data(Qt.UserRole + 1) or "").strip()
        if not zip_path:
            QMessageBox.warning(self, "Источники данных", "Файл архива не найден")
            return
        try:
            zip_path = os.path.normpath(os.path.abspath(zip_path))
        except Exception:
            QMessageBox.warning(self, "Источники данных", "Файл архива не найден")
            return
        if not os.path.isfile(zip_path):
            QMessageBox.warning(self, "Источники данных", "Файл архива не найден")
            return
        if os.name != "nt":
            QMessageBox.information(
                self,
                "Источники данных",
                "Открытие проводника с выделением файла поддерживается только в Windows.",
            )
            return
        try:
            subprocess.Popen(f'explorer /select,"{zip_path}"', shell=True)
        except Exception:
            QMessageBox.warning(self, "Источники данных", "Файл архива не найден")

    def _selected_package_id(self) -> str:
        row = self.table.currentRow()
        if row < 0:
            return ""
        it = self.table.item(row, 0)
        if it is None:
            return ""
        return str(it.data(Qt.UserRole) or "").strip()

    def _on_delete_clicked(self) -> None:
        pkg_id = self._selected_package_id()
        if not pkg_id:
            QMessageBox.information(self, "Удаление", "Выберите источник для удаления.")
            return

        if (
            QMessageBox.question(
                self,
                "Удаление",
                "Удалить импортированный источник?\nЛокальные данные текущего ПК затронуты не будут.",
                QMessageBox.Yes | QMessageBox.No,
            )
            != QMessageBox.Yes
        ):
            return

        ok = remove_imported_source(pkg_id)
        if not ok:
            QMessageBox.warning(self, "Удаление", "Источник не найден или уже удалён.")
            self.refresh()
            return

        try:
            if self._on_sources_changed is not None:
                self._on_sources_changed()
        except Exception:
            pass

        self.refresh()

