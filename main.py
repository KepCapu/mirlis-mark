# main.py
# Mirlis Mark — Система маркировки
# UI: “как на картинке” + редактор предпросмотра + ВИДИМЫЕ стрелочки в выпадающих списках
#
# ВАЖНО:
# - excel_loader.py / label_logic.py / printer.py НЕ ТРОГАЕМ
# - логотип берём по пути: D:\mirlis_mark\Mirlis software logo.png

import sys
import os
import time
import math
import json
from datetime import datetime

from PyQt6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QLabel,
    QComboBox,
    QLineEdit,
    QPushButton,
    QHBoxLayout,
    QCheckBox,
    QMessageBox,
    QCompleter,
    QFrame,
    QScrollArea,
    QSizePolicy,
    QTextEdit,
    QToolButton,
    QFontComboBox,
    QSpacerItem,
    QFileDialog,
)
from PyQt6.QtCore import QTimer, Qt, QUrl, QSize
from PyQt6.QtGui import (
    QDesktopServices,
    QIcon,
    QPixmap,
    QFont,
    QTextCharFormat,
    QTextBlockFormat,
    QTextCursor,
    QSurfaceFormat,
    QImage,
    QPainter,
    QColor,
)
from PyQt6.QtCore import QSizeF
from PyQt6.QtCore import QStringListModel
from PyQt6.QtMultimedia import QMediaPlayer, QAudioOutput
from PyQt6.QtMultimediaWidgets import QVideoWidget

from excel_loader import load_products, load_staff
from label_logic import build_label, format_dt
from printer import print_text_as_bitmap_tspl, print_raw
import win32print


# -------------------- CONFIG --------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_PATH = os.path.join(BASE_DIR, "data_sources", "products.xlsx")
CONFIG_PATH = os.path.join(BASE_DIR, "settings.json")
SHEET_PRODUCTS = "продукт"
SHEET_MADE = "изготовил"
SHEET_CHECKED = "проверил"

LOGO_PATH = os.path.join(BASE_DIR, "assets", "logo.png")
SPLASH_VIDEO_PATH = os.path.join(BASE_DIR, "assets", "loadingscreen.mp4")

APP_TITLE = "Mirlis Mark — Система маркировки"
APP_MARK = "Mark"
APP_VERSION = "1.0"
APP_SUBTITLE = "Система маркировки"


# -------------------- HELPERS --------------------
def _load_settings() -> dict:
    """Загрузить настройки из settings.json."""
    try:
        if os.path.isfile(CONFIG_PATH):
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def _save_settings(data: dict):
    """Сохранить настройки в settings.json."""
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _fmt_dt_local(ts: float) -> str:
    return datetime.fromtimestamp(ts).strftime("%d.%m.%Y %H:%M:%S")


def _safe_int(v, default=0):
    try:
        return int(v)
    except Exception:
        return default


# -------------------- UI Building Blocks --------------------
class Card(QFrame):
    """Скруглённая карточка с лёгкой рамкой."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("Card")
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)


class Pill(QLabel):
    """Плашка-‘pill’ (серый бэйдж)."""

    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self.setObjectName("Pill")
        self.setSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed)


class HeaderLabel(QLabel):
    """Заголовок секции по центру с короткой тенью-плашкой."""

    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self.setObjectName("SectionTitle")
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)


class ComboBoxFixedArrow(QComboBox):
    """
    Комбобокс с гарантированно видимой стрелкой:
    - оставляем drop-down зону
    - НЕ убираем down-arrow через QSS
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("ComboWithArrow")


class ToolBtn(QToolButton):
    def __init__(self, text="", parent=None):
        super().__init__(parent)
        self.setText(text)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setObjectName("ToolBtn")
        self.setCheckable(True)
        self.setAutoRaise(False)


class ActionBtn(QPushButton):
    def __init__(self, text="", kind="default", parent=None):
        super().__init__(text, parent)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setObjectName(f"Btn_{kind}")


# -------------------- SPLASH VIDEO --------------------
class SplashVideo(QWidget):
    """Окно загрузки с воспроизведением видео перед запуском основного приложения."""

    def __init__(self, video_path: str, on_finished_callback=None, parent=None):
        super().__init__(parent)
        self._on_finished = on_finished_callback
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        self.setStyleSheet("background-color: #000000;")
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, False)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self._video_widget = QVideoWidget()
        self._video_widget.setStyleSheet("background-color: #000000;")
        self._video_widget.setAspectRatioMode(Qt.AspectRatioMode.KeepAspectRatioByExpanding)
        layout.addWidget(self._video_widget)

        self._player = QMediaPlayer()
        self._audio_output = QAudioOutput()
        self._player.setAudioOutput(self._audio_output)
        self._player.setVideoOutput(self._video_widget)
        self._player.setSource(QUrl.fromLocalFile(video_path))
        self._player.mediaStatusChanged.connect(self._on_media_status_changed)

    def _on_media_status_changed(self, status):
        if status == QMediaPlayer.MediaStatus.EndOfMedia:
            if callable(self._on_finished):
                self._on_finished()
            self.close()

    def play(self):
        self._player.play()

    def showEvent(self, event):
        super().showEvent(event)
        self._player.play()


# -------------------- MAIN APP --------------------
class MirlisMarkApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)

        # окно растягиваемое, но с адекватным минимумом
        self.setMinimumSize(1100, 650)

        # печать (последний напечатанный текст для «Повторить»)
        self.last_printed_preview_text: str | None = None
        self._last_printed_tspl_bytes: bytes | None = None
        self.last_history_entry: dict | None = None

        # данные
        self.products = []
        self.staff_made = []
        self.staff_checked = []
        self.loaded_at_str = "—"

        # путь к Excel (из настроек или дефолтный)
        settings = _load_settings()
        self.excel_path = settings.get("excel_path", EXCEL_PATH)

        # флаги чтобы редактор НЕ “откатывал” форматирование
        self._updating_preview = False
        self._user_edited_preview = False

        # базовый размер шрифта в редакторе
        self._base_font_size = 20

        # состояние истории (должно существовать до init_ui / _rebuild_history_view)
        self.history_entries: list[dict] = []
        self._history_filter_text: str = ""
        self.history_page: int = 0
        self.history_page_size: int = 6
        self._loading_from_history = False
        self._selected_history_id = None

        self._apply_global_style()
        self.init_ui()
        self.reload_excel(show_message=False)

        # обновляем превью по таймеру (для “тикания” времени),
        # НО НЕ переписываем редактор, если пользователь редактировал.
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.refresh_preview)
        self.timer.start(5000)

        self.refresh_preview()

    # ---------------- STYLE ----------------
    def _apply_global_style(self):
        # Важно: Qt НЕ поддерживает align-self, image-repeat и т.п. — не используем.
        self.setStyleSheet(
            """
            QWidget {
                background: #f6f7f9;
                font-family: "Segoe UI";
                color: #111827;
            }

            /* верхний белый хедер-кард */
            #TopBar {
                background: #ffffff;
                border: 1px solid #e5e7eb;
                border-radius: 18px;
            }

            /* карточки */
            #Card {
                background: #ffffff;
                border: 1px solid #e5e7eb;
                border-radius: 18px;
            }

            /* история */
            #HistoryPanel {
                background: transparent;
            }

            #HistoryScroll {
                border: none;
                background: transparent;
            }

            #HistoryCard {
                background: #ffffff;
                border-radius: 14px;
                border: 1px solid #e5e7eb;
                padding: 10px 12px;
            }

            #HistoryCard:hover {
                background: #f9fafb;
            }

            #HistoryCard[selected="true"] {
                border: 2px solid #4f46e5;
                background: #eef2ff;
            }

            #LabelWrap {
                background: transparent;
            }

            /* заголовок секций — “короткая” плашка */
            #SectionTitle {
                background: #eef2f6;
                border-radius: 14px;
                padding: 10px 22px;
                font-size: 22px;
                font-weight: 800;
                letter-spacing: 0.2px;
            }

            /* серые подплашки-поля (только для нужных заголовков) */
            #FieldLabel {
                background: #eef2f6;
                border-radius: 14px;
                padding: 10px 16px;
                font-size: 16px;
                font-weight: 500;
                color: #111827;
            }

            #Pill {
                background: #eef2f6;
                border-radius: 14px;
                padding: 8px 14px;
                font-size: 14px;
                font-weight: 600;
                color: #111827;
            }

            /* отдельная широкая плашка статуса Excel в TopBar */
            #ExcelPill {
                background: #eef2f6;
                border-radius: 18px;
                padding: 14px 22px;
                font-size: 14px;
                font-weight: 600;
                color: #0f172a;
            }

            /* инпуты */
            QLineEdit, QTextEdit {
                background: #ffffff;
                border: 1px solid #cfd6e0;
                border-radius: 16px;
                padding: 10px 14px;
                font-size: 14px;
                selection-background-color: #cfe3ff;
            }
            QLineEdit:focus, QTextEdit:focus {
                border: 1px solid #6ea8fe;
            }

            /* чекбокс */
            QCheckBox {
                font-size: 14px;
                background: transparent;
            }

            QCheckBox:focus {
                outline: none;
            }

            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }

            /* кнопки (общая база) */
            QPushButton {
                border-radius: 16px;
                padding: 10px 18px;
                font-size: 14px;
                border: 1px solid #d0d7e2;
                background: #ffffff;
                color: #111827;
                box-shadow: none;
            }
            QPushButton:hover {
                background: #eef2ff;
            }
            QPushButton:pressed {
                background: #e5e7ff;
                border-color: #4f46e5;
            }
            QPushButton:disabled {
                background: #f3f4f6;
                color: #9ca3af;
                border-color: #e5e7eb;
            }

            /* плоские кнопки степперов (+ / -) */
            #StepperBtn {
                background: #f9fafb;
                border-radius: 12px;
                border: 1px solid #d1d5db;
                box-shadow: none;
            }
            #StepperBtn:hover {
                background: #e5e7eb;
                border-color: #d1d5db;
            }
            #StepperBtn:pressed {
                background: #d1d5db;
            }

            /* основная зелёная кнопка (ПЕЧАТЬ) */
            #Btn_primary {
                background: #16a34a;
                border: 1px solid #15803d;
                color: #ffffff;
                font-weight: 800;
                font-size: 18px;
                letter-spacing: 0.8px;
                padding: 18px 18px;
                border-radius: 18px;
            }
            #Btn_primary:hover {
                background: #15803d;
                border-color: #166534;
            }
            #Btn_primary:pressed {
                background: #166534;
                border-color: #14532d;
            }
            #Btn_primary:disabled {
                background: #d1fae5;
                border: 1px solid #bbf7d0;
                color: rgba(255,255,255,0.8);
            }

            /* вторичные кнопки (Повторить, Количество и т.п.) */
            #Btn_secondary {
                background: #f9fafb;
                border: 1px solid #d1d5db;
                color: #111827;
                font-weight: 700;
                font-size: 16px;
                padding: 18px 18px;
                border-radius: 18px;
            }
            #Btn_secondary:hover {
                background: #eef2ff;
                border-color: #4f46e5;
            }
            #Btn_secondary:pressed {
                background: #e0e7ff;
                border-color: #4338ca;
            }
            #Btn_secondary:disabled {
                color: #9ca3af;
                background: #f3f4f6;
                border-color: #e5e7eb;
            }

            /* опасные действия (Очистить) */
            #Btn_danger {
                border: 1px solid #ef4444;
                color: #b91c1c;
                background: #fef2f2;
                font-weight: 600;
            }
            #Btn_danger:hover {
                background: #fee2e2;
                border-color: #dc2626;
            }
            #Btn_danger:pressed {
                background: #fecaca;
                border-color: #b91c1c;
            }
            #Btn_danger:disabled {
                background: #fef2f2;
                color: #fca5a5;
                border-color: #fecaca;
            }

            /* тулбар редактора */
            #ToolBtn {
                border: 1px solid #d0d7e2;
                border-radius: 14px;
                padding: 8px 12px;
                background: #ffffff;
                min-width: 40px;
                font-weight: 800;
            }
            #ToolBtn:hover {
                background: #eef2ff;
                border-color: #4f46e5;
            }
            #ToolBtn:checked {
                background: #e0e7ff;
                border-color: #4f46e5;
            }

            /* выпадающие списки: светлое поле + аккуратная кнопка-стрелка справа */
            QComboBox,
            QFontComboBox {
                min-height: 40px;
                padding: 0 44px 0 14px;
                border: 1px solid #cfd6e0;
                border-radius: 12px;
                background: #ffffff;
                font-size: 14px;
                color: #111827;
            }

            QComboBox:editable,
            QFontComboBox:editable {
                background: #ffffff;
            }

            QComboBox:focus,
            QFontComboBox:focus {
                border: 1px solid #94a3b8;
            }

            /* правая зона — SVG-кнопка как фон drop-down */
            QComboBox::drop-down,
            QFontComboBox::drop-down {
                subcontrol-origin: border;
                subcontrol-position: center right;
                width: 36px;
                border: none;
                background: transparent;
                margin-right: 4px;
                image: url(assets/combo-btn.svg);
            }

            /* down-arrow полностью убран — визуал целиком через drop-down */
            QComboBox::down-arrow,
            QFontComboBox::down-arrow {
                image: none;
                width: 0px;
                height: 0px;
            }

            QComboBox:disabled::down-arrow,
            QFontComboBox:disabled::down-arrow {
                image: none;
                width: 0px;
                height: 0px;
            }

            /* выпадающий список: светлый фон, скругления, мягкий hover/selection */
            QComboBox QAbstractItemView,
            QFontComboBox QAbstractItemView {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 12px;
                selection-background-color: #eef2ff;
                selection-color: #3730a3;
                outline: none;
                padding: 8px 4px;
                margin: 4px 0 0 0;
            }

            QComboBox QAbstractItemView::item,
            QFontComboBox QAbstractItemView::item {
                min-height: 32px;
                padding: 4px 12px;
                border-radius: 8px;
            }

            QComboBox QAbstractItemView::item:hover,
            QFontComboBox QAbstractItemView::item:hover {
                background: #f1f5f9;
            }

            QComboBox QAbstractItemView::item:selected,
            QFontComboBox QAbstractItemView::item:selected {
                background: #eef2ff;
                color: #3730a3;
            }
            """
        )

    # ---------------- UI ----------------
    def init_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(14)

        # -------- Top Bar --------
        top = QFrame()
        top.setObjectName("TopBar")
        top_layout = QHBoxLayout(top)
        top_layout.setContentsMargins(18, 14, 18, 14)
        top_layout.setSpacing(14)

        # logo
        self.logo = QLabel()
        # чуть компактнее, чтобы не выглядел растянутым
        self.logo.setFixedSize(176, 64)
        self.logo.setScaledContents(True)
        self._load_logo()
        top_layout.addWidget(self.logo, 0, Qt.AlignmentFlag.AlignVCenter)

        # app title block
        title_block = QVBoxLayout()
        title_block.setSpacing(2)
        title_row = QHBoxLayout()
        title_row.setSpacing(10)

        self.title_mark = QLabel(APP_MARK)
        self.title_mark.setStyleSheet(
            'font-size: 32px; font-weight: 800; color: #0f172a; '
            'font-family: "Segoe UI Rounded","Segoe UI","Arial"; '
            "background: transparent;"
        )
        title_row.addWidget(self.title_mark)

        self.badge_ver = Pill(APP_VERSION)
        title_row.addWidget(self.badge_ver, 0, Qt.AlignmentFlag.AlignVCenter)

        title_row.addStretch(1)
        title_block.addLayout(title_row)

        self.subtitle = QLabel(APP_SUBTITLE)
        self.subtitle.setStyleSheet("font-size: 16px; color: #64748b; padding-left: 2px; margin-top: 0px; background: transparent;")
        title_block.addWidget(self.subtitle)

        top_layout.addLayout(title_block, 0)

        # excel status pill (фиксируем ширину, чтобы при full-screen не расползалась)
        self.excel_pill = QLabel("Excel: —")
        self.excel_pill.setObjectName("ExcelPill")
        self.excel_pill.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.excel_pill.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.excel_pill.setMinimumWidth(360)
        self.excel_pill.setMaximumWidth(520)
        self.excel_pill.setMinimumHeight(48)

        # центрируем плашку за счёт “растяжек” слева/справа
        top_layout.addStretch(1)
        top_layout.addWidget(self.excel_pill, 0, Qt.AlignmentFlag.AlignVCenter)
        top_layout.addStretch(1)

        self.reload_btn = ActionBtn("Обновить", kind="default")
        self.reload_btn.clicked.connect(self.reload_excel)
        top_layout.addWidget(self.reload_btn, 0, Qt.AlignmentFlag.AlignVCenter)

        self.open_folder_btn = ActionBtn("Папка", kind="default")
        self.open_folder_btn.clicked.connect(self.open_excel_folder)
        top_layout.addWidget(self.open_folder_btn, 0, Qt.AlignmentFlag.AlignVCenter)

        self.choose_path_btn = ActionBtn("Выбрать путь", kind="default")
        self.choose_path_btn.clicked.connect(self.choose_excel_path)
        top_layout.addWidget(self.choose_path_btn, 0, Qt.AlignmentFlag.AlignVCenter)

        self.clear_btn = ActionBtn("Очистить", kind="danger")
        self.clear_btn.clicked.connect(self.clear_fields)
        top_layout.addWidget(self.clear_btn, 0, Qt.AlignmentFlag.AlignVCenter)

        root.addWidget(top)

        # -------- Content Row --------
        row = QHBoxLayout()
        row.setSpacing(14)

        # left card (input)
        self.card_left = Card()
        left_layout = QVBoxLayout(self.card_left)
        left_layout.setContentsMargins(18, 18, 18, 18)
        left_layout.setSpacing(14)

        left_title = HeaderLabel("Ввод")
        left_layout.addWidget(left_title, 0, Qt.AlignmentFlag.AlignHCenter)

        # product label
        lab_prod = QLabel("Продукт")
        lab_prod.setObjectName("FieldLabel")
        lab_prod.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        left_layout.addWidget(lab_prod)

        self.product_combo = ComboBoxFixedArrow()
        self.product_combo.setEditable(True)
        self.product_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.product_combo.setPlaceholderText("Введите продукт или выберите из списка")
        left_layout.addWidget(self.product_combo)

        # completer for product
        self.product_model = QStringListModel([])
        self.product_completer = QCompleter(self.product_model, self)
        self.product_completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.product_completer.setFilterMode(Qt.MatchFlag.MatchContains)
        self.product_completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
        self.product_combo.setCompleter(self.product_completer)

        # units + qty row
        grid = QHBoxLayout()
        grid.setSpacing(12)

        col_units = QVBoxLayout()
        lab_units = QLabel("Ед. изм.")
        lab_units.setObjectName("FieldLabel")
        lab_units.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        col_units.addWidget(lab_units)

        self.unit_combo = ComboBoxFixedArrow()
        self.unit_combo.addItem("— выберите —")
        col_units.addWidget(self.unit_combo)
        grid.addLayout(col_units, 1)

        col_qty = QVBoxLayout()
        qty_row = QHBoxLayout()
        qty_row.setSpacing(10)

        self.minus_btn = ActionBtn("−", kind="default")
        self.minus_btn.setFixedWidth(60)

        self.qty_input = QLineEdit()
        self.qty_input.setPlaceholderText("Введите количество")

        self.plus_btn = ActionBtn("+", kind="default")
        self.plus_btn.setFixedWidth(60)

        qty_row.addWidget(self.minus_btn)
        qty_row.addWidget(self.qty_input, 1)
        qty_row.addWidget(self.plus_btn)
        col_qty.addLayout(qty_row)

        grid.addLayout(col_qty, 2)
        left_layout.addLayout(grid)

        # made by
        lab_made = QLabel("Изготовил")
        lab_made.setObjectName("FieldLabel")
        lab_made.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        left_layout.addWidget(lab_made)

        self.made_combo = ComboBoxFixedArrow()
        self.made_combo.addItem("— не выбрано —")
        left_layout.addWidget(self.made_combo)

        self.made_manual = QCheckBox("Ручной ввод")
        left_layout.addWidget(self.made_manual)

        self.made_input = QLineEdit()
        self.made_input.setPlaceholderText("ФИО (можно оставить пустым)")
        self.made_input.setVisible(False)
        left_layout.addWidget(self.made_input)

        # checked by
        lab_chk = QLabel("Проверил")
        lab_chk.setObjectName("FieldLabel")
        lab_chk.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        left_layout.addWidget(lab_chk)

        self.checked_combo = ComboBoxFixedArrow()
        self.checked_combo.addItem("— не выбрано —")
        left_layout.addWidget(self.checked_combo)

        self.checked_manual = QCheckBox("Ручной ввод")
        left_layout.addWidget(self.checked_manual)

        self.checked_input = QLineEdit()
        self.checked_input.setPlaceholderText("ФИО (можно оставить пустым)")
        self.checked_input.setVisible(False)
        left_layout.addWidget(self.checked_input)

        left_layout.addStretch(1)

        # оборачиваем левую карточку в контейнер для гибкой раскладки
        left_panel = QWidget()
        left_panel_layout = QVBoxLayout(left_panel)
        left_panel_layout.setContentsMargins(0, 0, 0, 0)
        left_panel_layout.setSpacing(0)
        left_panel_layout.addWidget(self.card_left)

        # right card (preview)
        self.card_right = Card()
        right_layout = QVBoxLayout(self.card_right)
        right_layout.setContentsMargins(18, 18, 18, 18)
        right_layout.setSpacing(14)

        right_title = HeaderLabel("Предпросмотр")
        right_layout.addWidget(right_title, 0, Qt.AlignmentFlag.AlignHCenter)

        # toolbar
        tb = QHBoxLayout()
        tb.setSpacing(10)

        self.btn_font_minus = ActionBtn("A-", kind="default")
        self.btn_font_minus.setFixedWidth(60)
        self.btn_font_plus = ActionBtn("A+", kind="default")
        self.btn_font_plus.setFixedWidth(60)

        # выпадающий список размеров шрифта (как в Word)
        self.font_size_combo = ComboBoxFixedArrow()
        self.font_size_combo.setEditable(True)
        self.font_size_combo.setFixedWidth(90)
        self.font_size_combo.addItems([str(s) for s in [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]])
        self.font_size_combo.setCurrentText(str(self._base_font_size))

        # как в русском Word: Ж / К / Ч
        self.btn_bold = ToolBtn("Ж")
        self.btn_bold.setFont(QFont("Segoe UI", 11, QFont.Weight.Black))
        self.btn_italic = ToolBtn("К")
        f_it = QFont("Segoe UI", 11, QFont.Weight.Black)
        f_it.setItalic(True)
        self.btn_italic.setFont(f_it)
        self.btn_underline = ToolBtn("Ч")
        f_un = QFont("Segoe UI", 11, QFont.Weight.Black)
        f_un.setUnderline(True)
        self.btn_underline.setFont(f_un)

        self.btn_align_left = ToolBtn("≡")
        self.btn_align_center = ToolBtn("≡")
        self.btn_align_right = ToolBtn("≡")
        self.btn_align_justify = ToolBtn("≡")

        # чтобы визуально отличались (как “иконки”)
        self.btn_align_left.setStyleSheet("#ToolBtn { font-weight: 900; }")
        self.btn_align_center.setStyleSheet("#ToolBtn { font-weight: 900; }")
        self.btn_align_right.setStyleSheet("#ToolBtn { font-weight: 900; }")
        self.btn_align_justify.setStyleSheet("#ToolBtn { font-weight: 900; }")

        self.font_combo = QFontComboBox()
        self.font_combo.setObjectName("ComboWithArrow")
        self.font_combo.setFixedWidth(220)

        tb.addWidget(self.btn_font_minus)
        tb.addWidget(self.btn_font_plus)
        tb.addWidget(self.font_size_combo)
        tb.addSpacing(8)
        tb.addWidget(self.btn_bold)
        tb.addWidget(self.btn_italic)
        tb.addWidget(self.btn_underline)
        tb.addSpacing(8)
        tb.addWidget(self.btn_align_left)
        tb.addWidget(self.btn_align_center)
        tb.addWidget(self.btn_align_right)
        tb.addWidget(self.btn_align_justify)
        tb.addStretch(1)
        tb.addWidget(self.font_combo)

        right_layout.addLayout(tb)

        # preview editor
        self.preview = QTextEdit()
        self.preview.setObjectName("PreviewEditor")
        self.preview.setAcceptRichText(True)
        self.preview.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.preview.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        # Редактор предпросмотра: фиксированный “лист этикетки” 80×60 (высота больше ширины)
        # Масштаб в пикселях, пропорция 80/60 = 4/3.
        self.preview.setFixedSize(450, 600)  # initial, будет подгоняться по окну

        # стиль редактора (рамка как у этикетки)
        self.preview.setStyleSheet(
            """
            QTextEdit {
                background: #ffffff;
                border: 1px solid #cfd6e0;
                border-radius: 18px;
                padding: 18px;
            }
            """
        )

        self.preview_wrap = QFrame()
        self.preview_wrap.setObjectName("LabelWrap")
        wrap_lay = QHBoxLayout(self.preview_wrap)
        wrap_lay.setContentsMargins(12, 12, 12, 12)
        wrap_lay.addStretch(1)
        wrap_lay.addWidget(self.preview, 0, Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        wrap_lay.addStretch(1)

        right_layout.addWidget(self.preview_wrap, 1)

        # print row
        pr = QHBoxLayout()
        pr.setSpacing(12)

        # 3 одинаковых блока: Печать / Повторить / Количество (кол-во копий)
        self.print_btn = ActionBtn("ПЕЧАТЬ", kind="primary")
        self.repeat_btn = ActionBtn("Повторить", kind="secondary")
        self.repeat_btn.setEnabled(False)

        self.copies_btn = ActionBtn("Количество", kind="secondary")
        self.copies_minus = ActionBtn("−", kind="default")
        self.copies_minus.setFixedWidth(44)
        self.copies_input = QLineEdit("1")
        self.copies_input.setFixedWidth(60)
        self.copies_input.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.copies_plus = ActionBtn("+", kind="default")
        self.copies_plus.setFixedWidth(44)

        copies_wrap = QWidget()
        cw = QHBoxLayout(copies_wrap)
        cw.setContentsMargins(0, 0, 0, 0)
        cw.setSpacing(8)
        cw.addWidget(self.copies_btn, 1)
        cw.addWidget(self.copies_minus, 0)
        cw.addWidget(self.copies_input, 0)
        cw.addWidget(self.copies_plus, 0)

        # одинаковая высота
        for w in (self.print_btn, self.repeat_btn, self.copies_btn, self.copies_minus, self.copies_plus, self.copies_input):
            w.setMinimumHeight(68)

        # плоские кнопки степперов
        for w in (self.minus_btn, self.plus_btn, self.copies_minus, self.copies_plus):
            w.setObjectName("StepperBtn")

        pr.addWidget(self.print_btn, 1)
        pr.addWidget(self.repeat_btn, 1)
        pr.addWidget(copies_wrap, 1)

        right_layout.addLayout(pr)

        # оборачиваем правую карточку предпросмотра в отдельный контейнер (центр)
        center_panel = QWidget()
        center_panel_layout = QVBoxLayout(center_panel)
        center_panel_layout.setContentsMargins(0, 0, 0, 0)
        center_panel_layout.setSpacing(0)
        center_panel_layout.addWidget(self.card_right)

        # -------- History panel (right) --------
        self.history_panel = QWidget()
        self.history_panel.setObjectName("HistoryPanel")
        history_layout = QVBoxLayout(self.history_panel)
        history_layout.setContentsMargins(18, 18, 18, 18)
        history_layout.setSpacing(12)

        history_title = HeaderLabel("История")
        history_layout.addWidget(history_title, 0, Qt.AlignmentFlag.AlignHCenter)

        self.history_search = QLineEdit()
        self.history_search.setPlaceholderText("Поиск по истории")
        history_layout.addWidget(self.history_search)

        self.history_scroll = QScrollArea()
        self.history_scroll.setObjectName("HistoryScroll")
        self.history_scroll.setWidgetResizable(True)
        history_layout.addWidget(self.history_scroll, 1)

        history_scroll_content = QWidget()
        self.history_list_layout = QVBoxLayout(history_scroll_content)
        self.history_list_layout.setContentsMargins(0, 0, 0, 0)
        self.history_list_layout.setSpacing(10)

        self.history_scroll.setWidget(history_scroll_content)

        # пагинация истории
        pager_row = QHBoxLayout()
        pager_row.setSpacing(8)

        self.history_prev_btn = ActionBtn("←", kind="default")
        self.history_next_btn = ActionBtn("→", kind="default")
        self.history_page_label = QLabel("Страница 1 из 1")
        self.history_page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.history_page_label.setStyleSheet("color: #6b7280; font-size: 12px;")

        pager_row.addWidget(self.history_prev_btn, 0)
        pager_row.addWidget(self.history_page_label, 1)
        pager_row.addWidget(self.history_next_btn, 0)

        history_layout.addLayout(pager_row)

        # добавляем три панели в основной ряд с пропорциями 3:4:3
        row.addWidget(left_panel, 3)
        row.addWidget(center_panel, 4)
        row.addWidget(self.history_panel, 3)
        root.addLayout(row)

# ---------------- Signals ----------------
        self.product_combo.currentTextChanged.connect(self.on_product_changed)
        self.unit_combo.currentTextChanged.connect(self.refresh_preview)
        self.qty_input.textChanged.connect(self.refresh_preview)

        self.plus_btn.clicked.connect(self.increase_qty)
        self.minus_btn.clicked.connect(self.decrease_qty)

        self.made_manual.stateChanged.connect(self.toggle_made_mode)
        self.checked_manual.stateChanged.connect(self.toggle_checked_mode)

        self.made_combo.currentTextChanged.connect(self.refresh_preview)
        self.checked_combo.currentTextChanged.connect(self.refresh_preview)
        self.made_input.textChanged.connect(self.refresh_preview)
        self.checked_input.textChanged.connect(self.refresh_preview)

        self.print_btn.clicked.connect(self.print_label)
        self.repeat_btn.clicked.connect(self.repeat_last_print)
        self.copies_plus.clicked.connect(self.increase_copies)
        self.copies_minus.clicked.connect(self.decrease_copies)
        self.copies_input.textChanged.connect(self._sanitize_copies)

        # editor signals: чтобы не откатывало форматирование
        self.preview.textChanged.connect(self._on_preview_text_changed)
        self.preview.cursorPositionChanged.connect(self._sync_format_toolbar_from_cursor)
        self.preview.selectionChanged.connect(self._sync_format_toolbar_from_cursor)

        # toolbar actions
        self.btn_font_minus.clicked.connect(lambda: self._change_font_size(-1))
        self.btn_font_plus.clicked.connect(lambda: self._change_font_size(+1))
        self.font_size_combo.currentTextChanged.connect(self.on_font_size_combo_changed)

        self.btn_bold.clicked.connect(self._toggle_bold_on_selection)
        self.btn_italic.clicked.connect(self._toggle_italic_on_selection)
        self.btn_underline.clicked.connect(self._toggle_underline_on_selection)

        self.btn_align_left.clicked.connect(lambda: self._set_alignment(Qt.AlignmentFlag.AlignLeft))
        self.btn_align_center.clicked.connect(lambda: self._set_alignment(Qt.AlignmentFlag.AlignHCenter))
        self.btn_align_right.clicked.connect(lambda: self._set_alignment(Qt.AlignmentFlag.AlignRight))
        self.btn_align_justify.clicked.connect(lambda: self._set_alignment(Qt.AlignmentFlag.AlignJustify))

        self.font_combo.currentFontChanged.connect(self._set_font_family_on_selection)

        # история: поиск и пагинация
        self.history_search.textChanged.connect(self._on_history_search_text_changed)
        self.history_prev_btn.clicked.connect(lambda: self._change_history_page(-1))
        self.history_next_btn.clicked.connect(lambda: self._change_history_page(+1))

        # дефолт шрифта редактора
        self.preview.setFont(QFont("Segoe UI", self._base_font_size))

        # ВАЖНО: применяем стиль комбобоксов явно (иначе иногда теряются подстили)
        for cb in (self.product_combo, self.unit_combo, self.made_combo, self.checked_combo, self.font_combo):
            cb.setObjectName("ComboWithArrow")

        # инициализация состояния истории
        self._rebuild_history_view()

    def _load_logo(self):
        if os.path.exists(LOGO_PATH):
            pix = QPixmap(LOGO_PATH)
            if not pix.isNull():
                self.logo.setPixmap(pix.scaled(self.logo.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
                return
        # fallback
        self.logo.setText("")


    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._resize_label_preview()
        # адаптивность: панель истории скрываем на узких окнах
        if hasattr(self, "history_panel"):
            self.history_panel.setVisible(self.width() >= 1400)

    def _resize_label_preview(self):
        # Подгоняем “лист этикетки” под доступное место, сохраняя пропорцию 80×60 (4:3)
        if not hasattr(self, "preview_wrap"):
            return
        rect = self.preview_wrap.contentsRect()
        avail_w = rect.width()
        avail_h = rect.height()
        if avail_w < 50 or avail_h < 50:
            return

        target_w = min(avail_w, int(avail_h * 3 / 4)) - 10
        target_w = max(260, target_w)
        target_h = int(target_w * 4 / 3)

        target_w = min(target_w, 520)
        target_h = min(target_h, int(520 * 4 / 3))

        self.preview.setFixedSize(int(target_w), int(target_h))

    # ---------------- Excel / data ----------------
    def reload_excel(self, show_message: bool = True):
        try:
            current_product = self.product_combo.currentText().strip()
            current_unit = self.unit_combo.currentText()
            current_qty = self.qty_input.text().strip()

            made_manual = self.made_manual.isChecked()
            checked_manual = self.checked_manual.isChecked()
            made_text = self.made_input.text().strip()
            checked_text = self.checked_input.text().strip()
            made_combo = self.made_combo.currentText()
            checked_combo = self.checked_combo.currentText()

            products_all = load_products(self.excel_path)
            self.products = [p for p in products_all if int(p.get("active", 0)) == 1]
            self.products.sort(key=lambda x: (x.get("name") or "").lower())

            self.staff_made = [s for s in load_staff(self.excel_path, SHEET_MADE) if int(s.get("active", 0)) == 1]
            self.staff_checked = [s for s in load_staff(self.excel_path, SHEET_CHECKED) if int(s.get("active", 0)) == 1]

            # excel_loader возвращает {"name": "..."} — а у нас UI ждёт fio.
            # поэтому нормализуем к "fio".
            self.staff_made = [{"fio": (x.get("fio") or x.get("name") or "").strip(), "active": x.get("active", 1)} for x in self.staff_made]
            self.staff_checked = [{"fio": (x.get("fio") or x.get("name") or "").strip(), "active": x.get("active", 1)} for x in self.staff_checked]

            self.staff_made.sort(key=lambda x: (x.get("fio") or "").lower())
            self.staff_checked.sort(key=lambda x: (x.get("fio") or "").lower())

            self.fill_products(current_product)
            self.fill_staff()

            self.loaded_at_str = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
            self.update_excel_status()

            # восстановим поля
            if current_qty:
                self.qty_input.setText(current_qty)

            self.made_manual.setChecked(made_manual)
            self.checked_manual.setChecked(checked_manual)
            self.made_input.setText(made_text)
            self.checked_input.setText(checked_text)

            if made_combo:
                idx = self.made_combo.findText(made_combo)
                if idx >= 0:
                    self.made_combo.setCurrentIndex(idx)
            if checked_combo:
                idx = self.checked_combo.findText(checked_combo)
                if idx >= 0:
                    self.checked_combo.setCurrentIndex(idx)

            # продукт — НЕ ставим автоматически при первом запуске
            # Если человек что-то вводил — оставим. Если пусто — оставим пустым.
            if current_product:
                self.on_product_changed(current_product)
                if current_unit:
                    idxu = self.unit_combo.findText(current_unit)
                    if idxu >= 0:
                        self.unit_combo.setCurrentIndex(idxu)
            else:
                # очистим units, чтобы не подставлялись
                self.unit_combo.clear()
                self.unit_combo.addItem("— выберите —")

            if show_message:
                QMessageBox.information(self, "Готово", "Excel обновлён.")

        except Exception as e:
            QMessageBox.critical(
                self,
                "Ошибка Excel",
                f"Не удалось загрузить Excel.\n\nФайл: {self.excel_path}\nОшибка: {e}",
            )

    def fill_products(self, current_product: str | None = None):
        self.product_combo.blockSignals(True)
        self.product_combo.clear()

        names = [p["name"] for p in self.products if p.get("name")]
        for n in names:
            self.product_combo.addItem(n)

        self.product_model.setStringList(names)

        # критично: НЕ выбирать первый элемент автоматически
        self.product_combo.setCurrentIndex(-1)
        self.product_combo.setEditText("")

        if current_product:
            idx = self.product_combo.findText(current_product)
            if idx >= 0:
                self.product_combo.setCurrentIndex(idx)
            else:
                self.product_combo.setEditText(current_product)

        self.product_combo.blockSignals(False)

    def fill_staff(self):
        self.made_combo.blockSignals(True)
        self.made_combo.clear()
        self.made_combo.addItem("— не выбрано —")
        for s in self.staff_made:
            fio = (s.get("fio") or "").strip()
            if fio:
                self.made_combo.addItem(fio)
        self.made_combo.blockSignals(False)

        self.checked_combo.blockSignals(True)
        self.checked_combo.clear()
        self.checked_combo.addItem("— не выбрано —")
        for s in self.staff_checked:
            fio = (s.get("fio") or "").strip()
            if fio:
                self.checked_combo.addItem(fio)
        self.checked_combo.blockSignals(False)

    def update_excel_status(self):
        try:
            mtime = os.path.getmtime(self.excel_path)
            mtime_str = _fmt_dt_local(mtime)
        except Exception:
            mtime_str = "—"

        default_dir = os.path.dirname(EXCEL_PATH)
        current_dir = os.path.dirname(self.excel_path)
        if os.path.normpath(current_dir) != os.path.normpath(default_dir):
            short_path = current_dir
            if len(short_path) > 40:
                short_path = "..." + short_path[-37:]
            self.excel_pill.setText(f"Excel: {short_path}\nобновлено {mtime_str}")
        else:
            self.excel_pill.setText(f"Excel: обновлено {mtime_str}")

    def open_excel_folder(self):
        folder = os.path.dirname(self.excel_path)
        QDesktopServices.openUrl(QUrl.fromLocalFile(folder))

    def choose_excel_path(self):
        current_dir = os.path.dirname(self.excel_path)

        folder = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку с файлом products.xlsx",
            current_dir,
            QFileDialog.Option.ShowDirsOnly | QFileDialog.Option.DontResolveSymlinks,
        )

        if not folder:
            return

        new_path = os.path.join(folder, "products.xlsx")

        if not os.path.isfile(new_path):
            QMessageBox.warning(
                self,
                "Файл не найден",
                f"В выбранной папке не найден файл products.xlsx.\n\n"
                f"Путь: {folder}\n\n"
                f"Убедитесь, что в папке есть файл products.xlsx "
                f"с листами «продукт», «изготовил», «проверил».",
            )
            return

        self.excel_path = new_path
        settings = _load_settings()
        settings["excel_path"] = new_path
        _save_settings(settings)

        self.reload_excel(show_message=True)

    # ---------------- Helpers ----------------
    def get_product(self, name: str):
        name = (name or "").strip()
        return next((p for p in self.products if (p.get("name") or "").strip() == name), None)

    def _clear_layout(self, layout: QHBoxLayout | QVBoxLayout):
        """Полностью очищает layout от вложенных виджетов/лейаутов."""
        while layout.count():
            item = layout.takeAt(0)
            w = item.widget()
            child_layout = item.layout()
            if child_layout is not None:
                self._clear_layout(child_layout)  # type: ignore[arg-type]
            if w is not None:
                w.deleteLater()

    # ---------------- History helpers ----------------
    def _filtered_history_entries(self) -> list[dict]:
        if not self._history_filter_text:
            return list(self.history_entries)
        q = self._history_filter_text
        result: list[dict] = []
        for e in self.history_entries:
            text = " ".join(
                [
                    str(e.get("product", "")),
                    str(e.get("qty", "")),
                    str(e.get("made", "")),
                    str(e.get("checked", "")),
                    str(e.get("time", "")),
                    str(e.get("batch", "")),
                ]
            ).lower()
            if q in text:
                result.append(e)
        return result

    def _rebuild_history_view(self):
        """Перестроить список карточек истории + пагинацию."""
        if not hasattr(self, "history_list_layout"):
            return

        filtered_history = self._filtered_history_entries()
        total = len(filtered_history)
        page_size = max(1, self.history_page_size)
        pages = max(1, math.ceil(total / page_size))
        self.history_page = max(0, min(self.history_page, pages - 1))

        start = self.history_page * page_size
        end = start + page_size
        page_items = filtered_history[start:end]

        self._clear_layout(self.history_list_layout)  # type: ignore[arg-type]

        for e in page_items:
            card = QFrame()
            card.setObjectName("HistoryCard")
            card.setCursor(Qt.CursorShape.PointingHandCursor)
            card.setProperty("selected", (e.get("id") == getattr(self, "_selected_history_id", None)))
            card_layout = QVBoxLayout(card)
            card_layout.setContentsMargins(10, 8, 10, 8)
            card_layout.setSpacing(4)

            top_row = QHBoxLayout()
            top_row.setSpacing(6)

            prod_label = QLabel(str(e.get("product", "")))
            prod_label.setStyleSheet("font-weight: 600;")
            prod_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
            qty_label = QLabel(str(e.get("qty", "")))
            qty_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            qty_label.setStyleSheet("font-weight: 600; color: #111827;")
            qty_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

            top_row.addWidget(prod_label, 1)
            top_row.addWidget(qty_label, 0)

            made = str(e.get("made", ""))
            checked = str(e.get("checked", ""))
            mid_parts = [p for p in [made, checked] if p]
            mid_text = " · ".join(mid_parts) if mid_parts else ""
            mid_row = QLabel(mid_text)
            mid_row.setStyleSheet("color: #6b7280; font-size: 12px;")
            mid_row.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

            bottom_row = QHBoxLayout()
            bottom_row.setSpacing(6)

            time_label = QLabel(str(e.get("time", "")))
            time_label.setStyleSheet("color: #9ca3af; font-size: 12px;")
            time_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

            batch = str(e.get("batch", ""))
            batch_label = QLabel(f"№ {batch}" if batch else "")
            batch_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            batch_label.setStyleSheet("color: #6b7280; font-size: 12px;")
            batch_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

            bottom_row.addWidget(time_label, 1)
            bottom_row.addWidget(batch_label, 0)

            card_layout.addLayout(top_row)
            card_layout.addWidget(mid_row)
            card_layout.addLayout(bottom_row)

            # важно: ent=e чтобы не было бага замыкания
            card.mousePressEvent = (lambda ev, ent=e: self._on_history_clicked(ent))
            card.style().unpolish(card)
            card.style().polish(card)
            self.history_list_layout.addWidget(card)

        self.history_list_layout.addStretch(1)

        # обновим подпись и активность кнопок пагинации
        if hasattr(self, "history_page_label"):
            self.history_page_label.setText(f"Страница {self.history_page + 1} из {pages}")
        if hasattr(self, "history_prev_btn"):
            self.history_prev_btn.setEnabled(self.history_page > 0)
        if hasattr(self, "history_next_btn"):
            self.history_next_btn.setEnabled(self.history_page < pages - 1)

    def _on_history_search_text_changed(self, text: str):
        self._history_filter_text = (text or "").strip().lower()
        self.history_page = 0
        self._rebuild_history_view()

    def _change_history_page(self, delta: int):
        self.history_page += delta
        self._rebuild_history_view()

    def _build_history_entry_from_label(self, label, qty_display: str, unit_ui: str) -> dict:
        produced_at = getattr(label, "produced_at", datetime.now())
        preview_text = (
            f"{getattr(label, 'weekday', '')}\n"
            f"Продукт: {getattr(label, 'product_name', '')}\n"
            f"Вес/шт: {getattr(label, 'qty_value', '')} {getattr(label, 'qty_unit_ru', '')}\n"
            f"Дата/время: {format_dt(produced_at)}\n"
            f"№ партии: {getattr(label, 'batch', '')}\n"
            f"Годен до: {format_dt(getattr(label, 'expires_at', produced_at))}\n"
            f"Изготовил: {getattr(label, 'made_by', '')}\n"
            f"Проверил: {getattr(label, 'checked_by', '')}\n"
        )

        return {
            "id": time.time_ns(),
            "ts": time.time(),
            "product_name": getattr(label, "product_name", ""),
            "unit_ui": unit_ui,
            "qty_value": str(qty_display),
            "made_by": getattr(label, "made_by", ""),
            "made_manual": bool(self.made_manual.isChecked()),
            "checked_by": getattr(label, "checked_by", ""),
            "checked_manual": bool(self.checked_manual.isChecked()),
            "preview_text": preview_text,
            # поля для отображения карточки
            "product": getattr(label, "product_name", ""),
            "qty": f"{qty_display} {unit_ui}".strip(),
            "made": getattr(label, "made_by", ""),
            "checked": getattr(label, "checked_by", ""),
            "time": format_dt(produced_at),
            "batch": getattr(label, "batch", ""),
        }

    def _append_history_entry(self, entry: dict):
        # newest-first: новая запись должна быть сверху
        if not hasattr(self, "history_entries") or not isinstance(self.history_entries, list):
            self.history_entries = []
        self.history_entries.insert(0, entry)
        # при новом элементе остаёмся на 1-й странице (там самые новые), если нет фильтра
        if not self._history_filter_text:
            self.history_page = 0
        self._rebuild_history_view()

    def _on_history_clicked(self, entry: dict):
        self._selected_history_id = entry.get("id")
        self.apply_history_entry(entry)
        self._rebuild_history_view()

    def apply_history_entry(self, entry: dict):
        self._loading_from_history = True
        try:
            # блокируем сигналы формы, чтобы не запускать лишние обработчики
            to_block = (
                self.product_combo,
                self.qty_input,
                self.made_manual,
                self.made_combo,
                self.made_input,
                self.checked_manual,
                self.checked_combo,
                self.checked_input,
            )
            for w in to_block:
                w.blockSignals(True)

            product_name = str(entry.get("product_name") or entry.get("product") or "").strip()
            unit_ui = str(entry.get("unit_ui") or "").strip()
            qty_value = str(entry.get("qty_value") or "").strip()
            made_by = str(entry.get("made_by") or entry.get("made") or "").strip()
            made_manual = bool(entry.get("made_manual", False))
            checked_by = str(entry.get("checked_by") or entry.get("checked") or "").strip()
            checked_manual = bool(entry.get("checked_manual", False))

            # продукт
            idx = self.product_combo.findText(product_name)
            if idx >= 0:
                self.product_combo.setCurrentIndex(idx)
            else:
                self.product_combo.setEditText(product_name)

            # наполняем единицы измерения через существующую логику (refresh_preview подавлен guard-ом)
            self.on_product_changed(product_name)

            # единица измерения
            self.unit_combo.blockSignals(True)
            try:
                if unit_ui:
                    idxu = self.unit_combo.findText(unit_ui)
                    if idxu >= 0:
                        self.unit_combo.setCurrentIndex(idxu)
            finally:
                self.unit_combo.blockSignals(False)

            # количество
            self.qty_input.setText(qty_value)

            # изготовил
            self.made_manual.setChecked(made_manual)
            self.toggle_made_mode()
            if made_manual:
                self.made_input.setText(made_by)
            else:
                idxm = self.made_combo.findText(made_by)
                if idxm >= 0:
                    self.made_combo.setCurrentIndex(idxm)
                elif made_by:
                    self.made_manual.setChecked(True)
                    self.toggle_made_mode()
                    self.made_input.setText(made_by)

            # проверил
            self.checked_manual.setChecked(checked_manual)
            self.toggle_checked_mode()
            if checked_manual:
                self.checked_input.setText(checked_by)
            else:
                idxc = self.checked_combo.findText(checked_by)
                if idxc >= 0:
                    self.checked_combo.setCurrentIndex(idxc)
                elif checked_by:
                    self.checked_manual.setChecked(True)
                    self.toggle_checked_mode()
                    self.checked_input.setText(checked_by)

        finally:
            for w in (
                self.product_combo,
                self.qty_input,
                self.made_manual,
                self.made_combo,
                self.made_input,
                self.checked_manual,
                self.checked_combo,
                self.checked_input,
            ):
                w.blockSignals(False)
            self._loading_from_history = False

        # предпросмотр выбранной записи (HTML сохраняет форматирование)
        preview_html = entry.get("preview_html")
        preview_text = entry.get("preview_text")
        if isinstance(preview_html, str) and preview_html.strip():
            self._set_preview_html_programmatically(preview_html)
            self._user_edited_preview = True
            _, can_print = self._build_label_plain_text()
            self.print_btn.setEnabled(can_print)
        elif isinstance(preview_text, str) and preview_text.strip():
            self._set_preview_text_programmatically(preview_text)
            self._user_edited_preview = True
            _, can_print = self._build_label_plain_text()
            self.print_btn.setEnabled(can_print)
        else:
            self._user_edited_preview = False
            self.refresh_preview()

    # ---------------- UI actions ----------------
    def toggle_made_mode(self):
        manual = self.made_manual.isChecked()
        self.made_combo.setVisible(not manual)
        self.made_input.setVisible(manual)
        if not getattr(self, "_loading_from_history", False):
            self.refresh_preview()

    def toggle_checked_mode(self):
        manual = self.checked_manual.isChecked()
        self.checked_combo.setVisible(not manual)
        self.checked_input.setVisible(manual)
        if not getattr(self, "_loading_from_history", False):
            self.refresh_preview()

    def on_product_changed(self, product_name: str):
        self.unit_combo.blockSignals(True)
        self.unit_combo.clear()
        self.unit_combo.addItem("— выберите —")

        product = self.get_product(product_name)
        if product:
            units = product.get("allowed_units", [])
            if isinstance(units, str):
                units = [u.strip() for u in units.split(",") if u.strip()]

            for u in units:
                if u == "kg":
                    self.unit_combo.addItem("кг")
                elif u == "pcs":
                    self.unit_combo.addItem("шт")
                else:
                    self.unit_combo.addItem(u)

        self.unit_combo.blockSignals(False)
        if not getattr(self, "_loading_from_history", False):
            self.refresh_preview()

    def _step_for_unit(self) -> float:
        unit_ru = self.unit_combo.currentText()
        return 1.0 if unit_ru == "шт" else 0.1

    def increase_qty(self):
        step = self._step_for_unit()
        try:
            value = float(self.qty_input.text().replace(",", "."))
        except Exception:
            value = 0.0
        value += step
        self.qty_input.setText(str(round(value, 3)).rstrip("0").rstrip("."))

    def decrease_qty(self):
        step = self._step_for_unit()
        try:
            value = float(self.qty_input.text().replace(",", "."))
        except Exception:
            value = 0.0
        value = max(0.0, value - step)
        self.qty_input.setText(str(round(value, 3)).rstrip("0").rstrip("."))

    def clear_fields(self):
        # НЕ выбираем продукт, очищаем всё
        self.product_combo.setCurrentIndex(-1)
        self.product_combo.setEditText("")

        self.unit_combo.setCurrentIndex(0)
        self.qty_input.clear()

        self.made_manual.setChecked(False)
        self.checked_manual.setChecked(False)

        self.made_combo.setCurrentIndex(0)
        self.checked_combo.setCurrentIndex(0)

        self.made_input.clear()
        self.checked_input.clear()

        # сброс редактора к авто-превью
        self._user_edited_preview = False

        self.refresh_preview()

    # ---------------- Preview / validation ----------------
    def _unit_code_from_ui(self, unit_text: str) -> str | None:
        if unit_text == "кг":
            return "kg"
        if unit_text == "шт":
            return "pcs"
        return None

    def _made_value(self) -> str:
        if self.made_manual.isChecked():
            return self.made_input.text().strip()
        val = self.made_combo.currentText().strip()
        return "" if val.startswith("—") else val

    def _checked_value(self) -> str:
        if self.checked_manual.isChecked():
            return self.checked_input.text().strip()
        val = self.checked_combo.currentText().strip()
        return "" if val.startswith("—") else val

    _WEEKDAYS = {"ПОНЕДЕЛЬНИК", "ВТОРНИК", "СРЕДА", "ЧЕТВЕРГ", "ПЯТНИЦА", "СУББОТА", "ВОСКРЕСЕНЬЕ"}

    def _set_preview_text_programmatically(self, text: str):
        """
        Вставка текста в редактор с дефолтным форматированием:
        - весь текст: размер _base_font_size (20)
        - первая строка (день недели): жирный, размер 26, по центру
        """
        self._updating_preview = True
        try:
            self.preview.blockSignals(True)
            self.preview.setPlainText(text)

            font_family = self.font_combo.currentFont().family()
            cursor = self.preview.textCursor()

            # 1) весь текст — базовый шрифт
            cursor.select(QTextCursor.SelectionType.Document)
            fmt_base = QTextCharFormat()
            fmt_base.setFont(QFont(font_family, self._base_font_size))
            cursor.mergeCharFormat(fmt_base)
            cursor.clearSelection()

            # 2) первая строка — жирный, 26pt, по центру (only if it’s a weekday)
            first_line = text.strip().split("\n")[0].strip()
            if first_line in self._WEEKDAYS:
                cursor.movePosition(QTextCursor.MoveOperation.Start)
                cursor.movePosition(QTextCursor.MoveOperation.EndOfBlock, QTextCursor.MoveMode.KeepAnchor)

                fmt_weekday = QTextCharFormat()
                fmt_weekday.setFont(QFont(font_family, 26, QFont.Weight.Bold))
                cursor.mergeCharFormat(fmt_weekday)

                block_fmt = QTextBlockFormat()
                block_fmt.setAlignment(Qt.AlignmentFlag.AlignHCenter)
                cursor.mergeBlockFormat(block_fmt)

            cursor.clearSelection()
            cursor.movePosition(QTextCursor.MoveOperation.Start)
            self.preview.setTextCursor(cursor)
        finally:
            self.preview.blockSignals(False)
            self._updating_preview = False

    def _set_preview_html_programmatically(self, html: str):
        """
        Восстановление HTML-форматирования в редактор (жирный, курсив, выравнивание и т.д.)
        без срабатывания «пользовательского редактирования».
        """
        self._updating_preview = True
        try:
            self.preview.blockSignals(True)
            self.preview.setHtml(html)
            cursor = self.preview.textCursor()
            cursor.movePosition(QTextCursor.MoveOperation.Start)
            self.preview.setTextCursor(cursor)
        finally:
            self.preview.blockSignals(False)
            self._updating_preview = False

    def _build_label_plain_text(self) -> tuple[str, bool]:
        """
        Возвращает (text, can_print)
        """
        product_name = self.product_combo.currentText().strip()
        product = self.get_product(product_name)

        unit_ui = self.unit_combo.currentText()
        unit_code = self._unit_code_from_ui(unit_ui)
        qty = self.qty_input.text().strip().replace(",", ".")

        if not product:
            return ("Выберите продукт.", False)

        if unit_code is None:
            return ("Выберите единицу измерения (кг или шт).", False)

        if not qty:
            return ("Введите количество.", False)

        try:
            qty_float = float(qty)
        except Exception:
            return ("Количество должно быть числом (например 2 или 0.5).", False)

        if qty_float <= 0:
            return ("Количество должно быть больше 0.", False)

        if unit_code == "pcs" and abs(qty_float - round(qty_float)) > 1e-9:
            return ("Для 'шт' количество должно быть целым числом.", False)

        made_by = self._made_value()
        checked_by = self._checked_value()

        label = build_label(
            product_name=product["name"],
            shelf_life_hours=product["shelf_life_hours"],
            qty_value=str(qty_float).rstrip("0").rstrip("."),
            unit=unit_code,
            made_by=made_by,
            checked_by=checked_by,
        )

        text = (
            f"{label.weekday}\n"
            f"Продукт: {label.product_name}\n"
            f"Вес/шт: {label.qty_value} {label.qty_unit_ru}\n"
            f"Дата/время: {format_dt(label.produced_at)}\n"
            f"№ партии: {label.batch}\n"
            f"Годен до: {format_dt(label.expires_at)}\n"
            f"Изготовил: {label.made_by}\n"
            f"Проверил: {label.checked_by}\n"
        )
        return (text, True)

    def refresh_preview(self):
        if getattr(self, "_loading_from_history", False):
            return
        text, can_print = self._build_label_plain_text()
        self.print_btn.setEnabled(can_print)

        # если пользователь уже начал редактировать предпросмотр — НЕ перетираем
        # (именно это раньше выглядело как “откат обратно”)
        if self._user_edited_preview:
            return

        self._set_preview_text_programmatically(text)

    def _on_preview_text_changed(self):
        if self._updating_preview:
            return
        # пользователь что-то менял руками — больше не перетираем автогенерацией
        self._user_edited_preview = True

    # ---------------- Editor formatting (selection-only) ----------------
    def _merge_format_on_selection(self, fmt: QTextCharFormat):
        cursor = self.preview.textCursor()
        if not cursor.hasSelection():
            # если нет выделения — применяем к текущему слову (как в Word)
            cursor.select(QTextCursor.SelectionType.WordUnderCursor)
        cursor.mergeCharFormat(fmt)
        self.preview.mergeCurrentCharFormat(fmt)

    def _toggle_bold_on_selection(self):
        fmt = QTextCharFormat()
        # состояние берём из самой кнопки (как в Word)
        desired_bold = self.btn_bold.isChecked()
        fmt.setFontWeight(QFont.Weight.Bold if desired_bold else QFont.Weight.Normal)
        self._merge_format_on_selection(fmt)
        self._sync_format_toolbar_from_cursor()

    def _toggle_italic_on_selection(self):
        fmt = QTextCharFormat()
        desired = self.btn_italic.isChecked()
        fmt.setFontItalic(desired)
        self._merge_format_on_selection(fmt)
        self._sync_format_toolbar_from_cursor()

    def _toggle_underline_on_selection(self):
        fmt = QTextCharFormat()
        desired = self.btn_underline.isChecked()
        fmt.setFontUnderline(desired)
        self._merge_format_on_selection(fmt)
        self._sync_format_toolbar_from_cursor()

    def _sync_format_toolbar_from_cursor(self):
        """Синхронизация состояний Ж/К/Ч с текущим форматированием, как в Word."""
        cursor = self.preview.textCursor()
        fmt = cursor.charFormat() if cursor.charFormat().isValid() else self.preview.currentCharFormat()

        self.btn_bold.blockSignals(True)
        self.btn_italic.blockSignals(True)
        self.btn_underline.blockSignals(True)
        try:
            self.btn_bold.setChecked(fmt.fontWeight() >= QFont.Weight.Bold)
            self.btn_italic.setChecked(fmt.fontItalic())
            self.btn_underline.setChecked(fmt.fontUnderline())
        finally:
            self.btn_bold.blockSignals(False)
            self.btn_italic.blockSignals(False)
            self.btn_underline.blockSignals(False)

        # параллельно обновляем UI размера шрифта
        self._sync_font_size_from_cursor()

    def _sync_font_size_from_cursor(self):
        """
        Синхронизировать поле размера шрифта с текущей позицией курсора
        (как в Word).
        """
        if not hasattr(self, "font_size_combo"):
            return

        cursor = self.preview.textCursor()
        fmt = cursor.charFormat()

        size = fmt.fontPointSize()
        if size <= 0:
            # если размер не задан явно в формате — считаем, что используется базовый
            size = float(self._base_font_size)

        size_int = int(round(size))
        if size_int <= 0:
            size_int = self._base_font_size

        self._base_font_size = size_int

        self.font_size_combo.blockSignals(True)
        self.font_size_combo.setCurrentText(str(size_int))
        self.font_size_combo.blockSignals(False)

    def _set_alignment(self, align_flag: Qt.AlignmentFlag):
        # alignment в QTextEdit — на уровне блока.
        # применяем к текущему блоку / блокам выделения (как в редакторах)
        if align_flag == Qt.AlignmentFlag.AlignLeft:
            self.preview.setAlignment(Qt.AlignmentFlag.AlignLeft)
        elif align_flag == Qt.AlignmentFlag.AlignHCenter:
            self.preview.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        elif align_flag == Qt.AlignmentFlag.AlignRight:
            self.preview.setAlignment(Qt.AlignmentFlag.AlignRight)
        elif align_flag == Qt.AlignmentFlag.AlignJustify:
            self.preview.setAlignment(Qt.AlignmentFlag.AlignJustify)

    def _set_font_family_on_selection(self, font: QFont):
        fmt = QTextCharFormat()
        fmt.setFontFamily(font.family())
        self._merge_format_on_selection(fmt)

    def _change_font_size(self, delta: int):
        self._base_font_size = max(8, min(72, self._base_font_size + delta))
        fmt = QTextCharFormat()
        fmt.setFontPointSize(float(self._base_font_size))
        self._merge_format_on_selection(fmt)
        # синхронизируем комбобокс размера
        if hasattr(self, "font_size_combo"):
            self.font_size_combo.blockSignals(True)
            self.font_size_combo.setCurrentText(str(self._base_font_size))
            self.font_size_combo.blockSignals(False)


    def on_font_size_combo_changed(self, text: str):
        t = (text or "").strip()
        if not t:
            return
        try:
            size = int(float(t))
        except Exception:
            return
        size = max(8, min(72, size))
        self._base_font_size = size
        fmt = QTextCharFormat()
        fmt.setFontPointSize(float(size))
        self._merge_format_on_selection(fmt)
        # синхронизируем комбобокс, если пользователь ввёл число
        if hasattr(self, "font_size_combo"):
            self.font_size_combo.blockSignals(True)
            self.font_size_combo.setCurrentText(str(size))
            self.font_size_combo.blockSignals(False)

    def _get_copies(self) -> int:
        if not hasattr(self, "copies_input"):
            return 1
        v = _safe_int(self.copies_input.text().strip(), 1)
        return max(1, min(999, v))

    def _sanitize_copies(self):
        if not hasattr(self, "copies_input"):
            return
        v = _safe_int(self.copies_input.text().strip(), 1)
        v = max(1, min(999, v))
        if self.copies_input.text().strip() != str(v):
            self.copies_input.blockSignals(True)
            self.copies_input.setText(str(v))
            self.copies_input.blockSignals(False)

    def increase_copies(self):
        if not hasattr(self, "copies_input"):
            return
        self.copies_input.setText(str(self._get_copies() + 1))

    def decrease_copies(self):
        if not hasattr(self, "copies_input"):
            return
        self.copies_input.setText(str(max(1, self._get_copies() - 1)))

    def _apply_copies_to_tspl(self, tspl: str, copies: int) -> str:
        lines = tspl.strip().splitlines()
        for i in range(len(lines) - 1, -1, -1):
            if lines[i].strip().upper().startswith("PRINT"):
                lines[i] = f"PRINT {copies}"
                return "\n".join(lines).strip()
        return (tspl.strip() + f"\nPRINT {copies}").strip()

# ---------------- Rendering preview → TSPL bitmap ----------------
    def _render_preview_to_tspl_bytes(
        self,
        label_w_mm: float = 58,
        label_h_mm: float = 80,
        dpi: int = 203,
        density: int = 10,
        speed: int = 4,
        threshold: int = 200,
        copies: int = 1,
    ) -> bytes:
        """
        WYSIWYG-рендеринг: берём QTextDocument из preview-редактора
        в тех же координатах, в которых пользователь видит его на экране,
        и масштабируем в пиксели принтера (203 DPI, 58×80 мм).
        Текст переносится ровно в тех же местах, что и в предпросмотре.
        """
        w_px = int(round(label_w_mm / 25.4 * dpi))
        h_px = int(round(label_h_mm / 25.4 * dpi))

        # клонируем документ, чтобы не трогать редактор
        doc = self.preview.document().clone()

        # используем размер viewport'а редактора — это ровно то,
        # что пользователь видит; текст уже уложен по этой ширине
        vp_w = self.preview.viewport().width()
        vp_h = self.preview.viewport().height()

        doc.setPageSize(QSizeF(vp_w, vp_h))

        # масштаб: из экранных пикселей редактора → в пиксели принтера
        scale_x = w_px / vp_w
        scale_y = h_px / vp_h

        # рисуем на QImage в принтерном разрешении
        img = QImage(w_px, h_px, QImage.Format.Format_RGB32)
        img.fill(QColor(255, 255, 255))

        painter = QPainter(img)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.setRenderHint(QPainter.RenderHint.TextAntialiasing, True)
        painter.scale(scale_x, scale_y)
        doc.drawContents(painter)
        painter.end()

        # --- конвертация в 1-bit TSPL bitmap ---
        width_bytes = (w_px + 7) // 8
        raster = bytearray()

        for y in range(h_px):
            for xb in range(width_bytes):
                byte_val = 0
                for bit in range(8):
                    x = xb * 8 + bit
                    if x >= w_px:
                        byte_val |= (1 << (7 - bit))
                        continue
                    pixel = img.pixel(x, y)
                    gray = ((pixel >> 16) & 0xFF) + ((pixel >> 8) & 0xFF) + (pixel & 0xFF)
                    gray = gray // 3
                    if gray > threshold:
                        byte_val |= (1 << (7 - bit))
                raster.append(byte_val)

        header = (
            f"SIZE {label_w_mm} mm, {label_h_mm} mm\r\n"
            f"GAP 2 mm, 0 mm\r\n"
            f"SPEED {speed}\r\n"
            f"DENSITY {density}\r\n"
            f"DIRECTION 1\r\n"
            f"REFERENCE 0,0\r\n"
            f"CLS\r\n"
            f"BITMAP 0,0,{width_bytes},{h_px},0,"
        ).encode("ascii")

        tail = f"\r\nPRINT {copies}\r\n".encode("ascii")

        return header + bytes(raster) + tail

# ---------------- Printing ----------------
    def print_label(self):
        preview_text = self.preview.toPlainText()
        preview_html = self.preview.toHtml()
        if not (preview_text or "").strip():
            QMessageBox.warning(self, "Печать", "Предпросмотр пуст. Заполните форму или введите текст в предпросмотр.")
            return

        printer_name = win32print.GetDefaultPrinter()

        try:
            tspl_bytes = self._render_preview_to_tspl_bytes(threshold=200, copies=self._get_copies())
            n_bytes = print_raw(printer_name, tspl_bytes)
            print("SENDING BITMAP...", n_bytes, "bytes")
        except Exception as e:
            QMessageBox.warning(self, "Печать", f"Не удалось отправить на печать:\n{e}")
            return

        self.repeat_btn.setEnabled(True)
        self.last_printed_preview_text = preview_text
        self._last_printed_tspl_bytes = tspl_bytes

        # только после успешной печати — в историю (старая структура entry для отображения)
        product_name = self.product_combo.currentText().strip()
        product = self.get_product(product_name)
        unit_ui = self.unit_combo.currentText()
        unit_code = self._unit_code_from_ui(unit_ui)
        qty = self.qty_input.text().strip().replace(",", ".")

        if product and unit_code and qty:
            try:
                label = build_label(
                    product_name=product["name"],
                    shelf_life_hours=product["shelf_life_hours"],
                    qty_value=qty,
                    unit=unit_code,
                    made_by=self._made_value(),
                    checked_by=self._checked_value(),
                )
                entry = self._build_history_entry_from_label(
                    label,
                    qty_display=qty,
                    unit_ui=unit_ui,
                )
                entry["preview_text"] = preview_text
                entry["preview_html"] = preview_html
            except Exception:
                lines = preview_text.strip().splitlines()
                first_line = (lines[0] if lines else "").strip() or "Этикетка"
                entry = {
                    "id": time.time_ns(),
                    "ts": time.time(),
                    "preview_text": preview_text,
                    "preview_html": preview_html,
                    "product": first_line,
                    "qty": "",
                    "made": "",
                    "checked": "",
                    "time": format_dt(datetime.now()),
                    "batch": "",
                }
        else:
            lines = preview_text.strip().splitlines()
            first_line = (lines[0] if lines else "").strip() or "Этикетка"
            entry = {
                "id": time.time_ns(),
                "ts": time.time(),
                "preview_text": preview_text,
                "preview_html": preview_html,
                "product": first_line,
                "qty": "",
                "made": "",
                "checked": "",
                "time": format_dt(datetime.now()),
                "batch": "",
            }
        self.last_history_entry = entry
        self._append_history_entry(entry)

    def repeat_last_print(self):
        tspl_bytes = getattr(self, "_last_printed_tspl_bytes", None)
        if not tspl_bytes:
            return
        try:
            printer_name = win32print.GetDefaultPrinter()
            # подставляем текущее количество копий
            copies = self._get_copies()
            if copies != 1:
                # заменяем PRINT N в сохранённых байтах
                tspl_bytes = tspl_bytes.rsplit(b"\r\nPRINT ", 1)[0] + f"\r\nPRINT {copies}\r\n".encode("ascii")
            print_raw(printer_name, tspl_bytes)
            if self.last_history_entry is not None:
                e = dict(self.last_history_entry)
                e["id"] = time.time_ns()
                e["ts"] = time.time()
                self._append_history_entry(e)
        except Exception as e:
            QMessageBox.warning(self, "Повтор", f"Не удалось повторить печать:\n{e}")


def main():
    # OpenGL rendering для качественного отображения splash video (2K/4K)
    fmt = QSurfaceFormat()
    fmt.setRenderableType(QSurfaceFormat.RenderableType.OpenGL)
    QSurfaceFormat.setDefaultFormat(fmt)

    app = QApplication(sys.argv)
    main_window = None

    def on_splash_finished():
        nonlocal main_window
        main_window = MirlisMarkApp()
        main_window.showMaximized()

    video_path = SPLASH_VIDEO_PATH
    if os.path.isfile(video_path):
        splash = SplashVideo(video_path, on_finished_callback=on_splash_finished)
        splash.showFullScreen()
        splash.play()
    else:
        on_splash_finished()

    sys.exit(app.exec())

if __name__ == "__main__":
    main()