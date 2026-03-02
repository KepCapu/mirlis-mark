# main.py
# Mirlis Mark — Система маркировки
# UI: “как на картинке” + редактор предпросмотра + ВИДИМЫЕ стрелочки в выпадающих списках
#
# ВАЖНО:
# - excel_loader.py / label_logic.py / printer.py НЕ ТРОГАЕМ
# - логотип берём по пути: D:\mirlis_mark\Mirlis software logo.png

import sys
import os
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
)
from PyQt6.QtCore import QTimer, Qt, QUrl, QSize
from PyQt6.QtGui import (
    QDesktopServices,
    QIcon,
    QPixmap,
    QFont,
    QTextCharFormat,
    QTextCursor,
)
from PyQt6.QtCore import QStringListModel

from excel_loader import load_products, load_staff
from label_logic import build_label, format_dt, generate_tspl
from printer import print_raw
import win32print


# -------------------- CONFIG --------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_PATH = os.path.join(BASE_DIR, "data_sources", "products.xlsx")
SHEET_PRODUCTS = "продукт"
SHEET_MADE = "изготовил"
SHEET_CHECKED = "проверил"

LOGO_PATH = os.path.join(BASE_DIR, "assets", "logo.png")

APP_TITLE = "Mirlis Mark — Система маркировки"
APP_MARK = "Mark"
APP_VERSION = "1.0"
APP_SUBTITLE = "Система маркировки"


# -------------------- HELPERS --------------------
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


# -------------------- MAIN APP --------------------
class MirlisMarkApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)

        # окно растягиваемое, но с адекватным минимумом
        self.setMinimumSize(1100, 650)

        # печать
        self.last_tspl: str | None = None
        self.last_tspl_human: str | None = None

        # данные
        self.products = []
        self.staff_made = []
        self.staff_checked = []
        self.loaded_at_str = "—"

        # флаги чтобы редактор НЕ “откатывал” форматирование
        self._updating_preview = False
        self._user_edited_preview = False

        # базовый размер шрифта в редакторе
        self._base_font_size = 12

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
            }
            QPushButton:hover {
                background: #eef2ff;
            }
            QPushButton:pressed {
                background: #e0e7ff;
                border-color: #4f46e5;
            }
            QPushButton:disabled {
                background: #f3f4f6;
                color: #9ca3af;
                border-color: #e5e7eb;
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

            /* выпадающие списки с визуальной стрелкой */
            QComboBox,
            QFontComboBox {
                min-height: 40px;
                padding: 0 40px 0 12px; /* место под стрелку справа */
                border: 1px solid #cfd6e0;
                border-radius: 12px;
                background: #f9fafb;
                font-size: 14px;
            }

            QComboBox:editable,
            QFontComboBox:editable {
                background: #ffffff;
            }

            QComboBox::drop-down,
            QFontComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 40px;
                border-left: 1px solid #cfd6e0;
                background: #f3f4f6;
                border-top-right-radius: 12px;
                border-bottom-right-radius: 12px;
            }

            QComboBox::down-arrow {
                image: url(assets/arrow-down.svg);
                width: 18px;
                height: 18px;
            }

            QFontComboBox::down-arrow {
                image: none;
                width: 0;
                height: 0;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 6px solid #6b7280; /* ▾ */
            }

            /* выпадающий список */
            QComboBox QAbstractItemView,
            QFontComboBox QAbstractItemView {
                background: #ffffff;
                border: 1px solid #cfd6e0;
                border-radius: 10px;
                selection-background-color: #e5f3ec;
                selection-color: #14532d;
                outline: none;
                padding: 6px;
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
        self.logo.setFixedSize(220, 80)
        self.logo.setScaledContents(True)
        self._load_logo()
        top_layout.addWidget(self.logo, 0, Qt.AlignmentFlag.AlignVCenter)

        # app title block
        title_block = QVBoxLayout()
        title_row = QHBoxLayout()
        title_row.setSpacing(10)

        self.title_mark = QLabel(APP_MARK)
        self.title_mark.setStyleSheet("font-size: 32px; font-weight: 900; color: #0f172a;; background: transparent;")
        title_row.addWidget(self.title_mark)

        self.badge_ver = Pill(APP_VERSION)
        title_row.addWidget(self.badge_ver, 0, Qt.AlignmentFlag.AlignVCenter)

        title_row.addStretch(1)
        title_block.addLayout(title_row)

        self.subtitle = QLabel(APP_SUBTITLE)
        self.subtitle.setStyleSheet("font-size: 16px; color: #64748b; padding-left: 2px; background: transparent;")
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

        # тестовые карточки истории (пока без реальных данных)
        sample_items = [
            {
                "product": "Гречка отварная",
                "qty": "3 кг",
                "made": "Буров Велорий",
                "checked": "Автономов Дмитрий",
                "time": "02.03.2026 14:54",
                "batch": "020326",
            },
            {
                "product": "Курица запечённая",
                "qty": "5 кг",
                "made": "Иванова Мария",
                "checked": "Петров Сергей",
                "time": "02.03.2026 13:10",
                "batch": "020325",
            },
            {
                "product": "Рис отварной",
                "qty": "2 кг",
                "made": "Сидоров Алексей",
                "checked": "Кузнецова Анна",
                "time": "02.03.2026 12:30",
                "batch": "020324",
            },
        ]

        for item in sample_items:
            card = QFrame()
            card.setObjectName("HistoryCard")
            card_layout = QVBoxLayout(card)
            card_layout.setContentsMargins(10, 8, 10, 8)
            card_layout.setSpacing(4)

            top_row = QHBoxLayout()
            top_row.setSpacing(6)

            prod_label = QLabel(item["product"])
            prod_label.setStyleSheet("font-weight: 600;")
            qty_label = QLabel(item["qty"])
            qty_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            qty_label.setStyleSheet("font-weight: 600; color: #111827;")

            top_row.addWidget(prod_label, 1)
            top_row.addWidget(qty_label, 0)

            mid_row = QLabel(f"{item['made']} · {item['checked']}")
            mid_row.setStyleSheet("color: #6b7280; font-size: 12px;")

            bottom_row = QHBoxLayout()
            bottom_row.setSpacing(6)

            time_label = QLabel(item["time"])
            time_label.setStyleSheet("color: #9ca3af; font-size: 12px;")

            batch_label = QLabel(f"№ {item['batch']}")
            batch_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            batch_label.setStyleSheet("color: #6b7280; font-size: 12px;")

            bottom_row.addWidget(time_label, 1)
            bottom_row.addWidget(batch_label, 0)

            card_layout.addLayout(top_row)
            card_layout.addWidget(mid_row)
            card_layout.addLayout(bottom_row)

            self.history_list_layout.addWidget(card)

        self.history_list_layout.addStretch(1)
        self.history_scroll.setWidget(history_scroll_content)

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

        # дефолт шрифта редактора
        self.preview.setFont(QFont("Segoe UI", self._base_font_size))

        # ВАЖНО: применяем стиль комбобоксов явно (иначе иногда теряются подстили)
        for cb in (self.product_combo, self.unit_combo, self.made_combo, self.checked_combo, self.font_combo):
            cb.setObjectName("ComboWithArrow")

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

            products_all = load_products(EXCEL_PATH)
            self.products = [p for p in products_all if int(p.get("active", 0)) == 1]
            self.products.sort(key=lambda x: (x.get("name") or "").lower())

            self.staff_made = [s for s in load_staff(EXCEL_PATH, SHEET_MADE) if int(s.get("active", 0)) == 1]
            self.staff_checked = [s for s in load_staff(EXCEL_PATH, SHEET_CHECKED) if int(s.get("active", 0)) == 1]

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
                f"Не удалось загрузить Excel.\n\nФайл: {EXCEL_PATH}\nОшибка: {e}",
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
            mtime = os.path.getmtime(EXCEL_PATH)
            mtime_str = _fmt_dt_local(mtime)
        except Exception:
            mtime_str = "—"

        # как в макете: “Excel: обновлено …”
        self.excel_pill.setText(f"Excel: обновлено {mtime_str}")

    def open_excel_folder(self):
        folder = os.path.dirname(EXCEL_PATH)
        QDesktopServices.openUrl(QUrl.fromLocalFile(folder))

    # ---------------- Helpers ----------------
    def get_product(self, name: str):
        name = (name or "").strip()
        return next((p for p in self.products if (p.get("name") or "").strip() == name), None)

    # ---------------- UI actions ----------------
    def toggle_made_mode(self):
        manual = self.made_manual.isChecked()
        self.made_combo.setVisible(not manual)
        self.made_input.setVisible(manual)
        self.refresh_preview()

    def toggle_checked_mode(self):
        manual = self.checked_manual.isChecked()
        self.checked_combo.setVisible(not manual)
        self.checked_input.setVisible(manual)
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

    def _set_preview_text_programmatically(self, text: str):
        """
        Вставка текста в редактор так, чтобы:
        - не сбивалось выделение
        - не считалось “пользовательским редактированием”
        - не откатывалось форматирование
        """
        self._updating_preview = True
        try:
            self.preview.blockSignals(True)
            self.preview.setPlainText(text)
            # базовый шрифт документа
            self.preview.selectAll()
            cursor = self.preview.textCursor()
            fmt = QTextCharFormat()
            fmt.setFont(QFont(self.font_combo.currentFont().family(), self._base_font_size))
            cursor.mergeCharFormat(fmt)
            cursor.clearSelection()
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
        cursor = self.preview.textCursor()
        if not cursor.hasSelection():
            cursor.select(QTextCursor.SelectionType.WordUnderCursor)

        current = self.preview.currentCharFormat().fontWeight()
        new_weight = QFont.Weight.Normal if current >= QFont.Weight.Bold else QFont.Weight.Bold

        fmt = QTextCharFormat()
        fmt.setFontWeight(new_weight)
        self._merge_format_on_selection(fmt)

    def _toggle_italic_on_selection(self):
        fmt = QTextCharFormat()
        fmt.setFontItalic(not self.preview.currentCharFormat().fontItalic())
        self._merge_format_on_selection(fmt)

    def _toggle_underline_on_selection(self):
        fmt = QTextCharFormat()
        fmt.setFontUnderline(not self.preview.currentCharFormat().fontUnderline())
        self._merge_format_on_selection(fmt)

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

# ---------------- Printing ----------------
    def print_label(self):
        product_name = self.product_combo.currentText().strip()
        product = self.get_product(product_name)

        unit_code = self._unit_code_from_ui(self.unit_combo.currentText())
        qty = self.qty_input.text().strip().replace(",", ".")

        if not product or unit_code is None or not qty:
            return

        made_by = self._made_value()
        checked_by = self._checked_value()

        label = build_label(
            product_name=product["name"],
            shelf_life_hours=product["shelf_life_hours"],
            qty_value=qty,
            unit=unit_code,
            made_by=made_by,
            checked_by=checked_by,
        )

        tspl_base = generate_tspl(label)
        tspl = self._apply_copies_to_tspl(tspl_base, self._get_copies())

        self.last_tspl = tspl_base
        self.last_tspl_human = f"{product['name']} / {qty} {self.unit_combo.currentText()}"
        self.repeat_btn.setEnabled(True)

        try:
            printer_name = win32print.GetDefaultPrinter()
            print_raw(printer_name, tspl)
        except Exception as e:
            QMessageBox.warning(self, "Печать", f"Не удалось отправить на печать:\n{e}")

    def repeat_last_print(self):
        if not self.last_tspl:
            return
        try:
            printer_name = win32print.GetDefaultPrinter()
            print_raw(printer_name, self._apply_copies_to_tspl(self.last_tspl, self._get_copies()))
        except Exception as e:
            QMessageBox.warning(self, "Повтор", f"Не удалось повторить печать:\n{e}")


def main():

    app = QApplication(sys.argv)
    w = MirlisMarkApp()

    # старт сразу в развёрнутом окне
    w.showMaximized()

    sys.exit(app.exec())

if __name__ == "__main__":
    main()