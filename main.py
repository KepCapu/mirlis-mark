# main.py
# Mirlis Mark — Система маркировки
# UI: "как на картинке" + редактор предпросмотра + ВИДИМЫЕ стрелочки в выпадающих списках
#
# ВАЖНО:
# - excel_loader.py / label_logic.py / printer.py НЕ ТРОГАЕМ
# - логотип берём по пути: D:\mirlis_mark\Mirlis software logo.png
#
# === PYQT5-СОВМЕСТИМАЯ ВЕРСИЯ ===
# Изменения относительно PyQt6:
# - Все импорты из PyQt5
# - Enum-стиль без вложенных флагов (Qt.AlignCenter вместо Qt.AlignmentFlag.AlignCenter)
# - QMediaPlayer / QVideoWidget из PyQt5.QtMultimedia / QtMultimediaWidgets
# - InsertPolicy, ScrollBarPolicy и т.п. — без вложенных enum-классов


import sys
import os
import shutil
import time
import math
import json
import calendar
from datetime import datetime

from PyQt5.QtWidgets import (
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
    QDateTimeEdit,
    QDateEdit,
    QTimeEdit,
    QCalendarWidget,
    QGridLayout,
    QDialog,
)
from PyQt5.QtCore import QTimer, Qt, QUrl, QSize, QDateTime, QDate, QTime, pyqtSignal, QPoint, QLocale, QEvent, QSizeF
from PyQt5.QtCore import QStringListModel
from PyQt5.QtGui import (
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

# PyQt5: мультимедиа
from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent
from PyQt5.QtMultimediaWidgets import QVideoWidget

from excel_loader import load_products, load_staff
from label_logic import build_label, format_dt
from printer import print_text_as_bitmap_tspl, print_raw
from openpyxl import load_workbook as _peek_workbook
import win32print


# -------------------- ПУТИ: ресурсы и пользовательские данные --------------------
def resource_path(relative_path: str) -> str:
    """Путь к встроенному ресурсу: из исходников — от корня проекта, из exe — из sys._MEIPASS."""
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, *relative_path.replace("/", os.sep).split(os.sep))


def app_data_dir() -> str:
    """Папка приложения в профиле пользователя Windows (%LOCALAPPDATA%\\MirlisMark). Создаётся при первом вызове."""
    if sys.platform == "win32":
        root = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
    else:
        root = os.path.expanduser("~")
    path = os.path.join(root, "MirlisMark")
    os.makedirs(path, exist_ok=True)
    return path


def get_config_path() -> str:
    """Путь к settings.json в пользовательской папке приложения."""
    return os.path.join(app_data_dir(), "settings.json")


def ensure_products_file() -> str:
    """
    Путь к products.xlsx для работы приложения.
    Файл лежит в пользовательской папке; при отсутствии копируется шаблон из ресурсов сборки.
    При отсутствии шаблона или ошибке копирования выбрасывает RuntimeError с понятным текстом.
    """
    app_dir = app_data_dir()
    target = os.path.join(app_dir, "products.xlsx")
    if os.path.isfile(target):
        return target
    template = resource_path("data_sources/products.xlsx")
    if not os.path.isfile(template):
        raise RuntimeError(
            "Встроенный шаблон products.xlsx не найден. "
            "Убедитесь, что при сборке добавлена папка data_sources (--add-data \"data_sources;data_sources\")."
        )
    try:
        shutil.copy2(template, target)
    except Exception as e:
        raise RuntimeError(
            f"Не удалось создать файл products.xlsx в папке приложения:\n{app_dir}\nОшибка: {e}"
        ) from e
    return target


# Константы листов Excel (не пути)
SHEET_PRODUCTS = "продукт"
SHEET_MADE = "изготовил"
SHEET_CHECKED = "цех"

# Ресурсы интерфейса (через resource_path для работы из исходников и из exe)
LOGO_PATH = resource_path("assets/logo.png")
SPLASH_VIDEO_PATH = resource_path("assets/loadingscreen.mp4")
HELP_IMG_PRODUCT = resource_path("assets/help_sheet_product.png")
HELP_IMG_MADE = resource_path("assets/help_sheet_made.png")
HELP_IMG_WORKSHOP = resource_path("assets/help_sheet_workshop.png")
HELP_IMG_TABS = resource_path("assets/help_sheet_tabs.png")

APP_TITLE = "Mirlis Mark — Система маркировки"
APP_MARK = "Mark"
APP_VERSION = "1.0"
APP_SUBTITLE = "Система маркировки"

# фон предпросмотра в ручном режиме: мягкий светло-жёлтый (чуть насыщеннее бледного)
MANUAL_PREVIEW_BG = "#fef3c7"


# -------------------- HELPERS --------------------
def _load_settings() -> dict:
    """Загрузить настройки из settings.json (в папке приложения пользователя)."""
    try:
        path = get_config_path()
        if os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        sys.stderr.write(f"[MirlisMark] Не удалось загрузить настройки: {e}\n")
    except Exception as e:
        sys.stderr.write(f"[MirlisMark] Ошибка при чтении settings.json: {e}\n")
    return {}


def _save_settings(data: dict):
    """Сохранить настройки в settings.json (в папке приложения пользователя)."""
    try:
        with open(get_config_path(), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except OSError as e:
        sys.stderr.write(f"[MirlisMark] Не удалось сохранить настройки: {e}\n")
    except Exception as e:
        sys.stderr.write(f"[MirlisMark] Ошибка при записи settings.json: {e}\n")


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
        self.setFrameShape(QFrame.NoFrame)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)


class Pill(QLabel):
    """Плашка-'pill' (серый бэйдж)."""

    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self.setObjectName("Pill")
        self.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)


class HeaderLabel(QLabel):
    """Заголовок секции по центру с короткой тенью-плашкой."""

    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self.setObjectName("SectionTitle")
        self.setAlignment(Qt.AlignCenter)


class PreviewHeaderLabel(HeaderLabel):
    """Заголовок «Предпросмотр»: двойной клик переключает ручной режим. Без hover/кнопки."""

    doubleClicked = pyqtSignal()

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.doubleClicked.emit()
        super().mouseDoubleClickEvent(event)


class ComboBoxFixedArrow(QComboBox):
    """
    Комбобокс с гарантированно видимой стрелкой.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("ComboWithArrow")


class ComboBoxPopupDown(ComboBoxFixedArrow):
    """
    Комбобокс, у которого выпадающий список всегда раскрывается вниз.
    """

    def showPopup(self):
        super().showPopup()
        view = self.view()
        if view and view.window():
            popup = view.window()
            bottom_left = self.mapToGlobal(self.rect().bottomLeft())
            popup.move(bottom_left.x(), bottom_left.y())


class ToolBtn(QToolButton):
    def __init__(self, text="", parent=None):
        super().__init__(parent)
        self.setText(text)
        self.setCursor(Qt.PointingHandCursor)
        self.setObjectName("ToolBtn")
        self.setCheckable(True)
        self.setAutoRaise(False)


class ActionBtn(QPushButton):
    def __init__(self, text="", kind="default", parent=None):
        super().__init__(text, parent)
        self.setCursor(Qt.PointingHandCursor)
        self.setObjectName(f"Btn_{kind}")


# -------------------- SPLASH VIDEO --------------------
class SplashVideo(QWidget):
    """Окно загрузки с воспроизведением видео перед запуском основного приложения."""

    def __init__(self, video_path: str, on_finished_callback=None, parent=None):
        super().__init__(parent)
        self._on_finished = on_finished_callback
        self._finished_called = False
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setStyleSheet("background-color: #000000;")
        self.setAttribute(Qt.WA_TranslucentBackground, False)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self._video_widget = QVideoWidget()
        self._video_widget.setStyleSheet("background-color: #000000;")
        self._video_widget.setAspectRatioMode(Qt.KeepAspectRatioByExpanding)
        layout.addWidget(self._video_widget)

        self._player = QMediaPlayer()
        self._player.setVideoOutput(self._video_widget)
        self._player.setMedia(QMediaContent(QUrl.fromLocalFile(video_path)))
        self._player.mediaStatusChanged.connect(self._on_media_status_changed)
        self._player.error.connect(self._on_error)

        # Таймаут-подстраховка: если видео не завершилось за 15 сек — запускаем приложение
        self._timeout = QTimer(self)
        self._timeout.setSingleShot(True)
        self._timeout.timeout.connect(self._finish)
        self._timeout.start(15000)

    def _finish(self):
        """Единоразовый запуск основного окна + закрытие splash."""
        if self._finished_called:
            return
        self._finished_called = True
        self._timeout.stop()
        self._player.stop()
        if callable(self._on_finished):
            self._on_finished()
        self.close()

    def _on_media_status_changed(self, status):
        if status == QMediaPlayer.EndOfMedia:
            self._finish()
        elif status == QMediaPlayer.InvalidMedia:
            self._finish()

    def _on_error(self):
        self._finish()

    def play(self):
        self._player.play()

    def showEvent(self, event):
        super().showEvent(event)
        self._player.play()


# -------------------- CUSTOM DATE-TIME PICKER --------------------
_MONTH_NAMES_RU = [
    "", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
]
_MONTH_NAMES_RU_GENITIVE = [
    "", "января", "февраля", "марта", "апреля", "мая", "июня",
    "июля", "августа", "сентября", "октября", "ноября", "декабря",
]
_WEEKDAY_HEADERS = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]


class CustomDateTimePicker(QWidget):
    """Кастомный picker даты и времени с popup-панелью (календарь + время)."""

    dateTimeChanged = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._date = QDate.currentDate()
        self._time = QTime.currentTime()
        self._popup_visible = False

        # --- триггер как у QComboBox ---
        combo_btn_path = resource_path("assets/combo-btn.svg").replace("\\", "/")
        self._trigger_frame = QFrame()
        self._trigger_frame.setObjectName("DateTimeTrigger")
        self._trigger_frame.setMinimumHeight(40)
        self._trigger_frame.setCursor(Qt.PointingHandCursor)
        self._trigger_frame.setStyleSheet(
            "#DateTimeTrigger { background: #ffffff; border: 1px solid #cfd6e0; border-radius: 12px; }"
            "#DateTimeTrigger:hover { border-color: #94a3b8; }"
        )
        self._trigger_frame.installEventFilter(self)
        inner = QHBoxLayout(self._trigger_frame)
        inner.setContentsMargins(14, 0, 4, 0)
        inner.setSpacing(0)
        self._text_label = QLabel()
        self._text_label.setStyleSheet(
            "background: transparent; border: none; font-size: 14px; color: #111827;"
        )
        self._text_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        inner.addWidget(self._text_label, 1)
        self._drop_icon = QLabel()
        self._drop_icon.setFixedSize(36, 36)
        self._drop_icon.setStyleSheet("background: transparent; border: none;")
        try:
            icon = QIcon(combo_btn_path)
            if not icon.isNull():
                self._drop_icon.setPixmap(icon.pixmap(QSize(36, 36)))
        except Exception:
            pass
        inner.addWidget(self._drop_icon, 0, Qt.AlignRight | Qt.AlignVCenter)
        self._text_label.installEventFilter(self)
        self._drop_icon.installEventFilter(self)

        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(self._trigger_frame)

        # --- popup ---
        self._popup = QFrame(self, Qt.Popup)
        self._popup.setStyleSheet(
            "QFrame { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 16px; }"
        )
        self._popup.setFixedWidth(625)
        self._build_popup()
        self._update_btn_text()

    # ---------- btn text ----------
    def _update_btn_text(self):
        d = self._date
        month_gen = _MONTH_NAMES_RU_GENITIVE[d.month()]
        txt = f"  📅   {d.day()} {month_gen} {d.year()}  {self._time.toString('HH:mm')}"
        self._text_label.setText(txt)

    # ---------- popup build ----------
    def _build_popup(self):
        popup_lay = QVBoxLayout(self._popup)
        popup_lay.setContentsMargins(16, 10, 16, 10)
        popup_lay.setSpacing(8)

        # --- навигация месяца ---
        nav = QHBoxLayout()
        nav.setSpacing(0)
        self._prev_btn = QPushButton("‹")
        self._prev_btn.setFixedSize(36, 36)
        self._prev_btn.setCursor(Qt.PointingHandCursor)
        self._prev_btn.setStyleSheet(
            "QPushButton { background: transparent; border: none; font-size: 22px; font-weight: 700; color: #64748b; }"
            "QPushButton:hover { color: #111827; }"
        )
        self._prev_btn.clicked.connect(lambda: self._change_month(-1))

        self._month_label = QLabel()
        self._month_label.setAlignment(Qt.AlignCenter)
        self._month_label.setStyleSheet(
            "font-size: 16px; font-weight: 700; color: #111827; background: transparent;"
        )

        self._next_btn = QPushButton("›")
        self._next_btn.setFixedSize(36, 36)
        self._next_btn.setCursor(Qt.PointingHandCursor)
        self._next_btn.setStyleSheet(
            "QPushButton { background: transparent; border: none; font-size: 22px; font-weight: 700; color: #64748b; }"
            "QPushButton:hover { color: #111827; }"
        )
        self._next_btn.clicked.connect(lambda: self._change_month(1))

        nav.addWidget(self._prev_btn)
        nav.addWidget(self._month_label, 1)
        nav.addWidget(self._next_btn)
        popup_lay.addLayout(nav)

        # --- разделитель ---
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("color: #e2e8f0; background: #e2e8f0; border: none; max-height: 1px;")
        popup_lay.addWidget(sep)

        # --- сетка дней ---
        self._cal_grid = QGridLayout()
        self._cal_grid.setSpacing(2)
        for c in range(7):
            self._cal_grid.setColumnMinimumWidth(c, 56)
        # заголовки
        for col, name in enumerate(_WEEKDAY_HEADERS):
            lbl = QLabel(name)
            lbl.setAlignment(Qt.AlignCenter)
            color = "#ef6c00" if col >= 5 else "#9ca3af"
            lbl.setStyleSheet(
                f"font-size: 12px; font-weight: 600; color: {color}; "
                "background: transparent; padding: 4px 0;"
            )
            self._cal_grid.addWidget(lbl, 0, col)
        popup_lay.addLayout(self._cal_grid)

        # --- разделитель ---
        sep2 = QFrame()
        sep2.setFrameShape(QFrame.HLine)
        sep2.setStyleSheet("color: #e2e8f0; background: #e2e8f0; border: none; max-height: 1px;")
        popup_lay.addWidget(sep2)

        # --- блок времени ---
        time_header = QLabel("Время")
        time_header.setAlignment(Qt.AlignCenter)
        time_header.setStyleSheet(
            "font-size: 21px; font-weight: 700; color: #111827; background: transparent; padding: 2px 0;"
        )
        time_header_row = QHBoxLayout()
        time_header_row.setContentsMargins(0, 0, 0, 0)
        time_header_row.addStretch(1)
        time_header_row.addWidget(time_header)
        time_header_row.addStretch(1)
        popup_lay.addLayout(time_header_row)

        arrow_style = (
            "QPushButton { background: #f9b233; border: none; border-radius: 10px; "
            "color: #ffffff; font-size: 16px; font-weight: 800; min-width: 34px; min-height: 28px; }"
            "QPushButton:hover { background: #e5a020; }"
            "QPushButton:pressed { background: #cc8c10; }"
        )
        btn_h_up = QPushButton("▲")
        btn_h_up.setStyleSheet(arrow_style)
        btn_h_up.setCursor(Qt.PointingHandCursor)
        btn_h_up.setAutoRepeat(True)
        btn_h_up.setAutoRepeatDelay(400)
        btn_h_up.setAutoRepeatInterval(100)
        btn_h_up.clicked.connect(lambda: self._step_time("h", 1))
        btn_h_down = QPushButton("▼")
        btn_h_down.setStyleSheet(arrow_style)
        btn_h_down.setCursor(Qt.PointingHandCursor)
        btn_h_down.setAutoRepeat(True)
        btn_h_down.setAutoRepeatDelay(400)
        btn_h_down.setAutoRepeatInterval(100)
        btn_h_down.clicked.connect(lambda: self._step_time("h", -1))
        btn_m_up = QPushButton("▲")
        btn_m_up.setStyleSheet(arrow_style)
        btn_m_up.setCursor(Qt.PointingHandCursor)
        btn_m_up.setAutoRepeat(True)
        btn_m_up.setAutoRepeatDelay(400)
        btn_m_up.setAutoRepeatInterval(60)
        btn_m_up.clicked.connect(lambda: self._step_time("m", 1))
        btn_m_down = QPushButton("▼")
        btn_m_down.setStyleSheet(arrow_style)
        btn_m_down.setCursor(Qt.PointingHandCursor)
        btn_m_down.setAutoRepeat(True)
        btn_m_down.setAutoRepeatDelay(400)
        btn_m_down.setAutoRepeatInterval(60)
        btn_m_down.clicked.connect(lambda: self._step_time("m", -1))

        self._hour_label = QLabel()
        self._hour_label.setAlignment(Qt.AlignCenter)
        self._hour_label.setMinimumSize(72, 40)
        self._hour_label.setStyleSheet(
            "font-size: 28px; font-weight: 700; color: #111827; "
            "background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px;"
        )
        self._min_label = QLabel()
        self._min_label.setAlignment(Qt.AlignCenter)
        self._min_label.setMinimumSize(72, 40)
        self._min_label.setStyleSheet(
            "font-size: 28px; font-weight: 700; color: #111827; "
            "background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px;"
        )
        colon = QLabel(":")
        colon.setAlignment(Qt.AlignCenter)
        colon.setFixedSize(24, 40)
        colon.setStyleSheet("font-size: 28px; font-weight: 700; color: #111827; background: transparent;")

        time_grid = QGridLayout()
        time_grid.setSpacing(4)
        time_grid.setContentsMargins(0, 2, 0, 0)
        time_grid.addWidget(self._hour_label, 0, 0, 1, 2, Qt.AlignCenter)
        time_grid.addWidget(colon, 0, 2, Qt.AlignCenter)
        time_grid.addWidget(self._min_label, 0, 3, 1, 2, Qt.AlignCenter)
        time_grid.addWidget(btn_h_up, 1, 0, Qt.AlignHCenter)
        time_grid.addWidget(btn_h_down, 1, 1, Qt.AlignHCenter)
        time_grid.addWidget(btn_m_up, 1, 3, Qt.AlignHCenter)
        time_grid.addWidget(btn_m_down, 1, 4, Qt.AlignHCenter)
        time_grid.setColumnMinimumWidth(0, 40)
        time_grid.setColumnMinimumWidth(1, 40)
        time_grid.setColumnMinimumWidth(2, 24)
        time_grid.setColumnMinimumWidth(3, 40)
        time_grid.setColumnMinimumWidth(4, 40)
        popup_lay.addLayout(time_grid)

        self._refresh_calendar()
        self._refresh_time_labels()

    # ---------- calendar grid ----------
    def _refresh_calendar(self):
        for r in range(7, 0, -1):
            for c in range(7):
                item = self._cal_grid.itemAtPosition(r, c)
                if item and item.widget():
                    item.widget().deleteLater()

        self._month_label.setText(f"{_MONTH_NAMES_RU[self._date.month()]} {self._date.year()}")

        year = self._date.year()
        month = self._date.month()
        cal = calendar.Calendar(firstweekday=0)
        weeks = cal.monthdayscalendar(year, month)

        for row_idx, week in enumerate(weeks):
            for col_idx, day in enumerate(week):
                btn = QPushButton("" if day == 0 else str(day))
                btn.setFixedSize(56, 36)
                btn.setCursor(Qt.PointingHandCursor if day else Qt.ArrowCursor)

                if day == 0:
                    btn.setEnabled(False)
                    btn.setStyleSheet(
                        "QPushButton { background: transparent; border: none; }"
                    )
                elif day == self._date.day():
                    btn.setStyleSheet(
                        "QPushButton { background: #f9b233; color: #ffffff; border: none; "
                        "border-radius: 10px; font-size: 14px; font-weight: 700; }"
                        "QPushButton:hover { background: #e5a020; }"
                    )
                elif col_idx >= 5:
                    btn.setStyleSheet(
                        "QPushButton { background: transparent; border: none; border-radius: 10px; "
                        "font-size: 14px; color: #ef6c00; }"
                        "QPushButton:hover { background: #fff3e0; }"
                    )
                else:
                    btn.setStyleSheet(
                        "QPushButton { background: transparent; border: none; border-radius: 10px; "
                        "font-size: 14px; color: #374151; }"
                        "QPushButton:hover { background: #f1f5f9; }"
                    )

                if day > 0:
                    btn.clicked.connect(lambda checked, d=day: self._select_day(d))

                self._cal_grid.addWidget(btn, row_idx + 1, col_idx)

    def _change_month(self, delta):
        m = self._date.month() + delta
        y = self._date.year()
        if m < 1:
            m = 12
            y -= 1
        elif m > 12:
            m = 1
            y += 1
        max_day = calendar.monthrange(y, m)[1]
        d = min(self._date.day(), max_day)
        self._date = QDate(y, m, d)
        self._refresh_calendar()
        self._update_btn_text()
        self.dateTimeChanged.emit()

    def _select_day(self, day):
        self._date = QDate(self._date.year(), self._date.month(), day)
        self._refresh_calendar()
        self._update_btn_text()
        self.dateTimeChanged.emit()

    # ---------- time ----------
    def _refresh_time_labels(self):
        self._hour_label.setText(f"{self._time.hour():02d}")
        self._min_label.setText(f"{self._time.minute():02d}")

    def _step_time(self, part, delta):
        h = self._time.hour()
        m = self._time.minute()
        if part == "h":
            h = (h + delta) % 24
        else:
            m = (m + delta) % 60
        self._time = QTime(h, m)
        self._refresh_time_labels()
        self._update_btn_text()
        self.dateTimeChanged.emit()

    # ---------- event filter ----------
    def eventFilter(self, obj, event):
        if event.type() == QEvent.MouseButtonPress and obj in (
            self._trigger_frame,
            self._text_label,
            self._drop_icon,
        ):
            self._toggle_popup()
            return True
        return super().eventFilter(obj, event)

    # ---------- popup toggle ----------
    def _toggle_popup(self):
        if self._popup.isVisible():
            self._popup.hide()
        else:
            self._popup.adjustSize()
            popup_h = self._popup.height() or self._popup.sizeHint().height() or 320
            win = self.window()
            win_rect = win.frameGeometry()
            f = self._trigger_frame
            btn_bottom_global = f.mapToGlobal(QPoint(0, f.height())).y()
            btn_top_global = f.mapToGlobal(QPoint(0, 0)).y()
            space_below = win_rect.bottom() - btn_bottom_global
            margin = 8
            popup_below_gap = 14

            if space_below - margin >= popup_h:
                pos = f.mapToGlobal(QPoint(0, f.height() + popup_below_gap))
            else:
                pos = f.mapToGlobal(QPoint(0, 0))
                pos.setY(btn_top_global - popup_h - margin)
            self._popup.move(pos)
            self._popup.show()

    # ---------- public API ----------
    def date(self) -> QDate:
        return self._date

    def time_(self) -> QTime:
        return self._time

    def setDate(self, d: QDate):
        self._date = d
        self._refresh_calendar()
        self._update_btn_text()

    def setTime(self, t: QTime):
        self._time = t
        self._refresh_time_labels()
        self._update_btn_text()


# -------------------- MAIN APP --------------------
class MirlisMarkApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon(resource_path("assets/mark_app.ico")))
        self.setWindowTitle(APP_TITLE)

        self.setMinimumSize(1100, 650)

        self.last_printed_preview_text = None
        self._last_printed_tspl_bytes = None
        self.last_history_entry = None

        self.label_w_mm = 58.0
        self.label_h_mm = 80.0

        self.products = []
        self.staff_made = []
        self.staff_checked = []
        self.loaded_at_str = "—"

        settings = _load_settings()
        try:
            default_excel = ensure_products_file()
        except RuntimeError as e:
            QMessageBox.critical(
                self,
                "Ошибка инициализации",
                str(e),
            )
            default_excel = os.path.join(app_data_dir(), "products.xlsx")
        self.excel_path = settings.get("excel_path", default_excel)

        self._updating_preview = False
        self._user_edited_preview = False
        self._preview_manual_mode = False
        self._base_font_size = 20
        self._preview_scale = 1.0

        self.history_entries = []
        self._history_filter_text = ""
        self.history_page = 0
        self.history_page_size = 6
        self._loading_from_history = False
        self._selected_history_id = None

        self._apply_global_style()
        self.init_ui()
        self.reload_excel(show_message=False)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.refresh_preview)
        self.timer.start(5000)

        # автообновление Excel раз в минуту (без всплывающих окон при фоновых ошибках)
        self._excel_autorefresh_timer = QTimer(self)
        self._excel_autorefresh_timer.timeout.connect(self._on_excel_autorefresh_tick)
        self._excel_autorefresh_timer.start(60_000)

        # архив напечатанных этикеток: очистка старше 31 дня при старте
        self._cleanup_old_label_archives(days=31)

        self.refresh_preview()

    # ---------------- STYLE ----------------
    def _apply_global_style(self):
        combo_btn_path = resource_path("assets/combo-btn.svg").replace("\\", "/")

        _qss = """
            QWidget {
                background: #f6f7f9;
                font-family: "Segoe UI";
                color: #111827;
            }

            #TopBar {
                background: #ffffff;
                border: 1px solid #e5e7eb;
                border-radius: 18px;
            }

            #Card {
                background: #ffffff;
                border: 1px solid #e5e7eb;
                border-radius: 18px;
            }

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

            #SectionTitle {
                background: #eef2f6;
                border-radius: 14px;
                padding: 10px 22px;
                font-size: 22px;
                font-weight: 800;
                letter-spacing: 0.2px;
            }

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

            #ExcelPill {
                background: #eef2f6;
                border-radius: 18px;
                padding: 14px 22px;
                font-size: 14px;
                font-weight: 600;
                color: #0f172a;
            }

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
        _qss = _qss.replace("url(assets/combo-btn.svg)", f"url({combo_btn_path})")
        self.setStyleSheet(_qss)

    # ---------------- Auto Excel refresh ----------------
    def _on_excel_autorefresh_tick(self):
        """Фоновое обновление Excel раз в минуту без MessageBox."""
        try:
            self.reload_excel(show_message=False, silent_errors=True)
        except TypeError:
            # на случай если сигнатура reload_excel не содержит silent_errors
            self.reload_excel(show_message=False)

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

        self.logo = QLabel()
        self.logo.setStyleSheet("background: transparent;")
        self.logo.setScaledContents(False)
        self._load_logo()
        top_layout.addWidget(self.logo, 0, Qt.AlignVCenter)

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
        title_row.addWidget(self.badge_ver, 0, Qt.AlignVCenter)

        title_row.addStretch(1)
        title_block.addLayout(title_row)

        self.subtitle = QLabel(APP_SUBTITLE)
        self.subtitle.setStyleSheet("font-size: 16px; color: #64748b; padding-left: 2px; margin-top: 0px; background: transparent;")
        title_block.addWidget(self.subtitle)

        top_layout.addLayout(title_block, 0)

        self.excel_pill = QLabel("Excel: —")
        self.excel_pill.setObjectName("ExcelPill")
        self.excel_pill.setAlignment(Qt.AlignCenter)
        self.excel_pill.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.excel_pill.setMinimumWidth(360)
        self.excel_pill.setMaximumWidth(520)
        self.excel_pill.setMinimumHeight(48)

        top_layout.addStretch(1)
        top_layout.addWidget(self.excel_pill, 0, Qt.AlignVCenter)
        top_layout.addStretch(1)

        self.reload_btn = ActionBtn("Обновить", kind="default")
        self.reload_btn.clicked.connect(self.reload_excel)
        top_layout.addWidget(self.reload_btn, 0, Qt.AlignVCenter)

        self.open_folder_btn = ActionBtn("Папка", kind="default")
        self.open_folder_btn.clicked.connect(self.open_excel_folder)
        top_layout.addWidget(self.open_folder_btn, 0, Qt.AlignVCenter)

        self.choose_path_btn = ActionBtn("Выбрать файл", kind="default")
        self.choose_path_btn.clicked.connect(self.choose_excel_path)
        top_layout.addWidget(self.choose_path_btn, 0, Qt.AlignVCenter)

        self.clear_btn = ActionBtn("Очистить", kind="danger")
        self.clear_btn.clicked.connect(self.clear_fields)
        top_layout.addWidget(self.clear_btn, 0, Qt.AlignVCenter)

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
        left_layout.addWidget(left_title, 0, Qt.AlignHCenter)

        lab_prod = QLabel("Продукт")
        lab_prod.setObjectName("FieldLabel")
        lab_prod.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        left_layout.addWidget(lab_prod)

        self.product_combo = ComboBoxFixedArrow()
        self.product_combo.setEditable(True)
        self.product_combo.setInsertPolicy(QComboBox.NoInsert)
        self.product_combo.setPlaceholderText("Введите продукт или выберите из списка")
        self.product_combo.setMaxVisibleItems(8)
        left_layout.addWidget(self.product_combo)

        self.product_model = QStringListModel([])
        self.product_completer = QCompleter(self.product_model, self)
        self.product_completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.product_completer.setFilterMode(Qt.MatchContains)
        self.product_completer.setCompletionMode(QCompleter.PopupCompletion)
        self.product_combo.setCompleter(self.product_completer)

        grid = QHBoxLayout()
        grid.setSpacing(12)

        col_units = QVBoxLayout()
        lab_units = QLabel("Ед. изм.")
        lab_units.setObjectName("FieldLabel")
        lab_units.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        col_units.addWidget(lab_units)

        self.unit_combo = ComboBoxFixedArrow()
        self.unit_combo.addItem("")
        col_units.addWidget(self.unit_combo)
        grid.addLayout(col_units, 1)

        col_qty = QVBoxLayout()
        qty_row = QHBoxLayout()
        qty_row.setSpacing(10)

        self.minus_btn = ActionBtn("−", kind="default")
        self.minus_btn.setFixedWidth(60)
        self.minus_btn.setAutoRepeat(True)
        self.minus_btn.setAutoRepeatDelay(400)
        self.minus_btn.setAutoRepeatInterval(80)

        self.qty_input = QLineEdit()
        self.qty_input.setPlaceholderText("Введите количество")

        self.plus_btn = ActionBtn("+", kind="default")
        self.plus_btn.setFixedWidth(60)
        self.plus_btn.setAutoRepeat(True)
        self.plus_btn.setAutoRepeatDelay(400)
        self.plus_btn.setAutoRepeatInterval(80)

        qty_row.addWidget(self.minus_btn)
        qty_row.addWidget(self.qty_input, 1)
        qty_row.addWidget(self.plus_btn)
        col_qty.addLayout(qty_row)

        grid.addLayout(col_qty, 2)
        left_layout.addLayout(grid)

        lab_made = QLabel("Изготовил")
        lab_made.setObjectName("FieldLabel")
        lab_made.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        left_layout.addWidget(lab_made)

        self.made_combo = ComboBoxFixedArrow()
        self.made_combo.addItem("— не выбрано —")
        self.made_combo.setMaxVisibleItems(8)
        left_layout.addWidget(self.made_combo)

        self.made_manual = QCheckBox("Ручной ввод")
        left_layout.addWidget(self.made_manual)

        self.made_input = QLineEdit()
        self.made_input.setPlaceholderText("ФИО (можно оставить пустым)")
        self.made_input.setVisible(False)
        left_layout.addWidget(self.made_input)

        lab_chk = QLabel("Цех")
        lab_chk.setObjectName("FieldLabel")
        lab_chk.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        left_layout.addWidget(lab_chk)

        self.checked_combo = ComboBoxPopupDown()
        self.checked_combo.addItem("— не выбрано —")
        self.checked_combo.setMaxVisibleItems(8)
        left_layout.addWidget(self.checked_combo)

        self.checked_manual = QCheckBox("Ручной ввод")
        left_layout.addWidget(self.checked_manual)

        self.checked_input = QLineEdit()
        self.checked_input.setPlaceholderText("Цех (можно оставить пустым)")
        self.checked_input.setVisible(False)
        left_layout.addWidget(self.checked_input)

        lab_dt = QLabel("Дата и время")
        lab_dt.setObjectName("FieldLabel")
        lab_dt.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.manual_datetime_label = lab_dt

        self.manual_datetime_picker = CustomDateTimePicker()

        self.manual_datetime_label.setVisible(False)
        self.manual_datetime_picker.setVisible(False)
        left_layout.addWidget(self.manual_datetime_label)
        left_layout.addWidget(self.manual_datetime_picker)

        left_layout.addStretch(1)

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

        self.preview_header = PreviewHeaderLabel("Предпросмотр")
        right_layout.addWidget(self.preview_header, 0, Qt.AlignHCenter)

        self.manual_mode_subtitle = QLabel("редактирование")
        self.manual_mode_subtitle.setAlignment(Qt.AlignCenter)
        self.manual_mode_subtitle.setStyleSheet(
            "font-size: 12px; font-weight: 700; color: #ffffff; "
            "background: #8b5cf6; border-radius: 8px; padding: 4px 10px; margin: 0;"
        )
        self.manual_mode_subtitle.setVisible(False)
        right_layout.addWidget(self.manual_mode_subtitle, 0, Qt.AlignHCenter)

        # toolbar
        tb = QHBoxLayout()
        tb.setSpacing(10)

        self.btn_font_minus = ActionBtn("A-", kind="default")
        self.btn_font_minus.setFixedWidth(60)
        self.btn_font_plus = ActionBtn("A+", kind="default")
        self.btn_font_plus.setFixedWidth(60)

        self.font_size_combo = ComboBoxFixedArrow()
        self.font_size_combo.setEditable(True)
        self.font_size_combo.setFixedWidth(90)
        self.font_size_combo.addItems([str(s) for s in [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]])
        self.font_size_combo.setCurrentText(str(self._base_font_size))

        self.btn_bold = ToolBtn("Ж")
        self.btn_bold.setFont(QFont("Segoe UI", 11, QFont.Black))
        self.btn_italic = ToolBtn("К")
        f_it = QFont("Segoe UI", 11, QFont.Black)
        f_it.setItalic(True)
        self.btn_italic.setFont(f_it)
        self.btn_underline = ToolBtn("Ч")
        f_un = QFont("Segoe UI", 11, QFont.Black)
        f_un.setUnderline(True)
        self.btn_underline.setFont(f_un)

        self.btn_align_left = ToolBtn("")
        self.btn_align_center = ToolBtn("")
        self.btn_align_right = ToolBtn("")

        for btn, mode in (
            (self.btn_align_left, "left"),
            (self.btn_align_center, "center"),
            (self.btn_align_right, "right"),
        ):
            btn.setIcon(self._make_align_icon(mode))
            btn.setIconSize(QSize(26, 18))
            btn.setFixedWidth(62)

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
        tb.addStretch(1)
        tb.addWidget(self.font_combo)

        right_layout.addLayout(tb)

        # preview editor
        self.preview = QTextEdit()
        self.preview.setObjectName("PreviewEditor")
        self.preview.setAcceptRichText(True)
        self.preview.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.preview.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        self.preview.setFixedSize(450, 600)

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
        self.preview.setReadOnly(True)

        self.preview_wrap = QFrame()
        self.preview_wrap.setObjectName("LabelWrap")

        wrap_lay = QGridLayout(self.preview_wrap)
        wrap_lay.setContentsMargins(12, 12, 12, 12)
        wrap_lay.setHorizontalSpacing(0)
        wrap_lay.setVerticalSpacing(0)

        # Этикетка строго по центру области предпросмотра
        wrap_lay.addWidget(self.preview, 0, 0, 1, 3, Qt.AlignHCenter | Qt.AlignTop)

        # Кнопка справа, не влияет на центрирование этикетки

        right_layout.addWidget(self.preview_wrap, 1)

        center_panel = QWidget()
        center_panel_layout = QVBoxLayout(center_panel)
        center_panel_layout.setContentsMargins(0, 0, 0, 0)
        center_panel_layout.setSpacing(10)
        center_panel_layout.addWidget(self.card_right)

        label_size_row = QHBoxLayout()
        label_size_row.setSpacing(8)
        label_size_lab = QLabel("Размер этикетки")
        label_size_lab.setObjectName("FieldLabel")
        self.label_size_combo = ComboBoxFixedArrow()
        self.label_size_combo.addItem("58×80 мм")
        self.label_size_combo.addItem("58×60 мм")
        self.label_size_combo.addItem("70×70 мм")
        self.label_size_combo.addItem("Цветные")
        self.label_size_combo.setCurrentIndex(0)
        self.label_size_combo.setMinimumWidth(140)
        label_size_row.addWidget(label_size_lab)
        label_size_row.addWidget(self.label_size_combo, 1)
        center_panel_layout.addLayout(label_size_row)

        # print row
        pr = QHBoxLayout()
        pr.setSpacing(12)

        self.print_btn = ActionBtn("ПЕЧАТЬ", kind="primary")
        self.repeat_btn = ActionBtn("Повторить", kind="secondary")
        self.repeat_btn.setEnabled(False)

        self.copies_btn = ActionBtn("Количество", kind="secondary")
        self.copies_minus = ActionBtn("−", kind="default")
        self.copies_minus.setFixedWidth(44)
        self.copies_minus.setAutoRepeat(True)
        self.copies_minus.setAutoRepeatDelay(400)
        self.copies_minus.setAutoRepeatInterval(80)
        self.copies_input = QLineEdit("1")
        self.copies_input.setFixedWidth(60)
        self.copies_input.setAlignment(Qt.AlignCenter)
        self.copies_plus = ActionBtn("+", kind="default")
        self.copies_plus.setFixedWidth(44)
        self.copies_plus.setAutoRepeat(True)
        self.copies_plus.setAutoRepeatDelay(400)
        self.copies_plus.setAutoRepeatInterval(80)

        copies_wrap = QWidget()
        cw = QHBoxLayout(copies_wrap)
        cw.setContentsMargins(0, 0, 0, 0)
        cw.setSpacing(8)
        cw.addWidget(self.copies_btn, 1)
        cw.addWidget(self.copies_minus, 0)
        cw.addWidget(self.copies_input, 0)
        cw.addWidget(self.copies_plus, 0)

        for w in (self.print_btn, self.repeat_btn, self.copies_btn, self.copies_minus, self.copies_plus, self.copies_input):
            w.setMinimumHeight(68)

        for w in (self.minus_btn, self.plus_btn, self.copies_minus, self.copies_plus):
            w.setObjectName("StepperBtn")

        pr.addWidget(self.print_btn, 1)
        pr.addWidget(self.repeat_btn, 1)
        pr.addWidget(copies_wrap, 1)

        center_panel_layout.addLayout(pr)

        # -------- History panel (right) --------
        self.history_panel = QWidget()
        self.history_panel.setObjectName("HistoryPanel")
        history_layout = QVBoxLayout(self.history_panel)
        history_layout.setContentsMargins(18, 18, 18, 18)
        history_layout.setSpacing(12)

        history_title = HeaderLabel("История")
        history_layout.addWidget(history_title, 0, Qt.AlignHCenter)

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

        pager_row = QHBoxLayout()
        pager_row.setSpacing(8)

        self.history_prev_btn = ActionBtn("←", kind="default")
        self.history_next_btn = ActionBtn("→", kind="default")
        self.history_page_label = QLabel("Страница 1 из 1")
        self.history_page_label.setAlignment(Qt.AlignCenter)
        self.history_page_label.setStyleSheet("color: #6b7280; font-size: 12px;")

        pager_row.addWidget(self.history_prev_btn, 0)
        pager_row.addWidget(self.history_page_label, 1)
        pager_row.addWidget(self.history_next_btn, 0)

        history_layout.addLayout(pager_row)

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
        self.label_size_combo.currentIndexChanged.connect(self._on_label_size_changed)

        self.preview.textChanged.connect(self._on_preview_text_changed)
        self.preview.cursorPositionChanged.connect(self._sync_format_toolbar_from_cursor)
        self.preview.selectionChanged.connect(self._sync_format_toolbar_from_cursor)

        self.btn_font_minus.clicked.connect(lambda: self._change_font_size(-1))
        self.btn_font_plus.clicked.connect(lambda: self._change_font_size(+1))
        self.font_size_combo.currentTextChanged.connect(self.on_font_size_combo_changed)

        self.btn_bold.clicked.connect(self._toggle_bold_on_selection)
        self.btn_italic.clicked.connect(self._toggle_italic_on_selection)
        self.btn_underline.clicked.connect(self._toggle_underline_on_selection)

        self.btn_align_left.clicked.connect(lambda: self._set_alignment(Qt.AlignLeft))
        self.btn_align_center.clicked.connect(lambda: self._set_alignment(Qt.AlignHCenter))
        self.btn_align_right.clicked.connect(lambda: self._set_alignment(Qt.AlignRight))

        self.font_combo.currentFontChanged.connect(self._set_font_family_on_selection)
        self.preview_header.doubleClicked.connect(self._toggle_preview_manual_mode)
        self.manual_datetime_picker.dateTimeChanged.connect(self.refresh_preview)

        self.history_search.textChanged.connect(self._on_history_search_text_changed)
        self.history_prev_btn.clicked.connect(lambda: self._change_history_page(-1))
        self.history_next_btn.clicked.connect(lambda: self._change_history_page(+1))

        self.preview.setFont(QFont("Segoe UI", self._base_font_size))

        for cb in (self.product_combo, self.unit_combo, self.made_combo, self.checked_combo, self.font_combo, self.label_size_combo):
            cb.setObjectName("ComboWithArrow")

        self._rebuild_history_view()

    def _make_align_icon(self, mode):
        pix = QPixmap(26, 18)
        pix.fill(Qt.transparent)

        painter = QPainter(pix)
        painter.setRenderHint(QPainter.Antialiasing, False)

        color = QColor("#111827")
        bar_h = 2
        gaps = [2, 8, 14]

        if mode == "left":
            widths = [16, 20, 14]
            xs = [3, 3, 3]
        elif mode == "center":
            widths = [16, 20, 14]
            xs = [(26 - w) // 2 for w in widths]
        elif mode == "right":
            widths = [16, 20, 14]
            xs = [26 - w - 3 for w in widths]
        else:
            widths = [18, 18, 18]
            xs = [4, 4, 4]

        for x, y, w in zip(xs, gaps, widths):
            painter.fillRect(x, y, w, bar_h, color)

        painter.end()
        return QIcon(pix)
    def _load_logo(self):
        if os.path.exists(LOGO_PATH):
            pix = QPixmap(LOGO_PATH)
            if not pix.isNull():
                target_h = 56
                dpr = self.devicePixelRatioF() if hasattr(self, 'devicePixelRatioF') else 2.0
                scaled = pix.scaledToHeight(
                    int(target_h * dpr),
                    Qt.SmoothTransformation,
                )
                scaled.setDevicePixelRatio(dpr)
                self.logo.setPixmap(scaled)
                return
        self.logo.setText("")

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._resize_label_preview()
        if hasattr(self, "history_panel"):
            self.history_panel.setVisible(self.width() >= 1100)

    def _resize_label_preview(self):
        if not hasattr(self, "preview_wrap"):
            return

        rect = self.preview_wrap.contentsRect()
        avail_w = rect.width()
        avail_h = rect.height()
        if avail_w < 50 or avail_h < 50:
            return

        # Общая система масштаба для всех форматов:
        # сравниваем реальные размеры этикеток относительно максимального формата.
        max_w_mm = 70.0
        max_h_mm = 80.0

        margin_w = 28
        margin_h = 28

        px_per_mm = min(
            (avail_w - margin_w) / max_w_mm,
            (avail_h - margin_h) / max_h_mm
        )

        target_w = int(round(self.label_w_mm * px_per_mm))
        target_h = int(round(self.label_h_mm * px_per_mm))

        target_w = max(260, target_w)
        target_h = max(260, target_h)

        # Локальный масштаб только для preview (не влияет на печать).
        # Делаем шрифт чуть меньше на небольших экранах, пропорционально размеру preview.
        preview_ref_w = 450.0
        scale = float(target_w) / preview_ref_w if preview_ref_w > 0 else 1.0
        self._preview_scale = max(0.65, min(1.0, scale))

        self.preview.setFixedSize(target_w, target_h)
    def _on_label_size_changed(self, index):
        sizes = {0: (58.0, 80.0), 1: (58.0, 60.0), 2: (70.0, 70.0), 3: (70.0, 70.0)}
        self.label_w_mm, self.label_h_mm = sizes.get(index, (58.0, 80.0))
        self._resize_label_preview()

        self._user_edited_preview = False
        text, can_print = self._build_label_plain_text()
        self.print_btn.setEnabled(can_print)
        self._set_preview_text_programmatically(text)
        QApplication.processEvents()

    def _toggle_preview_manual_mode(self):
        self._preview_manual_mode = not self._preview_manual_mode
        on = self._preview_manual_mode
        self.manual_datetime_label.setVisible(on)
        self.manual_datetime_picker.setVisible(on)
        self.manual_mode_subtitle.setVisible(on)
        self.preview.setReadOnly(not on)
        self._user_edited_preview = False
        self.refresh_preview()

    def _resolve_sheet_names(self, excel_path, silent: bool = False):
        """
        Проверяет, что в файле есть все 3 обязательных листа.
        Возвращает (sheet_products, sheet_made, sheet_checked) или None если файл не подходит.
        """
        wb = _peek_workbook(excel_path, read_only=True, data_only=True)
        available = wb.sheetnames
        wb.close()

        # проверяем наличие каждого листа
        sheet_products = SHEET_PRODUCTS if SHEET_PRODUCTS in available else ("products" if "products" in available else None)
        sheet_made = SHEET_MADE if SHEET_MADE in available else None
        sheet_checked = SHEET_CHECKED if SHEET_CHECKED in available else None

        missing = []
        if not sheet_products:
            missing.append(f"«{SHEET_PRODUCTS}»")
        if not sheet_made:
            missing.append(f"«{SHEET_MADE}»")
        if not sheet_checked:
            missing.append(f"«{SHEET_CHECKED}»")

        if missing:
            available_str = ", ".join([f"«{s}»" for s in available]) if available else "(пусто)"
            if not silent:
                QMessageBox.warning(
                    self,
                    "Неподходящий файл",
                    f"В выбранном файле отсутствуют обязательные листы:\n"
                    f"{', '.join(missing)}\n\n"
                    f"Листы в файле: {available_str}\n\n"
                    f"Файл Excel должен содержать 3 листа:\n\n"
                    f"1. «{SHEET_PRODUCTS}» — список продукции\n"
                    f"   Заголовки: Код | Наименование | Срок годности (ч) | Ед. измер. | Активен\n\n"
                    f"2. «{SHEET_MADE}» — сотрудники\n"
                    f"   Заголовки: ФИО | Активен\n\n"
                    f"3. «{SHEET_CHECKED}» — цехи\n"
                    f"   Заголовки: Цех | Активен\n\n"
                    f"Первая строка каждого листа — заголовки, данные со второй строки.",
                )
            return None

        return (sheet_products, sheet_made, sheet_checked)

    def reload_excel(self, show_message=True, silent_errors: bool = False):
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

            # определяем имена листов
            resolved = self._resolve_sheet_names(self.excel_path, silent=silent_errors)
            if resolved is None:
                # пользователь отменил выбор листов
                return
            sheet_products, sheet_made, sheet_checked = resolved

            # загружаем продукты
            if sheet_products:
                products_all = load_products(self.excel_path, sheet_name=sheet_products)
                self.products = [p for p in products_all if int(p.get("active", 0)) == 1]
                self.products.sort(key=lambda x: (x.get("name") or "").lower())
            else:
                self.products = []

            # загружаем сотрудников «изготовил»
            if sheet_made:
                self.staff_made = [s for s in load_staff(self.excel_path, sheet_made) if int(s.get("active", 0)) == 1]
            else:
                self.staff_made = []

            # загружаем «цех»
            if sheet_checked:
                self.staff_checked = [s for s in load_staff(self.excel_path, sheet_checked) if int(s.get("active", 0)) == 1]
            else:
                self.staff_checked = []

            self.staff_made = [{"fio": (x.get("fio") or x.get("name") or "").strip(), "active": x.get("active", 1)} for x in self.staff_made]
            self.staff_checked = [{"fio": (x.get("fio") or x.get("name") or "").strip(), "active": x.get("active", 1)} for x in self.staff_checked]

            self.staff_made.sort(key=lambda x: (x.get("fio") or "").lower())
            self.staff_checked.sort(key=lambda x: (x.get("fio") or "").lower())

            self.fill_products(current_product)
            self.fill_staff()

            self.loaded_at_str = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
            self.update_excel_status()

            if current_qty:
                self.qty_input.setText(current_qty)

            self.made_manual.setChecked(made_manual)
            self.checked_manual.setChecked(checked_manual)
            self.made_input.setText(made_text)
            self.checked_input.setText(checked_text)

            if made_combo and not made_combo.startswith("—"):
                idx = self.made_combo.findText(made_combo)
                if idx >= 0:
                    self.made_combo.setCurrentIndex(idx)
            if checked_combo and not checked_combo.startswith("—"):
                idx = self.checked_combo.findText(checked_combo)
                if idx >= 0:
                    self.checked_combo.setCurrentIndex(idx)

            if current_product:
                self.on_product_changed(current_product)
                if current_unit:
                    idxu = self.unit_combo.findText(current_unit)
                    if idxu >= 0:
                        self.unit_combo.setCurrentIndex(idxu)
            else:
                self.unit_combo.clear()
                self.unit_combo.addItem("")

            if show_message and not silent_errors:
                QMessageBox.information(self, "Готово", "Excel обновлён.")

        except Exception as e:
            if silent_errors:
                sys.stderr.write(f"[MirlisMark] Excel auto-refresh error: {e}\n")
                return
            QMessageBox.critical(
                self,
                "Ошибка Excel",
                f"Не удалось загрузить Excel.\n\nФайл: {self.excel_path}\nОшибка: {e}",
            )

    def fill_products(self, current_product=None):
        self.product_combo.blockSignals(True)
        self.product_combo.clear()

        names = [p["name"] for p in self.products if p.get("name")]
        for n in names:
            self.product_combo.addItem(n)

        self.product_model.setStringList(names)

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
        # если только один сотрудник — выбираем автоматически
        if self.made_combo.count() == 2:
            self.made_combo.setCurrentIndex(1)
        self.made_combo.blockSignals(False)

        self.checked_combo.blockSignals(True)
        self.checked_combo.clear()
        self.checked_combo.addItem("— не выбрано —")
        for s in self.staff_checked:
            fio = (s.get("fio") or "").strip()
            if fio:
                self.checked_combo.addItem(fio)
        # если только один цех — выбираем автоматически
        if self.checked_combo.count() == 2:
            self.checked_combo.setCurrentIndex(1)
        self.checked_combo.blockSignals(False)

    def update_excel_status(self):
        try:
            mtime = os.path.getmtime(self.excel_path)
            mtime_str = _fmt_dt_local(mtime)
        except Exception:
            mtime_str = "—"

        default_dir = app_data_dir()
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
        # показываем инструкцию с картинками перед выбором файла
        dlg = QDialog(self)
        dlg.setWindowTitle("Структура файла Excel")
        dlg.setMinimumSize(700, 500)
        dlg.resize(750, 600)

        dlg_layout = QVBoxLayout(dlg)
        dlg_layout.setContentsMargins(0, 0, 0, 12)
        dlg_layout.setSpacing(0)

        # скролл-область
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { border: none; background: #ffffff; }")

        content = QWidget()
        content.setStyleSheet("background: #ffffff;")
        lay = QVBoxLayout(content)
        lay.setContentsMargins(28, 24, 28, 24)
        lay.setSpacing(16)

        def add_title(text):
            lbl = QLabel(text)
            lbl.setStyleSheet("font-size: 20px; font-weight: 800; color: #111827; background: transparent;")
            lbl.setWordWrap(True)
            lay.addWidget(lbl)

        def add_text(text):
            lbl = QLabel(text)
            lbl.setStyleSheet("font-size: 14px; color: #374151; background: transparent;")
            lbl.setWordWrap(True)
            lay.addWidget(lbl)

        def add_image(path, max_width=640):
            if not os.path.isfile(path):
                return
            lbl = QLabel()
            lbl.setStyleSheet("background: transparent;")
            pix = QPixmap(path)
            if not pix.isNull():
                if pix.width() > max_width:
                    pix = pix.scaledToWidth(max_width, Qt.SmoothTransformation)
                lbl.setPixmap(pix)
            lay.addWidget(lbl)

        def add_separator():
            sep = QFrame()
            sep.setFrameShape(QFrame.HLine)
            sep.setStyleSheet("color: #e5e7eb; background: #e5e7eb; border: none; max-height: 1px;")
            lay.addWidget(sep)

        # --- содержимое инструкции ---
        add_title("Структура файла Excel")
        add_text(
            "Файл должен содержать 3 листа с точными названиями: «продукт», «изготовил», «цех».\n"
            "Первая строка каждого листа — заголовки колонок, данные начинаются со второй строки."
        )

        add_separator()

        add_text("Внизу файла должны быть видны 3 вкладки:")
        add_image(HELP_IMG_TABS)

        add_separator()

        add_title("1. Лист «продукт»")
        add_text(
            "Заголовки: Код | Наименование | Срок годности (ч) | Ед. измер. | Активен | Комментарий\n\n"
            "• «Ед. измер.» — через запятую: кг, шт или кг,шт\n"
            "• «Активен» — Да или 1 (показывать), Нет или 0 (скрыть)"
        )
        add_image(HELP_IMG_PRODUCT)

        add_separator()

        add_title("2. Лист «изготовил»")
        add_text("Заголовки: ФИО | Активен")
        add_image(HELP_IMG_MADE)

        add_separator()

        add_title("3. Лист «цех»")
        add_text("Заголовки: Цех | Активен")
        add_image(HELP_IMG_WORKSHOP)

        add_separator()

        add_text(
            "Если в листе «изготовил» или «цех» только одна запись — "
            "она подставится автоматически, и работнику не нужно будет её выбирать каждый раз."
        )

        scroll.setWidget(content)
        dlg_layout.addWidget(scroll, 1)

        # кнопки
        btn_row = QHBoxLayout()
        btn_row.setContentsMargins(20, 0, 20, 0)
        btn_row.setSpacing(12)
        btn_row.addStretch(1)

        btn_cancel = QPushButton("Отмена")
        btn_cancel.setMinimumHeight(40)
        btn_cancel.setCursor(Qt.PointingHandCursor)
        btn_cancel.clicked.connect(dlg.reject)

        btn_ok = QPushButton("Выбрать файл")
        btn_ok.setMinimumHeight(40)
        btn_ok.setCursor(Qt.PointingHandCursor)
        btn_ok.setStyleSheet(
            "QPushButton { background: #16a34a; color: #ffffff; font-weight: 700; "
            "border: 1px solid #15803d; border-radius: 12px; padding: 8px 24px; }"
            "QPushButton:hover { background: #15803d; }"
        )
        btn_ok.clicked.connect(dlg.accept)

        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        dlg_layout.addLayout(btn_row)

        if dlg.exec_() != QDialog.Accepted:
            return

        current_dir = os.path.dirname(self.excel_path)

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл Excel с данными",
            current_dir,
            "Excel файлы (*.xlsx *.xls);;Все файлы (*)",
        )

        if not file_path:
            return

        self.excel_path = file_path
        settings = _load_settings()
        settings["excel_path"] = file_path
        _save_settings(settings)

        self.reload_excel(show_message=True)

    # ---------------- Helpers ----------------
    def get_product(self, name):
        name = (name or "").strip()
        return next((p for p in self.products if (p.get("name") or "").strip() == name), None)

    def _clear_layout(self, layout):
        while layout.count():
            item = layout.takeAt(0)
            w = item.widget()
            child_layout = item.layout()
            if child_layout is not None:
                self._clear_layout(child_layout)
            if w is not None:
                w.deleteLater()

    # ---------------- History helpers ----------------
    def _filtered_history_entries(self):
        if not self._history_filter_text:
            return list(self.history_entries)
        q = self._history_filter_text
        result = []
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

        self._clear_layout(self.history_list_layout)

        for e in page_items:
            card = QFrame()
            card.setObjectName("HistoryCard")
            card.setCursor(Qt.PointingHandCursor)
            card.setProperty("selected", (e.get("id") == getattr(self, "_selected_history_id", None)))
            card_layout = QVBoxLayout(card)
            card_layout.setContentsMargins(10, 8, 10, 8)
            card_layout.setSpacing(4)

            top_row = QHBoxLayout()
            top_row.setSpacing(6)

            prod_label = QLabel(str(e.get("product", "")))
            prod_label.setStyleSheet("font-weight: 600;")
            prod_label.setAttribute(Qt.WA_TransparentForMouseEvents, True)
            qty_label = QLabel(str(e.get("qty", "")))
            qty_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            qty_label.setStyleSheet("font-weight: 600; color: #111827;")
            qty_label.setAttribute(Qt.WA_TransparentForMouseEvents, True)

            top_row.addWidget(prod_label, 1)
            top_row.addWidget(qty_label, 0)

            made = str(e.get("made", ""))
            checked = str(e.get("checked", ""))
            mid_parts = [p for p in [made, checked] if p]
            mid_text = " · ".join(mid_parts) if mid_parts else ""
            mid_row = QLabel(mid_text)
            mid_row.setStyleSheet("color: #6b7280; font-size: 12px;")
            mid_row.setAttribute(Qt.WA_TransparentForMouseEvents, True)

            bottom_row = QHBoxLayout()
            bottom_row.setSpacing(6)

            time_label = QLabel(str(e.get("time", "")))
            time_label.setStyleSheet("color: #9ca3af; font-size: 12px;")
            time_label.setAttribute(Qt.WA_TransparentForMouseEvents, True)

            batch = str(e.get("batch", ""))
            batch_label = QLabel(f"№ {batch}" if batch else "")
            batch_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            batch_label.setStyleSheet("color: #6b7280; font-size: 12px;")
            batch_label.setAttribute(Qt.WA_TransparentForMouseEvents, True)

            bottom_row.addWidget(time_label, 1)
            bottom_row.addWidget(batch_label, 0)

            card_layout.addLayout(top_row)
            card_layout.addWidget(mid_row)
            card_layout.addLayout(bottom_row)

            card.mousePressEvent = (lambda ev, ent=e: self._on_history_clicked(ent))
            card.style().unpolish(card)
            card.style().polish(card)
            self.history_list_layout.addWidget(card)

        self.history_list_layout.addStretch(1)

        if hasattr(self, "history_page_label"):
            self.history_page_label.setText(f"Страница {self.history_page + 1} из {pages}")
        if hasattr(self, "history_prev_btn"):
            self.history_prev_btn.setEnabled(self.history_page > 0)
        if hasattr(self, "history_next_btn"):
            self.history_next_btn.setEnabled(self.history_page < pages - 1)

    def _on_history_search_text_changed(self, text):
        self._history_filter_text = (text or "").strip().lower()
        self.history_page = 0
        self._rebuild_history_view()

    def _change_history_page(self, delta):
        self.history_page += delta
        self._rebuild_history_view()

    def _build_history_entry_from_label(self, label, qty_display, unit_ui):
        produced_at = getattr(label, "produced_at", datetime.now())
        preview_text = (
            f"{getattr(label, 'weekday', '')}\n"
            f"Продукт: {getattr(label, 'product_name', '')}\n"
            f"Вес/шт: {getattr(label, 'qty_value', '')} {getattr(label, 'qty_unit_ru', '')}\n"
            f"Дата/время: {format_dt(produced_at)}\n"
            f"№ партии: {getattr(label, 'batch', '')}\n"
            f"Годен до: {format_dt(getattr(label, 'expires_at', produced_at))}\n"
            f"Изготовил: {getattr(label, 'made_by', '')}\n"
            f"Цех: {getattr(label, 'checked_by', '')}\n"
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
            "product": getattr(label, "product_name", ""),
            "qty": f"{qty_display} {unit_ui}".strip(),
            "made": getattr(label, "made_by", ""),
            "checked": getattr(label, "checked_by", ""),
            "time": format_dt(produced_at),
            "batch": getattr(label, "batch", ""),
        }

    def _append_history_entry(self, entry):
        if not hasattr(self, "history_entries") or not isinstance(self.history_entries, list):
            self.history_entries = []
        self.history_entries.insert(0, entry)
        if not self._history_filter_text:
            self.history_page = 0
        self._rebuild_history_view()

    def _on_history_clicked(self, entry):
        self._selected_history_id = entry.get("id")
        self.apply_history_entry(entry)
        self._rebuild_history_view()

    def apply_history_entry(self, entry):
        self._loading_from_history = True
        try:
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

            idx = self.product_combo.findText(product_name)
            if idx >= 0:
                self.product_combo.setCurrentIndex(idx)
            else:
                self.product_combo.setEditText(product_name)

            self.on_product_changed(product_name)

            self.unit_combo.blockSignals(True)
            try:
                if unit_ui:
                    idxu = self.unit_combo.findText(unit_ui)
                    if idxu >= 0:
                        self.unit_combo.setCurrentIndex(idxu)
            finally:
                self.unit_combo.blockSignals(False)

            self.qty_input.setText(qty_value)

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

    def on_product_changed(self, product_name):
        self.unit_combo.blockSignals(True)
        self.unit_combo.clear()
        self.unit_combo.addItem("")

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

    def _step_for_unit(self):
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
        self.product_combo.setCurrentIndex(-1)
        self.product_combo.setEditText("")

        self.unit_combo.setCurrentIndex(0)
        self.qty_input.clear()

        self.made_manual.setChecked(False)
        self.checked_manual.setChecked(False)

        # если в списке один вариант — оставляем автовыбор, иначе сбрасываем
        if self.made_combo.count() == 2:
            self.made_combo.setCurrentIndex(1)
        else:
            self.made_combo.setCurrentIndex(0)

        if self.checked_combo.count() == 2:
            self.checked_combo.setCurrentIndex(1)
        else:
            self.checked_combo.setCurrentIndex(0)

        self.made_input.clear()
        self.checked_input.clear()

        # Сбрасываем только следы ручного редактирования,
        # но НЕ навязываем единый размер шрифта всем форматам.
        if hasattr(self, "preview"):
            self.preview.clearFocus()
            self.preview.clear()

        if hasattr(self, "btn_bold"):
            self.btn_bold.blockSignals(True)
            self.btn_bold.setChecked(False)
            self.btn_bold.blockSignals(False)

        if hasattr(self, "btn_italic"):
            self.btn_italic.blockSignals(True)
            self.btn_italic.setChecked(False)
            self.btn_italic.blockSignals(False)

        if hasattr(self, "btn_underline"):
            self.btn_underline.blockSignals(True)
            self.btn_underline.setChecked(False)
            self.btn_underline.blockSignals(False)

        if hasattr(self, "btn_align_left"):
            self.btn_align_left.blockSignals(True)
            self.btn_align_left.setChecked(False)
            self.btn_align_left.blockSignals(False)

        if hasattr(self, "btn_align_center"):
            self.btn_align_center.blockSignals(True)
            self.btn_align_center.setChecked(False)
            self.btn_align_center.blockSignals(False)

        if hasattr(self, "btn_align_right"):
            self.btn_align_right.blockSignals(True)
            self.btn_align_right.setChecked(False)
            self.btn_align_right.blockSignals(False)

        self._user_edited_preview = False

        # Важно: после сброса пусть каждая этикетка заново
        # применит СВОИ дефолтные автонастройки.
        self.refresh_preview(force=True)
    # ---------------- Preview / validation ----------------
    def _unit_code_from_ui(self, unit_text):
        if unit_text == "кг":
            return "kg"
        if unit_text == "шт":
            return "pcs"
        return None

    def _made_value(self):
        if self.made_manual.isChecked():
            return self.made_input.text().strip()
        val = self.made_combo.currentText().strip()
        return "" if val.startswith("—") else val

    def _checked_value(self):
        if self.checked_manual.isChecked():
            return self.checked_input.text().strip()
        val = self.checked_combo.currentText().strip()
        return "" if val.startswith("—") else val

    _WEEKDAYS = {"ПОНЕДЕЛЬНИК", "ВТОРНИК", "СРЕДА", "ЧЕТВЕРГ", "ПЯТНИЦА", "СУББОТА", "ВОСКРЕСЕНЬЕ"}

    def _set_preview_text_programmatically(self, text):
        self._updating_preview = True
        try:
            self.preview.blockSignals(True)
            self.preview.setPlainText(text)

            font_family = self.font_combo.currentFont().family()
            preview_scale = getattr(self, "_preview_scale", 1.0) or 1.0
            effective_base_font = float(self._base_font_size) * float(preview_scale)
            cursor = self.preview.textCursor()

            cursor.select(QTextCursor.Document)
            fmt_base = QTextCharFormat()
            fmt_base.setFontFamily(font_family)
            fmt_base.setFontPointSize(effective_base_font)
            fmt_base.setFontWeight(QFont.Normal)
            cursor.mergeCharFormat(fmt_base)
            cursor.clearSelection()

            is_colored = hasattr(self, "label_size_combo") and self.label_size_combo.currentText() == "Цветные"
            weekday_line_index = 1 if is_colored else 0

            lines = text.split("\n")
            weekday_line = ""
            if len(lines) > weekday_line_index:
                weekday_line = lines[weekday_line_index].strip()

            if weekday_line in self._WEEKDAYS:
                cursor.movePosition(QTextCursor.Start)

                for _ in range(weekday_line_index):
                    cursor.movePosition(QTextCursor.NextBlock)

                cursor.movePosition(QTextCursor.StartOfBlock)
                cursor.movePosition(QTextCursor.EndOfBlock, QTextCursor.KeepAnchor)

                fmt_weekday = QTextCharFormat()
                fmt_weekday.setFontFamily(font_family)

                if is_colored:
                    fmt_weekday.setFontPointSize(max(8.0, effective_base_font * 1.8))
                else:
                    fmt_weekday.setFontPointSize(max(8.0, effective_base_font * 1.3))

                fmt_weekday.setFontWeight(QFont.Bold)
                cursor.mergeCharFormat(fmt_weekday)

                block_fmt = QTextBlockFormat()
                block_fmt.setAlignment(Qt.AlignHCenter)
                cursor.mergeBlockFormat(block_fmt)

            cursor.clearSelection()
            cursor.movePosition(QTextCursor.Start)
            self.preview.setTextCursor(cursor)
        finally:
            self.preview.blockSignals(False)
            self._updating_preview = False

    def _set_preview_html_programmatically(self, html):
        self._updating_preview = True
        try:
            self.preview.blockSignals(True)
            self.preview.setHtml(html)
            cursor = self.preview.textCursor()
            cursor.movePosition(QTextCursor.Start)
            self.preview.setTextCursor(cursor)
        finally:
            self.preview.blockSignals(False)
            self._updating_preview = False

    def _build_label_plain_text(self):
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

        produced_at = None
        if getattr(self, "_preview_manual_mode", False):
            d = self.manual_datetime_picker.date()
            t = self.manual_datetime_picker.time_()
            produced_at = datetime(d.year(), d.month(), d.day(), t.hour(), t.minute())

        label = build_label(
            product_name=product["name"],
            shelf_life_hours=product["shelf_life_hours"],
            qty_value=str(qty_float).rstrip("0").rstrip("."),
            unit=unit_code,
            made_by=made_by,
            checked_by=checked_by,
            produced_at=produced_at,
        )

        text_parts = []
        is_colored = hasattr(self, "label_size_combo") and self.label_size_combo.currentText() == "Цветные"

        if is_colored:
            text_parts.append("")
            text_parts.append("")
            text_parts.append("")
            text_parts.append("")
        else:
            text_parts.append(f"{label.weekday}")

        text_parts.append(f"Продукт: {label.product_name}")
        text_parts.append(f"Вес/шт: {label.qty_value} {label.qty_unit_ru}")
        text_parts.append(f"Дата/время: {format_dt(label.produced_at)}")
        text_parts.append(f"№ партии: {label.batch}")
        text_parts.append(f"Годен до: {format_dt(label.expires_at)}")
        text_parts.append(f"Изготовил: {label.made_by}")
        text_parts.append(f"Цех: {label.checked_by}")

        text = "\n".join(text_parts) + "\n"
        return (text, True)

    def refresh_preview(self, force=False):
        if getattr(self, "_loading_from_history", False):
            return
        text, can_print = self._build_label_plain_text()
        self.print_btn.setEnabled(can_print)

        if self._user_edited_preview:
            return

        if self.preview.hasFocus() and not force:
            return

        self._set_preview_text_programmatically(text)

    def _on_preview_text_changed(self):
        if self._updating_preview:
            return
        self._user_edited_preview = True

    # ---------------- Editor formatting (selection-only) ----------------
    def _merge_format_on_selection(self, fmt):
        cursor = self.preview.textCursor()
        if not cursor.hasSelection():
            cursor.select(QTextCursor.WordUnderCursor)
        cursor.mergeCharFormat(fmt)
        self.preview.mergeCurrentCharFormat(fmt)

    def _toggle_bold_on_selection(self):
        fmt = QTextCharFormat()
        desired_bold = self.btn_bold.isChecked()
        fmt.setFontWeight(QFont.Bold if desired_bold else QFont.Normal)
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
        cursor = self.preview.textCursor()
        fmt = cursor.charFormat() if cursor.charFormat().isValid() else self.preview.currentCharFormat()

        self.btn_bold.blockSignals(True)
        self.btn_italic.blockSignals(True)
        self.btn_underline.blockSignals(True)
        self.btn_align_left.blockSignals(True)
        self.btn_align_center.blockSignals(True)
        self.btn_align_right.blockSignals(True)
        try:
            self.btn_bold.setChecked(fmt.fontWeight() >= QFont.Bold)
            self.btn_italic.setChecked(fmt.fontItalic())
            self.btn_underline.setChecked(fmt.fontUnderline())

            align = cursor.blockFormat().alignment()
            self.btn_align_left.setChecked(align == Qt.AlignLeft)
            self.btn_align_center.setChecked(align == Qt.AlignHCenter)
            self.btn_align_right.setChecked(align == Qt.AlignRight)
        finally:
            self.btn_bold.blockSignals(False)
            self.btn_italic.blockSignals(False)
            self.btn_underline.blockSignals(False)
            self.btn_align_left.blockSignals(False)
            self.btn_align_center.blockSignals(False)
            self.btn_align_right.blockSignals(False)
    
        self._sync_font_size_from_cursor()

    def _sync_font_size_from_cursor(self):
        if not hasattr(self, "font_size_combo"):
            return

        cursor = self.preview.textCursor()
        fmt = cursor.charFormat()

        size = fmt.fontPointSize()
        if size <= 0:
            size = fmt.font().pointSizeF()
        if size <= 0:
            size = float(self._base_font_size)

        size_int = int(round(size))
        if size_int <= 0:
            size_int = self._base_font_size

        self.font_size_combo.blockSignals(True)
        self.font_size_combo.lineEdit().setText(str(size_int))
        self.font_size_combo.blockSignals(False)

    def _set_alignment(self, align_flag):
        if align_flag == Qt.AlignLeft:
            self.preview.setAlignment(Qt.AlignLeft)
        elif align_flag == Qt.AlignHCenter:
            self.preview.setAlignment(Qt.AlignHCenter)
        elif align_flag == Qt.AlignRight:
            self.preview.setAlignment(Qt.AlignRight)
        elif align_flag == Qt.AlignJustify:
            self.preview.setAlignment(Qt.AlignJustify)
        self._sync_format_toolbar_from_cursor()

    def _set_font_family_on_selection(self, font):
        fmt = QTextCharFormat()
        fmt.setFontFamily(font.family())
        self._merge_format_on_selection(fmt)

    def _change_font_size(self, delta):
        self._base_font_size = max(8, min(72, self._base_font_size + delta))
        fmt = QTextCharFormat()
        fmt.setFontPointSize(float(self._base_font_size))
        self._merge_format_on_selection(fmt)
        if hasattr(self, "font_size_combo"):
            self.font_size_combo.blockSignals(True)
            self.font_size_combo.setCurrentText(str(self._base_font_size))
            self.font_size_combo.blockSignals(False)

    def on_font_size_combo_changed(self, text):
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
        if hasattr(self, "font_size_combo"):
            self.font_size_combo.blockSignals(True)
            self.font_size_combo.setCurrentText(str(size))
            self.font_size_combo.blockSignals(False)

    def _get_copies(self):
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

    def _apply_copies_to_tspl(self, tspl, copies):
        lines = tspl.strip().splitlines()
        for i in range(len(lines) - 1, -1, -1):
            if lines[i].strip().upper().startswith("PRINT"):
                lines[i] = f"PRINT {copies}"
                return "\n".join(lines).strip()
        return (tspl.strip() + f"\nPRINT {copies}").strip()

    # ---------------- Rendering preview → TSPL bitmap ----------------
    def _render_preview_to_tspl_bytes(
        self,
        label_w_mm=None,
        label_h_mm=None,
        dpi=203,
        density=10,
        speed=4,
        threshold=200,
        copies=1,
    ):
        w_mm = label_w_mm if label_w_mm is not None else getattr(self, "label_w_mm", 58.0)
        h_mm = label_h_mm if label_h_mm is not None else getattr(self, "label_h_mm", 80.0)
        w_px = int(w_mm / 25.4 * dpi)
        h_px = int(h_mm / 25.4 * dpi)

        doc = self.preview.document().clone()

        vp_w = self.preview.viewport().width()
        vp_h = self.preview.viewport().height()

        doc.setPageSize(QSizeF(vp_w, vp_h))

        scale_x = w_px / vp_w
        scale_y = h_px / vp_h

        img = QImage(w_px, h_px, QImage.Format_RGB32)
        img.fill(QColor(255, 255, 255))

        painter = QPainter(img)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setRenderHint(QPainter.TextAntialiasing, True)
        painter.scale(scale_x, scale_y)
        doc.drawContents(painter)
        painter.end()

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

        gap_mm = 2.0
        header = (
            f"SIZE {w_mm} mm, {h_mm} mm\r\n"
            f"GAP {gap_mm} mm, 0 mm\r\n"
        ).encode("ascii")
        header += (
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
            copies = self._get_copies()
            tspl_bytes = self._render_preview_to_tspl_bytes(threshold=200, copies=copies)
            n_bytes = print_raw(printer_name, tspl_bytes)
            print(f"SENDING BITMAP... {n_bytes} bytes, copies={copies}")
        except Exception as e:
            QMessageBox.warning(self, "Печать", f"Не удалось отправить на печать:\n{e}")
            return

        self.repeat_btn.setEnabled(True)
        self.last_printed_preview_text = preview_text
        self._last_printed_tspl_bytes = tspl_bytes

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
        try:
            self._archive_printed_label(preview_text=preview_text, entry=entry, copies=copies)
        except Exception as e:
            sys.stderr.write(f"[MirlisMark] Archive save error: {e}\n")
        self._append_history_entry(entry)

    def repeat_last_print(self):
        tspl_bytes = getattr(self, "_last_printed_tspl_bytes", None)
        if not tspl_bytes:
            return
        try:
            printer_name = win32print.GetDefaultPrinter()
            copies = self._get_copies()
            if copies != 1:
                tspl_bytes = tspl_bytes.rsplit(b"\r\nPRINT ", 1)[0] + f"\r\nPRINT {copies}\r\n".encode("ascii")
            print_raw(printer_name, tspl_bytes)
            if self.last_history_entry is not None:
                e = dict(self.last_history_entry)
                e["id"] = time.time_ns()
                e["ts"] = time.time()
                try:
                    self._archive_printed_label(preview_text=e.get("preview_text", ""), entry=e, copies=copies)
                except Exception as arch_e:
                    sys.stderr.write(f"[MirlisMark] Archive save error: {arch_e}\n")
                self._append_history_entry(e)
        except Exception as e:
            QMessageBox.warning(self, "Повтор", f"Не удалось повторить печать:\n{e}")

    # ---------------- Printed labels archive ----------------
    def _labels_archive_root(self) -> str:
        return os.path.join(app_data_dir(), "Готовые этикетки")

    def _labels_archive_day_dir(self) -> str:
        day = datetime.now().strftime("%d.%m.%Y")
        return os.path.join(self._labels_archive_root(), day)

    def _sanitize_filename_part(self, s: str) -> str:
        s = (s or "").strip()
        if not s:
            return "—"
        for ch in ['<', '>', ':', '"', '/', '\\', '|', '?', '*']:
            s = s.replace(ch, "_")
        s = s.replace("\n", " ").replace("\r", " ").strip()
        while "  " in s:
            s = s.replace("  ", " ")
        return s[:120].strip() or "—"

    def _cleanup_old_label_archives(self, days: int = 31):
        root = self._labels_archive_root()
        try:
            if not os.path.isdir(root):
                return
            today = datetime.now().date()
            for name in os.listdir(root):
                p = os.path.join(root, name)
                if not os.path.isdir(p):
                    continue
                try:
                    d = datetime.strptime(name.strip(), "%d.%m.%Y").date()
                except Exception:
                    continue
                age_days = (today - d).days
                if age_days >= days:
                    shutil.rmtree(p, ignore_errors=True)
        except Exception as e:
            sys.stderr.write(f"[MirlisMark] Archive cleanup error: {e}\n")

    def _archive_printed_label(self, preview_text: str, entry: dict | None, copies: int):
        # очистка старых папок — при каждом сохранении
        self._cleanup_old_label_archives(days=31)

        e = entry or {}
        product = self._sanitize_filename_part(e.get("product") or e.get("product_name") or "Этикетка")
        weight = self._sanitize_filename_part((e.get("qty") or "").replace(" ", "") or "—")
        batch = self._sanitize_filename_part(e.get("batch") or "—")

        base_name = f"{product}_{weight}_{batch}"
        day_dir = self._labels_archive_day_dir()
        os.makedirs(day_dir, exist_ok=True)

        content = (preview_text or "").rstrip()
        try:
            copies_int = int(copies)
        except Exception:
            copies_int = 1
        content = f"{content}\nКоличество: {copies_int}"

        n = 1
        while True:
            suffix = "" if n == 1 else f"_{n}"
            fname = f"{base_name}{suffix}.txt"
            fpath = os.path.join(day_dir, fname)
            if not os.path.exists(fpath):
                break
            n += 1

        with open(fpath, "w", encoding="utf-8") as f:
            f.write(content + "\n")


def main():
    try:
        import ctypes
        user32 = ctypes.windll.user32
        user32.SetProcessDPIAware()
        screen_w = user32.GetSystemMetrics(0)
        screen_h = user32.GetSystemMetrics(1)
        scale_w = screen_w / 1920
        scale_h = screen_h / 1080
        scale = min(scale_w, scale_h)
        if scale < 0.95:
            os.environ["QT_SCALE_FACTOR"] = str(round(scale, 3))
    except Exception:
        pass

    fmt = QSurfaceFormat()
    fmt.setRenderableType(QSurfaceFormat.OpenGL)
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

    sys.exit(app.exec_())

if __name__ == "__main__":
    main()



















































