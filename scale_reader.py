# scale_reader.py
# Чтение веса с весов Тензо-М (и совместимых) через RS-232/USB-Serial.
# Автоопределение COM-порта, парсинг строки, фоновый QThread.

from __future__ import annotations

import time
import re
from typing import Optional

from PyQt5.QtCore import QThread, pyqtSignal

try:
    import serial
    from serial.tools import list_ports
    SERIAL_AVAILABLE = True
except ImportError:
    SERIAL_AVAILABLE = False


# USB-Serial VID для приоритетного сканирования
_PRIORITY_VIDS = {
    0x1A86,  # CH340/CH341 (WCH) — адаптер пользователя
    0x0403,  # FTDI
    0x067B,  # Prolific PL2303
    0x10C4,  # Silicon Labs CP210x
    0x2341,  # Arduino-compatible
}


def list_scale_ports() -> list:
    """Возвращает список доступных COM-портов в порядке приоритета.
    Сначала идут известные USB-Serial адаптеры (CH340, FTDI и т.п.),
    затем все остальные.
    """
    if not SERIAL_AVAILABLE:
        return []
    try:
        ports = list(list_ports.comports())
    except Exception:
        return []

    def priority(p):
        vid = getattr(p, "vid", None) or 0
        return 0 if vid in _PRIORITY_VIDS else 1

    ports.sort(key=priority)
    result = []
    for p in ports:
        if p.device and p.device not in result:
            result.append(p.device)
    return result


def parse_weight_line(line: str) -> Optional[float]:
    """Парсит строку от весов, возвращает вес в КГ или None.

    Поддерживает форматы вроде:
        "+  1.234"        → 1.234
        "001.234 kg"      → 1.234
        "1234"            → 1.234 (граммы → кг, эвристика)
        "  0.50 kg"       → 0.5
    """
    if not line:
        return None
    m = re.search(r"[+-]?\s*(\d+[.,]\d+|\d+)", line)
    if not m:
        return None
    raw_num = m.group(1)
    val_str = raw_num.replace(",", ".")
    try:
        val = float(val_str)
    except Exception:
        return None
    # Эвристика: целое число > 200 без дробной точки — вероятно граммы
    if val > 200 and "." not in raw_num and "," not in raw_num:
        val = val / 1000.0
    return round(val, 3)


class ScaleReader(QThread):
    """Автоопределение COM-порта и чтение веса в фоновом потоке.

    Сигналы:
        weight_received(float) — успешно распознанный вес в кг
        raw_received(str)      — сырая диагностика (для лога)
        error_occurred(str)    — текст ошибки для пользователя
        progress(str)          — текущее действие (для индикации на кнопке)
        port_found(str)        — какой порт оказался рабочим (для сохранения)
    """

    weight_received = pyqtSignal(float)
    raw_received    = pyqtSignal(str)
    error_occurred  = pyqtSignal(str)
    progress        = pyqtSignal(str)
    port_found      = pyqtSignal(str)

    def __init__(self, preferred_port: str = "", baud: int = 9600,
                 per_port_timeout: float = 1.5, saved_port_timeout: float = 2.5,
                 parent=None):
        super().__init__(parent)
        self.preferred_port     = preferred_port or ""
        self.baud               = baud
        self.per_port_timeout   = per_port_timeout
        self.saved_port_timeout = saved_port_timeout

    def _try_port(self, port: str, timeout: float) -> Optional[float]:
        """Пытается прочитать вес с одного порта. Возвращает кг или None."""
        try:
            with serial.Serial(
                port=port,
                baudrate=self.baud,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=0.4,
            ) as ser:
                deadline = time.monotonic() + timeout
                while time.monotonic() < deadline:
                    raw = ser.readline()
                    if not raw:
                        continue
                    try:
                        line = raw.decode("ascii", errors="replace").strip()
                    except Exception:
                        continue
                    if not line:
                        continue
                    val = parse_weight_line(line)
                    if val is not None:
                        self.raw_received.emit(f"{port}: {line!r}  →  {val:.3f} кг")
                        return val
        except Exception as exc:
            self.raw_received.emit(f"{port}: {exc}")
        return None

    def run(self):
        if not SERIAL_AVAILABLE:
            self.error_occurred.emit(
                "Модуль pyserial не установлен.\n"
                "Выполните: python -m pip install pyserial"
            )
            return

        # 1) Сначала сохранённый порт (если есть), потом остальные по приоритету
        all_ports = list_scale_ports()
        ports_order = []
        if self.preferred_port:
            ports_order.append(self.preferred_port)
        for p in all_ports:
            if p not in ports_order:
                ports_order.append(p)

        if not ports_order:
            self.error_occurred.emit(
                "Не найдено ни одного COM-порта.\n\n"
                "Проверьте:\n"
                "  • USB-кабель адаптера подключён к компьютеру\n"
                "  • Адаптер виден в Диспетчере устройств\n"
                "  • Установлен драйвер CH340 (или другого адаптера)"
            )
            return

        # 2) Перебор: сохранённый порт с увеличенным таймаутом, остальные быстро
        for idx, port in enumerate(ports_order):
            is_preferred = (idx == 0 and port == self.preferred_port)
            timeout = self.saved_port_timeout if is_preferred else self.per_port_timeout
            self.progress.emit(f"Поиск весов: {port}")
            value = self._try_port(port, timeout)
            if value is not None:
                self.port_found.emit(port)
                self.weight_received.emit(value)
                return

        # 3) Никто не ответил
        self.error_occurred.emit(
            "Весы не найдены.\n\n"
            "Проверьте:\n"
            "  • USB-кабель адаптера подключён\n"
            "  • Весы включены и кабель RS-232 от них подключён к адаптеру\n"
            "  • На весы положен какой-то груз (или нажмите кнопку выдачи)\n"
            "  • В настройках весов установлен режим автоматической передачи"
        )
