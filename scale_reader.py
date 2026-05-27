# scale_reader.py
# Чтение веса с преобразователя ТВ-003/05Н и других приборов ВИК "Тензо-М"
# по протоколу обмена "Тензо-М" через RS-232 / USB-Serial.
#
# Прибор работает в режиме "запрос-ответ": компьютер посылает команду,
# прибор отвечает (РЭ ТВ-003/05Н, п. 23.1, 30.5).
#
# Что делает этот модуль:
#   • сам находит COM-порт, перебирает СКОРОСТЬ, сетевой адрес и режим CRC;
#   • когда находит рабочую связку — сообщает её (сигнал config_found),
#     чтобы приложение запомнило и в следующий раз подключалось мгновенно;
#   • дожидается "успокоения" веса и отдаёт стабильное значение.
#
# Кадр ПК -> прибор:  FF Adr COP [CRC] FF FF   (конец кадра — два FF подряд)
# Ответ:              FF Adr COP Data... [CRC] FF FF
# Команды веса: 0xC2 = НЕТТО, 0xC3 = БРУТТО.
# Ответ веса: Adr, COP, W0, W1, W2 (BCD, младший первым), CON.
# CON: D7=знак минус, D4=успокоение, D3=перегруз, D2..D0=позиция запятой.

from __future__ import annotations

import time
from typing import List, Optional, Tuple

from PyQt5.QtCore import QThread, pyqtSignal

try:
    import serial
    from serial.tools import list_ports
    SERIAL_AVAILABLE = True
except ImportError:
    SERIAL_AVAILABLE = False


# ── Константы протокола ─────────────────────────────────────────────
_SEP   = 0xFF
_STUFF = 0xFE
COP_NETTO  = 0xC2
COP_BRUTTO = 0xC3
_COP_WEIGHT = (COP_NETTO, COP_BRUTTO)

# Скорости для перебора (по убыванию вероятности). Заданная baud идёт первой.
_DEFAULT_BAUDS = [9600, 4800, 2400, 19200, 38400, 57600, 115200, 14400, 28800]

_PRIORITY_VIDS = {0x1A86, 0x0403, 0x067B, 0x10C4, 0x2341}
_PRIORITY_DESC = ("tenso", "uart bridge", "ch340", "ch341", "cp210", "ftdi", "prolific")


def list_scale_ports() -> list:
    if not SERIAL_AVAILABLE:
        return []
    try:
        ports = list(list_ports.comports())
    except Exception:
        return []

    def priority(p):
        vid = getattr(p, "vid", None) or 0
        desc = (getattr(p, "description", "") or "").lower()
        return 0 if (vid in _PRIORITY_VIDS or any(k in desc for k in _PRIORITY_DESC)) else 1

    ports.sort(key=priority)
    result = []
    for p in ports:
        if p.device and p.device not in result:
            result.append(p.device)
    return result


# ── CRC "Тензо-М" (полином 0x69) ────────────────────────────────────
def _crc_byte(b_input: int, b_crc: int) -> int:
    al, ah = b_input & 0xFF, b_crc & 0xFF
    for _ in range(8):
        cf = (al >> 7) & 1
        al = ((al << 1) | cf) & 0xFF
        new_cf = (ah >> 7) & 1
        ah = ((ah << 1) | cf) & 0xFF
        if new_cf:
            ah ^= 0x69
    return ah & 0xFF


def tenso_crc(data: bytes) -> int:
    crc = 0
    for b in data:
        crc = _crc_byte(b, crc)
    return _crc_byte(0x00, crc)


# ── Сборка запроса / разбор кадра ───────────────────────────────────
def _stuff(payload: bytes) -> bytes:
    out = bytearray()
    for b in payload:
        out.append(b)
        if b == _SEP:
            out.append(_STUFF)
    return bytes(out)


def build_request(addr: int, cop: int, with_crc: bool) -> bytes:
    body = bytes([addr & 0xFF, cop & 0xFF])
    if with_crc:
        body += bytes([tenso_crc(body)])
    return bytes([_SEP]) + _stuff(body) + bytes([_SEP, _SEP])


def _extract_frame(buf: bytearray) -> Optional[bytes]:
    i, n = 0, len(buf)
    while i < n and buf[i] == _SEP:
        i += 1
    if i >= n:
        return None
    body = bytearray()
    while i < n:
        b = buf[i]
        if b == _SEP:
            if i + 1 >= n:
                return None
            if buf[i + 1] == _STUFF:
                body.append(_SEP)
                i += 2
                continue
            return bytes(body)
        body.append(b)
        i += 1
    return None


def _bcd(b: int) -> int:
    return (b >> 4) * 10 + (b & 0x0F)


def decode_weight(body: bytes) -> Optional[Tuple[float, bool, bool]]:
    """(вес_кг, стабильно, перегруз) либо None."""
    if not body or len(body) < 6 or body[1] not in _COP_WEIGHT:
        return None
    w0, w1, w2, con = body[2], body[3], body[4], body[5]
    intval = _bcd(w2) * 10000 + _bcd(w1) * 100 + _bcd(w0)
    point  = con & 0x07
    sign   = -1 if (con & 0x80) else 1
    stable = bool(con & 0x10)
    overload = bool(con & 0x08)
    return (round(sign * intval / (10 ** point), 3), stable, overload)


class ScaleReader(QThread):
    """Фоновый поиск и чтение веса.

    Сигналы:
        weight_received(float)            — вес в кг
        config_found(str, int, int, bool) — (порт, скорость, адрес, CRC) для запоминания
        port_found(str)                   — порт (для обратной совместимости)
        raw_received(str)                 — диагностика в лог
        error_occurred(str)               — ошибка для пользователя
        progress(str)                     — текущее действие (надпись на кнопке)
    """

    weight_received = pyqtSignal(float)
    config_found    = pyqtSignal(str, int, int, bool)
    port_found      = pyqtSignal(str)
    raw_received    = pyqtSignal(str)
    error_occurred  = pyqtSignal(str)
    progress        = pyqtSignal(str)

    def __init__(self, preferred_port: str = "", baud: int = 9600,
                 per_port_timeout: float = 1.5, saved_port_timeout: float = 2.5,
                 address: int = 1, command: int = COP_BRUTTO,
                 stabilize_timeout: float = 4.0,
                 # запомненная связка для мгновенного подключения:
                 preferred_baud: int = 0, preferred_address: int = 0,
                 preferred_crc: Optional[bool] = None,
                 parent=None):
        super().__init__(parent)
        self.preferred_port    = preferred_port or ""
        self.command           = command
        self.stabilize_timeout = stabilize_timeout

        # Перебор скоростей: заданная — первой, без дублей
        bauds = [baud] + [b for b in _DEFAULT_BAUDS if b != baud]
        self._bauds = bauds

        # Перебор адресов: заданный — первым, затем 1..8
        self._addrs = [address] + [a for a in range(1, 9) if a != address]

        # Запомненная связка (если есть)
        self.preferred_baud    = preferred_baud
        self.preferred_address = preferred_address
        self.preferred_crc     = preferred_crc

    # ── низкоуровневое ──
    def _open(self, port: str, baud: int):
        return serial.Serial(
            port=port, baudrate=baud,
            bytesize=serial.EIGHTBITS, parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE, timeout=0.2, write_timeout=0.5,
        )

    def _read_raw(self, ser, timeout: float) -> bytes:
        deadline = time.monotonic() + timeout
        buf = bytearray()
        while time.monotonic() < deadline:
            chunk = ser.read(128)
            if chunk:
                buf.extend(chunk)
                if _extract_frame(buf) is not None:
                    break
        return bytes(buf)

    def _exchange(self, ser, addr: int, with_crc: bool, timeout: float) -> bytes:
        try:
            ser.reset_input_buffer()
            ser.write(build_request(addr, self.command, with_crc))
            ser.flush()
        except Exception as exc:
            self.raw_received.emit(f"write error: {exc}")
            return b""
        return self._read_raw(ser, timeout)

    def _decode(self, raw: bytes):
        frame = _extract_frame(bytearray(raw)) if raw else None
        return decode_weight(frame) if frame is not None else None

    def _crc_order(self):
        if self.preferred_crc is True:
            return (True, False)
        if self.preferred_crc is False:
            return (False, True)
        return (True, False)

    # ── стабилизация и отдача результата ──
    def _stabilize(self, ser, port: int, baud: int, addr: int, with_crc: bool):
        """Возвращает ('ok', вес, (port,baud,addr,crc)) | ('overload',...) | None."""
        self.progress.emit("Весы найдены, ожидание стабильности…")
        deadline = time.monotonic() + self.stabilize_timeout
        last = None
        while time.monotonic() < deadline:
            dec = self._decode(self._exchange(ser, addr, with_crc, 0.4))
            if dec is not None:
                value, stable, overload = dec
                self.raw_received.emit(
                    f"{port}@{baud} a{addr} crc={int(with_crc)}: "
                    f"{value:+.3f} стаб={stable} перегруз={overload}")
                if overload:
                    return ("overload", None, None)
                last = value
                if stable:
                    return ("ok", value, (port, baud, addr, with_crc))
            time.sleep(0.15)
        if last is not None:
            return ("ok", last, (port, baud, addr, with_crc))
        return None

    # ── поиск конфигурации ──
    def _baud_alive(self, ser) -> bool:
        """Быстрая проверка: приходит ли вообще что-то на этой скорости."""
        deadline = time.monotonic() + 0.2
        while time.monotonic() < deadline:           # пассивно
            if ser.read(64):
                return True
        probe = self._addrs[:2] if len(self._addrs) >= 2 else self._addrs
        for with_crc in self._crc_order():
            for addr in probe:
                if self._exchange(ser, addr, with_crc, 0.25):
                    return True
        return False

    def _find_config(self, ser):
        """Ищет (адрес, режим CRC), при котором приходит корректный вес."""
        for with_crc in self._crc_order():
            for addr in self._addrs:
                dec = self._decode(self._exchange(ser, addr, with_crc, 0.3))
                if dec is not None:
                    return (addr, with_crc)
        return None

    def _try_known(self, port: str, baud: int, addr: int, with_crc: bool):
        """Быстрая попытка по запомненной связке."""
        try:
            ser = self._open(port, baud)
        except Exception as exc:
            self.raw_received.emit(f"{port}@{baud}: {exc}")
            return None
        try:
            if self._decode(self._exchange(ser, addr, with_crc, 0.5)) is None:
                return None
            return self._stabilize(ser, port, baud, addr, with_crc)
        finally:
            ser.close()

    def _discover_on_port(self, port: str):
        for baud in self._bauds:
            self.progress.emit(f"Поиск весов: {port} @ {baud}")
            try:
                ser = self._open(port, baud)
            except Exception as exc:
                self.raw_received.emit(f"{port}@{baud}: {exc}")
                continue
            try:
                if not self._baud_alive(ser):
                    continue
                cfg = self._find_config(ser)
                if cfg is None:
                    continue
                addr, with_crc = cfg
                res = self._stabilize(ser, port, baud, addr, with_crc)
                if res is not None:
                    return res
            finally:
                ser.close()
        return None

    def _emit(self, res) -> bool:
        kind = res[0]
        if kind == "overload":
            self.error_occurred.emit(
                "Перегрузка весов!\n\n"
                "Снимите лишний груз с платформы и повторите взвешивание.")
            return True
        if kind == "ok":
            _, value, cfg = res
            port, baud, addr, with_crc = cfg
            self.port_found.emit(port)
            self.config_found.emit(port, int(baud), int(addr), bool(with_crc))
            self.weight_received.emit(float(value))
            return True
        return False

    # ── основной поток ──
    def run(self):
        if not SERIAL_AVAILABLE:
            self.error_occurred.emit(
                "Модуль pyserial не установлен.\n"
                "Выполните: python -m pip install pyserial")
            return

        ports = []
        if self.preferred_port:
            ports.append(self.preferred_port)
        for p in list_scale_ports():
            if p not in ports:
                ports.append(p)

        if not ports:
            self.error_occurred.emit(
                "Не найдено ни одного COM-порта.\n\n"
                "Проверьте:\n"
                "  • USB-кабель преобразователя/адаптера подключён\n"
                "  • Порт виден в Диспетчере устройств\n"
                "  • Установлен USB-драйвер «Тензо-М» (PreInstaller.exe)")
            return

        # 1) Мгновенная попытка по запомненной связке
        if (self.preferred_port and self.preferred_baud
                and self.preferred_address and self.preferred_crc is not None):
            self.progress.emit("Подключение к весам…")
            res = self._try_known(self.preferred_port, self.preferred_baud,
                                  self.preferred_address, self.preferred_crc)
            if res is not None and self._emit(res):
                return

        # 2) Полный автопоиск
        for port in ports:
            res = self._discover_on_port(port)
            if res is not None and self._emit(res):
                return

        self.error_occurred.emit(
            "Весы не отвечают.\n\n"
            "Проверьте:\n"
            "  • Преобразователь включён, кабель связи подключён\n"
            "  • Протокол обмена F4→D0 = «tEnSo» (не «6.43»)\n"
            "  • Подключён только один интерфейс (RS-232 / RS-485 / USB)\n"
            "  • Для RS-232 линии RXD/TXD перекрёстные (РЭ стр. 28)")