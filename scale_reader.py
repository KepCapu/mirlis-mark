# scale_reader.py
# Чтение веса с весов через RS-232 / RS-485 / USB-Serial.
# Движок автоопределения: перебор портов и скоростей + пассивное прослушивание
# ASCII-потока (Фаза A) + активный опрос по реестру драйверов (Фаза B:
# tEnSo «Тензо-М», generic-ASCII/MT-SICS, 6.43). Найденная комбинация кэшируется.
#
# Ключевые принципы (важно для надёжности и скорости):
#   • tEnSo С CRC — принимаем по ПЕРВОМУ валидному кадру: контрольная сумма уже
#     гарантирует достоверность, повторы не нужны (иначе из-за тайминга на
#     виртуальных/медленных портах ответ теряется и поиск идёт бесконечно).
#   • без CRC и ASCII — требуем краткое подтверждение (значение дважды).
#   • весь поиск ограничен по времени (общий дедлайн), чтобы не «зависать».
#
# Публичный интерфейс ПОЛНОСТЬЮ совместим со старым main.py:
#   класс ScaleReader(QThread) с сигналами weight_received(float),
#   raw_received(str), error_occurred(str), progress(str), port_found(str);
#   конструктор ScaleReader(preferred_port, baud, per_port_timeout,
#   saved_port_timeout, parent); флаг SERIAL_AVAILABLE.

from __future__ import annotations

import os
import re
import json
import time
from typing import Optional, List, Tuple, Dict, Any

from PyQt5.QtCore import QThread, pyqtSignal

try:
    import serial
    from serial.tools import list_ports
    SERIAL_AVAILABLE = True
except ImportError:
    SERIAL_AVAILABLE = False


# ─────────────────────────────────────────────────────────────────────────────
#  Параметры перебора
# ─────────────────────────────────────────────────────────────────────────────

_PRIORITY_VIDS = {0x1A86, 0x0403, 0x067B, 0x10C4, 0x2341}  # CH340/FTDI/PL2303/CP210x/Arduino
_BAUDS = [9600, 4800, 19200, 38400, 2400, 57600, 115200]   # частые — первыми
_DEFAULT_ADDRESSES = [1, 2, 3, 4, 5, 6, 7, 8]

# Тайминги одного шага (короткие — приём подтверждается контентом, не ожиданием)
_SETTLE_S = 0.03          # пауза после отправки запроса перед чтением
_READ_S   = 0.20          # окно чтения ответа на активный запрос
_PASSIVE_S = 0.30         # прослушивание ASCII-потока на каждой скорости

# Общий лимит времени всего поиска (защита от «зависания»)
_TOTAL_BUDGET_S = 60.0

# Правдоподобный диапазон веса (кг)
_MIN_KG, _MAX_KG = -1000.0, 5000.0


# ─────────────────────────────────────────────────────────────────────────────
#  Кэш найденной комбинации (свой файл, main.py не трогаем)
# ─────────────────────────────────────────────────────────────────────────────

def _cache_path() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    folder = os.path.join(base, "MirlisMark")
    try:
        os.makedirs(folder, exist_ok=True)
    except Exception:
        folder = base
    return os.path.join(folder, "scale_link.json")


def _load_link_cache() -> Optional[Dict[str, Any]]:
    try:
        with open(_cache_path(), "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict) and data.get("port"):
            return data
    except Exception:
        pass
    return None


def _save_link_cache(combo: Dict[str, Any]) -> None:
    try:
        with open(_cache_path(), "w", encoding="utf-8") as f:
            json.dump(combo, f, ensure_ascii=False)
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────────────────
#  Перечисление портов
# ─────────────────────────────────────────────────────────────────────────────

def list_scale_ports() -> list:
    if not SERIAL_AVAILABLE:
        return []
    try:
        ports = list(list_ports.comports())
    except Exception:
        return []
    ports.sort(key=lambda p: 0 if (getattr(p, "vid", None) or 0) in _PRIORITY_VIDS else 1)
    result = []
    for p in ports:
        if p.device and p.device not in result:
            result.append(p.device)
    return result


# ─────────────────────────────────────────────────────────────────────────────
#  Универсальный парсер ASCII-веса
# ─────────────────────────────────────────────────────────────────────────────

_NUM_RE = re.compile(r"[+-]?\d{1,7}(?:[.,]\d{1,4})?")


def parse_ascii_weight(line: str) -> Optional[float]:
    """ASCII-строка весов → вес в КГ или None."""
    if not line:
        return None
    up = line.upper()
    unit = None
    if re.search(r"KG|КГ", up):
        unit = "kg"
    elif re.search(r"\bLB\b", up):
        unit = "lb"
    elif re.search(r"(?<![A-ZА-Я])T(?![A-ZА-Я])|ТОНН", up):
        unit = "t"
    elif re.search(r"(?<![A-ZА-Я])G(?![A-ZА-Я])|ГР|(?<![А-Я])Г(?![А-Я])", up):
        unit = "g"

    nums = _NUM_RE.findall(line)
    if not nums:
        return None
    chosen = None
    for tok in nums:
        if "." in tok or "," in tok:
            chosen = tok
    if chosen is None:
        chosen = nums[-1]
    try:
        val = float(chosen.replace(",", "."))
    except Exception:
        return None

    if unit == "g":
        val /= 1000.0
    elif unit == "t":
        val *= 1000.0
    elif unit == "lb":
        val *= 0.45359237
    elif unit is None:
        if ("." not in chosen and "," not in chosen) and abs(val) > 200:
            val /= 1000.0

    if not (_MIN_KG <= val <= _MAX_KG):
        return None
    return round(val, 3)


def parse_weight_line(line: str) -> Optional[float]:   # обратная совместимость
    return parse_ascii_weight(line)


# ─────────────────────────────────────────────────────────────────────────────
#  CRC и кадры протокола «Тензо-М» (tEnSo)
# ─────────────────────────────────────────────────────────────────────────────

def crc_tenso(data: bytes) -> int:
    """CRC-8 «Тензо-М»: полином 0x69, MSB-first, init 0x00 (vks.pdf)."""
    crc = 0
    for byte in data:
        d = byte
        for _ in range(8):
            if crc & 0x80:
                crc = (crc << 1) & 0xFF
                if d & 0x80:
                    crc |= 1
                crc ^= 0x69
            else:
                crc = (crc << 1) & 0xFF
                if d & 0x80:
                    crc |= 1
            d = (d << 1) & 0xFF
    return crc


def tenso_frame(addr: int, cop: int, with_crc: bool = True) -> bytes:
    body = [addr & 0xFF, cop & 0xFF]
    if with_crc:
        body.append(crc_tenso(bytes(body + [0x00])))
    stuffed = []
    for b in body:
        stuffed.append(b)
        if b == 0xFF:
            stuffed.append(0xFE)
    return bytes([0xFF] + stuffed + [0xFF, 0xFF])


def _tenso_extract_frames(raw: bytes) -> List[bytes]:
    frames: List[bytes] = []
    cur: List[int] = []
    i, n = 0, len(raw)
    while i < n:
        b = raw[i]
        if b == 0xFF:
            if i + 1 < n and raw[i + 1] == 0xFE:
                cur.append(0xFF)
                i += 2
                continue
            if cur:
                frames.append(bytes(cur)); cur = []
            i += 1
        else:
            cur.append(b); i += 1
    if cur:
        frames.append(bytes(cur))
    return frames


def _bcd(b: int) -> Optional[int]:
    hi, lo = (b >> 4), (b & 0x0F)
    if hi > 9 or lo > 9:
        return None
    return hi * 10 + lo


def tenso_parse_weight(raw: bytes, require_crc: bool = False) -> Optional[float]:
    """Ищет в ответе кадр веса (COP C2/C3) → кг. При require_crc=True принимает
    только кадр с корректной контрольной суммой (для приёма «по первому кадру»)."""
    for fr in _tenso_extract_frames(raw):
        variants = (True,) if require_crc else (True, False)
        for has_crc in variants:
            need = 7 if has_crc else 6
            if len(fr) < need:
                continue
            f = fr[:need]
            if f[1] not in (0xC2, 0xC3):
                continue
            if has_crc and crc_tenso(f) != 0:
                continue
            d0, d1, d2 = _bcd(f[2]), _bcd(f[3]), _bcd(f[4])
            if None in (d0, d1, d2):
                continue
            con = f[5]
            kg = (d2 * 10000 + d1 * 100 + d0) / (10 ** (con & 0x07))
            if con & 0x80:
                kg = -kg
            if _MIN_KG <= kg <= _MAX_KG:
                return round(kg, 3)
    if not require_crc:
        try:
            return parse_ascii_weight(raw.decode("ascii", errors="ignore"))
        except Exception:
            return None
    return None


# ─────────────────────────────────────────────────────────────────────────────
#  Ввод-вывод
# ─────────────────────────────────────────────────────────────────────────────

class _IO:
    def __init__(self, ser, emit_raw, port: str, baud: int, deadline: float):
        self._ser = ser
        self._emit = emit_raw
        self.port = port
        self.baud = baud
        self.deadline = deadline

    def expired(self) -> bool:
        return time.monotonic() >= self.deadline

    def flush(self):
        try:
            self._ser.reset_input_buffer()
            self._ser.reset_output_buffer()
        except Exception:
            pass

    def write(self, data: bytes):
        try:
            self._ser.write(data)
            self._ser.flush()
        except Exception:
            pass

    def read_for(self, seconds: float) -> bytes:
        out = bytearray()
        end = time.monotonic() + seconds
        while time.monotonic() < end:
            try:
                chunk = self._ser.read(256)
            except Exception:
                break
            if chunk:
                out.extend(chunk)
                end = time.monotonic() + 0.10
            else:
                time.sleep(0.01)
        return bytes(out)

    def ask(self, frame: bytes, read_s: float = _READ_S) -> bytes:
        """Отправить запрос и прочитать ответ (с короткой паузой на разворот)."""
        self.flush()
        self.write(frame)
        time.sleep(_SETTLE_S)
        return self.read_for(read_s)

    def log(self, msg: str):
        if self._emit:
            self._emit(f"{self.port}@{self.baud}: {msg}")


def _confirm2(a: Optional[float], b: Optional[float]) -> Optional[float]:
    if a is not None and a == b:
        return a
    return None


# ─────────────────────────────────────────────────────────────────────────────
#  Драйверы (Фаза B)
# ─────────────────────────────────────────────────────────────────────────────

class _Driver:
    name = "base"
    def poll(self, io: "_IO", addresses: List[int]) -> Optional[Tuple[float, Dict[str, Any]]]:
        raise NotImplementedError


class _TensoDriver(_Driver):
    """«Тензо-М» (tEnSo). С CRC — приём по первому валидному кадру; без CRC —
    с подтверждением (значение дважды)."""
    name = "tenso"

    def poll(self, io, addresses):
        # 1) С CRC: принимаем по первому корректному кадру.
        for cop in (0xC2, 0xC3):
            for addr in addresses:
                if io.expired():
                    return None
                raw = io.ask(tenso_frame(addr, cop, True))
                v = tenso_parse_weight(raw, require_crc=True)
                if v is not None:
                    io.log(f"tEnSo адрес={addr} COP={cop:#x} CRC → {v} кг")
                    return v, {"driver": self.name, "address": addr, "cop": cop, "crc": True}
        # 2) Без CRC: подтверждение двумя одинаковыми ответами.
        for cop in (0xC2, 0xC3):
            for addr in addresses:
                if io.expired():
                    return None
                v1 = tenso_parse_weight(io.ask(tenso_frame(addr, cop, False)))
                if v1 is None:
                    continue
                v2 = tenso_parse_weight(io.ask(tenso_frame(addr, cop, False)))
                v = _confirm2(v1, v2)
                if v is not None:
                    io.log(f"tEnSo адрес={addr} COP={cop:#x} без CRC → {v} кг")
                    return v, {"driver": self.name, "address": addr, "cop": cop, "crc": False}
        return None


class _AsciiPollDriver(_Driver):
    """Опрос ASCII-весов «запрос-ответ» (в т.ч. MT-SICS)."""
    name = "ascii-poll"
    _TRIGGERS = [b"SI\r\n", b"\x05", b"W\r\n"]

    def _read_value(self, io) -> Optional[float]:
        for trig in self._TRIGGERS:
            if io.expired():
                return None
            raw = io.ask(trig, read_s=0.3)
            for line in re.split(rb"[\r\n]+", raw):
                v = parse_ascii_weight(line.decode("ascii", errors="ignore"))
                if v is not None:
                    return v
        return None

    def poll(self, io, addresses):
        v1 = self._read_value(io)
        if v1 is None:
            return None
        v = _confirm2(v1, self._read_value(io))
        if v is not None:
            io.log(f"ASCII-опрос → {v} кг")
            return v, {"driver": self.name}
        return None


class _Protocol643Driver(_Driver):
    """Протокол 6.43: команда 0x10 (ответ ASCII). Адрес 0 отвечает всегда;
    1..4 — активация 0x01 + 4 ASCII-цифры."""
    name = "643"

    def _indicator(self, io) -> Optional[float]:
        raw = io.ask(b"\x10", read_s=0.3)
        if not raw:
            return None
        s = raw.decode("ascii", errors="ignore")
        m = re.search(r"=([^\r\n]{1,12})", s)
        return parse_ascii_weight(m.group(1) if m else s)

    def poll(self, io, addresses):
        v = _confirm2(self._indicator(io), self._indicator(io))   # адрес 0
        if v is not None:
            io.log(f"6.43 адрес=0 → {v} кг")
            return v, {"driver": self.name, "address": 0}
        for addr in addresses[:4]:
            if io.expired():
                return None
            io.flush()
            io.write(b"\x01" + f"{addr:04d}".encode("ascii"))
            time.sleep(0.03)
            v = _confirm2(self._indicator(io), self._indicator(io))
            io.write(b"\x02")
            if v is not None:
                io.log(f"6.43 адрес={addr} → {v} кг")
                return v, {"driver": self.name, "address": addr}
        return None


def _build_drivers() -> List[_Driver]:
    return [_TensoDriver(), _AsciiPollDriver(), _Protocol643Driver()]


# ─────────────────────────────────────────────────────────────────────────────
#  Поток
# ─────────────────────────────────────────────────────────────────────────────

class ScaleReader(QThread):
    weight_received = pyqtSignal(float)
    raw_received    = pyqtSignal(str)
    error_occurred  = pyqtSignal(str)
    progress        = pyqtSignal(str)
    port_found      = pyqtSignal(str)

    def __init__(self, preferred_port: str = "", baud: int = 9600,
                 per_port_timeout: float = 1.5, saved_port_timeout: float = 2.5,
                 parent=None, addresses: Optional[List[int]] = None,
                 total_budget: float = _TOTAL_BUDGET_S, **_ignored):
        super().__init__(parent)
        self.preferred_port     = preferred_port or ""
        self.baud               = int(baud or 9600)
        self.per_port_timeout   = per_port_timeout
        self.saved_port_timeout = saved_port_timeout
        self.addresses          = addresses or _DEFAULT_ADDRESSES
        self.total_budget       = total_budget
        self._abort             = False
        self._deadline          = 0.0
        self._drivers           = _build_drivers()

    def stop(self):
        self._abort = True

    def _baud_order(self) -> List[int]:
        return [self.baud] + [b for b in _BAUDS if b != self.baud]

    def _open(self, port: str, baud: int):
        return serial.Serial(
            port=port, baudrate=baud,
            bytesize=serial.EIGHTBITS, parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE, timeout=0.15, write_timeout=0.5,
        )

    def _passive(self, io: "_IO", listen_s: float) -> Optional[float]:
        raw = io.read_for(listen_s)
        if not raw:
            return None
        v1 = v2 = None
        for line in re.split(rb"[\r\n]+", raw):
            s = line.decode("ascii", errors="ignore").strip()
            if not s:
                continue
            v = parse_ascii_weight(s)
            if v is None:
                continue
            if v1 is None:
                v1 = v
            elif v == v1:
                v2 = v
                break
        v = _confirm2(v1, v2)
        if v is not None:
            io.log(f"поток ASCII → {v} кг")
        return v

    def _scan(self, port: str, baud: int, listen_s: float
              ) -> Optional[Tuple[float, Dict[str, Any]]]:
        try:
            with self._open(port, baud) as ser:
                io = _IO(ser, self.raw_received.emit, port, baud, self._deadline)
                v = self._passive(io, listen_s)
                if v is not None:
                    return v, {"driver": "ascii-stream"}
                for drv in self._drivers:
                    if self._abort or io.expired():
                        return None
                    res = drv.poll(io, self.addresses)
                    if res is not None:
                        return res
        except Exception as exc:
            self.raw_received.emit(f"{port}@{baud}: {exc}")
        return None

    def _try_cached(self, cached: Dict[str, Any]) -> Optional[Tuple[float, Dict[str, Any]]]:
        port = cached.get("port")
        baud = int(cached.get("baud") or self.baud)
        if not port:
            return None
        if SERIAL_AVAILABLE and port not in list_scale_ports() and port != self.preferred_port:
            return None
        self.progress.emit(f"Весы: {port} @ {baud}")
        try:
            with self._open(port, baud) as ser:
                io = _IO(ser, self.raw_received.emit, port, baud, self._deadline)
                drv = cached.get("driver")
                # Точное восстановление выигравшего опроса — мгновенно
                if drv == "tenso":
                    addr = int(cached.get("address", 1))
                    cop = int(cached.get("cop", 0xC2))
                    crc = bool(cached.get("crc", True))
                    raw = io.ask(tenso_frame(addr, cop, crc))
                    v = tenso_parse_weight(raw, require_crc=crc)
                    if v is None and not crc:
                        v = _confirm2(v, tenso_parse_weight(io.ask(tenso_frame(addr, cop, crc))))
                    if v is not None:
                        return v, cached
                elif drv == "ascii-stream":
                    v = self._passive(io, 0.8)
                    if v is not None:
                        return v, cached
                else:
                    for d in self._drivers:
                        res = d.poll(io, self.addresses)
                        if res is not None:
                            return res
        except Exception as exc:
            self.raw_received.emit(f"{port}@{baud}: {exc}")
        return None

    def _accept(self, port: str, baud: int, value: float, combo: Dict[str, Any]):
        combo = dict(combo); combo["port"] = port; combo["baud"] = baud
        _save_link_cache(combo)
        self.port_found.emit(port)
        self.weight_received.emit(value)

    def run(self):
        if not SERIAL_AVAILABLE:
            self.error_occurred.emit(
                "Модуль pyserial не установлен.\n"
                "Выполните: python -m pip install pyserial"
            )
            return

        self._deadline = time.monotonic() + self.total_budget

        cached = _load_link_cache()
        if cached:
            res = self._try_cached(cached)
            if res is not None:
                v, combo = res
                self._accept(cached["port"], int(cached.get("baud") or self.baud), v, combo)
                return

        ports_order: List[str] = []
        if self.preferred_port:
            ports_order.append(self.preferred_port)
        for p in list_scale_ports():
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

        for port in ports_order:
            if self._abort or time.monotonic() >= self._deadline:
                break
            listen_s = _PASSIVE_S * (2 if port == self.preferred_port else 1)
            for baud in self._baud_order():
                if self._abort or time.monotonic() >= self._deadline:
                    break
                self.progress.emit(f"Поиск весов: {port} @ {baud}")
                res = self._scan(port, baud, listen_s=listen_s)
                if res is not None:
                    v, combo = res
                    self._accept(port, baud, v, combo)
                    return

        self.error_occurred.emit(
            "Весы не найдены.\n\n"
            "Проверьте:\n"
            "  • USB-кабель адаптера подключён\n"
            "  • Кабель от весов идёт на правильный интерфейс (RS-232 или RS-485)\n"
            "  • Для RS-485 — не перепутаны линии A и B (попробуйте поменять местами)\n"
            "  • На весы положен груз; в меню весов включена выдача по каналу"
        )