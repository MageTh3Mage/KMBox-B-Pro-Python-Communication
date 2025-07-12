import serial
from serial.tools.list_ports import comports

class Kmbox:
    __slots__ = ('kmbox', 'port', 'debug', 'write_buffer')

    def __init__(self, port=None, baudrate=115200, timeout=0.05, debug=False):
        self.kmbox = None
        self.port = None
        self.debug = debug
        self.write_buffer = bytearray()

        if not hasattr(serial, 'Serial'):
            raise ImportError("KMBox requires pyserial. Run: pip install pyserial")

        if debug:
            print("KMBox Library - By Mage")

        self._connect(port, baudrate, timeout)

    def _connect(self, port, baudrate, timeout):
        self.port = port or self._find_port()
        if not self.port:
            if self.debug:
                print("[ERROR] No compatible CH340 device found")
            return

        try:
            self.kmbox = serial.Serial(self.port, baudrate, timeout=timeout, write_timeout=timeout)
            if self.debug:
                print(f"[INFO] Connected to {self.port}")
        except serial.SerialException as e:
            if self.debug:
                print(f"[ERROR] Connection failed: {e}")

    def _find_port(self):
        for port in comports():
            desc = (port.description or "").upper()
            hwid = (port.hwid or "").upper()
            if 'CH340' in desc or 'CH340' in hwid or 'USB-SERIAL' in desc + hwid:
                if self.debug:
                    print(f"[INFO] Found device on {port.device}")
                try:
                    with serial.Serial(port.device) as s:
                        return port.device
                except serial.SerialException:
                    continue
        return None

    def is_connected(self):
        return self.kmbox is not None and self.kmbox.is_open

    def _send(self, cmd: str):
        try:
            self.write_buffer.clear()
            self.write_buffer.extend(cmd.encode())
            self.write_buffer.append(10)  # \n = 10
            self.kmbox.write(self.write_buffer)
            self.kmbox.flush()
        except Exception as e:
            if self.debug:
                print(f"[ERROR] Send failed: {e}")

    def move(self, x: int, y: int):
        self._send(f"km.move({x},{y})")

    def left_click(self):
        self._send("km.click(0)")

    def right_click(self):
        self._send("km.click(1)")

    def middle_click(self):
        self._send("km.click(2)")

    def close(self):
        if self.is_connected():
            self.kmbox.close()
            if self.debug:
                print("[INFO] Connection closed")