import os
import threading
import time
from dataclasses import dataclass
from typing import Callable, Optional


@dataclass
class CardRead:
    uid: str
    holder_name: Optional[str]
    atr: Optional[str]


class CardReaderService:
    def __init__(
        self,
        on_card: Callable[[CardRead], None],
        on_remove: Optional[Callable[[], None]] = None,
    ) -> None:
        self._on_card = on_card
        self._on_remove = on_remove
        self._stop_event = threading.Event()
        self._thread: Optional[threading.Thread] = None
        self._last_uid: Optional[str] = None
        self._last_ts: float = 0.0
        self._ready = False
        self._init_error: Optional[str] = None

    @property
    def ready(self) -> bool:
        return self._ready

    @property
    def init_error(self) -> Optional[str]:
        return self._init_error

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        if self._thread:
            self._thread.join(timeout=3)

    def _run(self) -> None:
        try:
            from smartcard.CardMonitoring import CardMonitor, CardObserver
        except Exception as exc:
            self._init_error = f"pyscard import failed: {exc}"
            return

        service = self

        class Observer(CardObserver):
            def update(self, observable, actions):
                added_cards, removed_cards = actions
                for card in added_cards:
                    service._handle_card(card)
                if removed_cards and service._on_remove:
                    service._on_remove()

        monitor = CardMonitor()
        observer = Observer()
        monitor.addObserver(observer)
        self._ready = True

        while not self._stop_event.is_set():
            time.sleep(0.05)

        monitor.deleteObserver(observer)

    def _handle_card(self, card) -> None:
        uid, atr = self._read_uid_and_atr(card)
        if not uid:
            return
        now = time.time()
        if uid == self._last_uid and (now - self._last_ts) < 0.4:
            return
        self._last_uid = uid
        self._last_ts = now
        holder_name = self._read_holder_name(card)
        self._on_card(CardRead(uid=uid, holder_name=holder_name, atr=atr))

    def _read_uid_and_atr(self, card):
        try:
            from smartcard.util import toHexString
        except Exception:
            return None, None

        debug = os.environ.get("DEBUG_UID", "0") == "1"

        try:
            connection = card.createConnection()
            connection.connect()
            atr = toHexString(connection.getATR())

            for attempt in range(2):
                for length in (0x00, 0x04, 0x07, 0x0A):
                    data, sw1, sw2 = connection.transmit([0xFF, 0xCA, 0x00, 0x00, length])
                    if (sw1, sw2) == (0x90, 0x00) and data:
                        uid = "".join(f"{byte:02X}" for byte in data)
                        if debug:
                            print(f"UID APDU len={length} SW={sw1:02X}{sw2:02X} DATA={uid} ATR={atr}")
                        return uid, atr
                    if debug:
                        payload = "".join(f"{byte:02X}" for byte in data) if data else ""
                        print(
                            f"UID APDU len={length} SW={sw1:02X}{sw2:02X} DATA={payload} ATR={atr}"
                        )
                time.sleep(0.03)
            return None, atr
        except Exception:
            return None, None

    def _read_holder_name(self, card) -> Optional[str]:
        # Placeholder: the APDU to read the holder name depends on card type/app.
        return None
