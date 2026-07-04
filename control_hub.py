from __future__ import annotations

from dataclasses import dataclass, asdict
from datetime import datetime
from threading import Lock
from typing import Optional, Dict, Any, List


@dataclass
class BotStatus:
    running: bool = False
    paused: bool = True
    stopped: bool = False
    current_excel_row: Optional[int] = None
    current_index: int = 0
    total_rows: int = 0
    current_hm: str = ""
    current_vehicle: str = ""
    current_image_index: int = 0
    total_images: int = 0
    message: str = "Idle"
    last_error: str = ""
    current_image_path: str = ""
    updated_at: str = ""


class ControlHub:
    """Thread-safe bridge between FastAPI buttons and Playwright worker."""

    def __init__(self):
        self._lock = Lock()
        self.status = BotStatus(updated_at=datetime.now().isoformat(timespec="seconds"))
        self._commands: List[Dict[str, Any]] = []

    def push_command(self, name: str, payload: Optional[Dict[str, Any]] = None) -> None:
        payload = payload or {}
        name = name.upper().strip()
        with self._lock:
            if name == "PAUSE":
                self.status.paused = True
            elif name in {"START", "RESUME"}:
                self.status.paused = False
                self.status.stopped = False
            elif name == "STOP":
                self.status.stopped = True
                self.status.paused = False
            self._commands.append({"name": name, "payload": payload, "ts": datetime.now().isoformat(timespec="seconds")})
            self.status.updated_at = datetime.now().isoformat(timespec="seconds")

    def pop_command(self, *names: str) -> Optional[Dict[str, Any]]:
        wanted = {n.upper() for n in names}
        with self._lock:
            for i, cmd in enumerate(self._commands):
                if not wanted or cmd["name"] in wanted:
                    return self._commands.pop(i)
        return None

    def drain_commands(self) -> List[Dict[str, Any]]:
        with self._lock:
            cmds = list(self._commands)
            self._commands.clear()
            return cmds

    def has_command(self, *names: str) -> bool:
        wanted = {n.upper() for n in names}
        with self._lock:
            return any(cmd["name"] in wanted for cmd in self._commands)

    def update_status(self, **kwargs: Any) -> None:
        with self._lock:
            for k, v in kwargs.items():
                if hasattr(self.status, k):
                    setattr(self.status, k, v)
            self.status.updated_at = datetime.now().isoformat(timespec="seconds")

    def get_status(self) -> Dict[str, Any]:
        with self._lock:
            return asdict(self.status)

    def should_stop(self) -> bool:
        with self._lock:
            return self.status.stopped

    def is_paused(self) -> bool:
        with self._lock:
            return self.status.paused

    def clear_runtime_flags_for_next_row(self) -> None:
        # Keep PAUSE/STOP status. Remove old one-shot navigation commands.
        with self._lock:
            self._commands = [cmd for cmd in self._commands if cmd["name"] in {"PAUSE", "STOP"}]
            self.status.current_image_index = 0
            self.status.total_images = 0
            self.status.current_image_path = ""
            self.status.updated_at = datetime.now().isoformat(timespec="seconds")
