from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Any, Dict

ENV_PATTERN = re.compile(r"^\$\{([A-Z0-9_]+)\}$")


def _resolve_env(value: Any) -> Any:
    if isinstance(value, str):
        m = ENV_PATTERN.match(value.strip())
        if m:
            return os.getenv(m.group(1), "")
    if isinstance(value, dict):
        return {k: _resolve_env(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_resolve_env(v) for v in value]
    return value


def load_config(config_path: str | None = None) -> Dict[str, Any]:
    path = Path(config_path or os.getenv("EFLEET_CONFIG", "config.example.json"))
    if not path.exists():
        raise FileNotFoundError(f"Config file not found: {path}")
    with path.open("r", encoding="utf-8") as f:
        cfg = json.load(f)
    cfg = _resolve_env(cfg)
    cfg.setdefault("excel_file", "data/VehicleDetails.xlsx")
    cfg.setdefault("sheet_name", "DATA")
    cfg.setdefault("headless", True)
    cfg.setdefault("auto_start", False)
    cfg.setdefault("image_auto_close_sec", 2.0)
    cfg.setdefault("max_images", 20)
    cfg.setdefault("start_excel_row", 2)
    cfg.setdefault("save_after_each_query", True)
    cfg.setdefault("wait_for_save_next_after_images", True)
    return cfg
