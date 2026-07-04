from __future__ import annotations

from typing import Dict, List, Tuple, Optional

# Keyboard-style short keys are kept because user is already used to them.
COLUMN_MAP: Dict[str, str] = {
    "c": "CN",
    "r": "RCM",
    "f": "FCM",
    "e": "E- Way bill",
    "h": "Challan",
    "t": "Tax invoice",
    "w": "weight slip",
    "k": "Tracking details are not available",
    "z": "RC",
    "v": "Vehicle photo",
    "p": "Permit certificate",
    "x": "Tax document",
    "i": "Insurance",
    "n": "Fitness",
    "u": "PUC",
    "d": "DL",
}

AUDIT_COLUMNS: List[str] = list(COLUMN_MAP.values())

DEFAULT_VALUES: Dict[str, str] = {col: "ok" for col in AUDIT_COLUMNS}
DEFAULT_VALUES.update({"RCM": "Yes", "FCM": "Yes"})

QUERY_MASTER: Dict[str, Dict[str, str]] = {
    "CN": {
        "1": "Vehicle no. is not mention in bilty",
        "2": "E-way bill no is not mention in bilty",
        "3": "CN document is not found in efleet",
    },
    "RCM": {
        "1": "No",
        "2": "wrong bilty/CN i.e. it should be FCM",
    },
    "FCM": {
        "1": "No",
        "2": "This party is not into FCM List",
    },
    "E- Way bill": {
        "1": "E-Way bill is not Available in efleet",
        "2": "Wrong e-way bill uploaded in efleet",
        "3": "E-way bill expired before delivery date",
    },
    "Challan": {
        "1": "Driver's signature missing on challan",
        "2": "Challan document is not found in efleet",
        "3": "Wrong challan uploaded in efleet",
    },
    "Tax invoice": {
        "1": "Tax invoice is not found in efleet",
        "2": "Wrong Tax invoice uploaded in efleet",
    },
    "weight slip": {
        "1": "weight slip is not found",
        "2": "Wrong weight slip uploaded in efleet",
    },
    "Tracking details are not available": {
        "1": "Tracking details are not available",
        "2": "Tracking screenshot is not found in efleet",
    },
    "RC": {
        "1": "RC document is not found",
        "2": "Wrong RC uploaded in efleet",
    },
    "Vehicle photo": {
        "1": "Vehicle photo is not uploaded",
        "2": "Wrong vehicle photo uploaded in efleet",
    },
    "Permit certificate": {
        "1": "Permit certificate is not found",
        "2": "Permit certificate expired",
    },
    "Tax document": {
        "1": "Tax document is not found",
        "2": "Tax document expired",
    },
    "Insurance": {
        "1": "Insurance Expired",
        "2": "Insurance Expired before delivery date",
        "3": "Insurance document is not found",
    },
    "Fitness": {
        "1": "Fitness Expired",
        "2": "Fitness document is not found",
    },
    "PUC": {
        "1": "PUC Expired",
        "2": "PUC is not found",
    },
    "DL": {
        "1": "DL is not found",
        "2": "DL expired",
        "3": "Wrong DL uploaded in efleet",
    },
}

OK_VALUES = {"", "ok", "yes", "nan", "none", "#n/a"}


def normalize(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def is_issue_value(value: object) -> bool:
    text = normalize(value)
    return text.lower() not in OK_VALUES


def resolve_query(column: str, query_no: str, custom_text: Optional[str] = None) -> str:
    if custom_text and custom_text.strip():
        return custom_text.strip()
    column = normalize(column)
    query_no = normalize(query_no)
    if column not in QUERY_MASTER:
        raise KeyError(f"Unknown query column: {column}")
    if query_no not in QUERY_MASTER[column]:
        raise KeyError(f"Query no {query_no} not found for {column}")
    return QUERY_MASTER[column][query_no]


def build_audit_queries(row_values: Dict[str, object]) -> str:
    issues: List[str] = []
    seen = set()
    for col in AUDIT_COLUMNS:
        value = normalize(row_values.get(col, ""))
        if is_issue_value(value) and value not in seen:
            issues.append(value)
            seen.add(value)
    return ", ".join(issues)


def query_options_for_ui() -> List[Dict[str, str]]:
    options: List[Dict[str, str]] = []
    for key, col in COLUMN_MAP.items():
        entries = QUERY_MASTER.get(col, {})
        for no, text in entries.items():
            options.append({"key": key, "column": col, "query_no": no, "text": text})
    return options


def default_row_values() -> Dict[str, str]:
    return dict(DEFAULT_VALUES)
