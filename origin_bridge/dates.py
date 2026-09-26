"""Consistent ISO timestamp parsing across supported Python versions."""
import re
from datetime import datetime


def parse_iso_datetime(value: str) -> datetime:
    token = value.strip()
    if token.endswith(("Z", "z")):
        token = token[:-1] + "+00:00"
    # Python 3.10 only accepts 3 or 6 fractional digits. Keep the original
    # source strings in DataColumn; export preflight reports precision limits.
    token = re.sub(
        r"(\d{2}:\d{2}:\d{2})[.,](\d+)",
        lambda match: match[1] + "." + match[2][:6].ljust(6, "0"),
        token,
    )
    return datetime.fromisoformat(token)
