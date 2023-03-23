from __future__ import annotations
from typing import NamedTuple


class FrameOptionsProperties(NamedTuple):
    editable: str  # EditInReadonly
    printable: str  # Print
    write_mode: str  # WritingMode
