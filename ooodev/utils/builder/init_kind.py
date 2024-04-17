from __future__ import annotations
from enum import IntEnum


class InitKind(IntEnum):
    """Import kind."""

    NONE = 0
    """Class constructor has no parameters."""
    COMPONENT = 1
    """Class constructor has only a component parameter."""
    COMPONENT_INTERFACE = 2
    """Class constructor has a component parameter and an interface parameter."""
    CALLBACK = 3
    """Args are gotten from a callback."""
    LO_INST = 4
    """Class constructor has a LoInst parameter."""
