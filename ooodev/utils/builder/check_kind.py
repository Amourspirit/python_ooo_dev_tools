from __future__ import annotations
from enum import IntEnum


class CheckKind(IntEnum):
    """Check kind."""

    NONE = 0
    """NO CHECK"""
    SERVICE = 1
    """Check for service."""
    INTERFACE = 2
    """Check for interface."""
