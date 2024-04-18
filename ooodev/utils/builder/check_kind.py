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
    SERVICE_ALL = 3
    """Check that all services are matched."""
    INTERFACE_ALL = 4
    """Check that all interfaces are matched."""
    SERVICE_ONLY = 5
    """Check that only the first service in a services sequence is matched."""
    INTERFACE_ONLY = 6
    """Check that only the first interface in an interfaces sequence is matched."""
