from __future__ import annotations
from dataclasses import dataclass


@dataclass(frozen=True)
class Options:
    """
    Lo Load options

    .. versionadded:: 0.6.10
    """

    verbose: bool = False
    """Determines if various info is sent to console. Default ``False``"""
