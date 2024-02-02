"""
This module is DEPRECATED since version ``0.26.0``.
It is no longer recommended for use and may be removed in the future.

Use :py:mod:`ooodev.loader.inst.lo_loader` instead.
"""

from ooodev.loader.inst.lo_loader import LoLoader as LoLoader
import warnings

warnings.warn(
    "ooodev.inst.lo.lo_loader module is deprecated. Use ooodev.loader.inst.lo_loader instead.",
    DeprecationWarning,
    stacklevel=2,
)
