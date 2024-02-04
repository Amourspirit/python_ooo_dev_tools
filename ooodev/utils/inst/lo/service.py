"""
This module is DEPRECATED since version ``0.26.0``.
It is no longer recommended for use and may be removed in the future.

Use :py:mod:`ooodev.loader.inst.options` instead.
"""

from ooodev.loader.inst.service import Service as Service
import warnings

warnings.warn(
    "ooodev.inst.lo.options module is deprecated. Use ooodev.loader.inst.options instead.",
    DeprecationWarning,
    stacklevel=2,
)
