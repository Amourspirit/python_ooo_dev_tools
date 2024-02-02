"""
This module is DEPRECATED since version ``0.26.0``.
It is no longer recommended for use and may be removed in the future.

Use :py:class:`ooodev.loader.Lo` instead.
"""

from ooodev.loader.lo import Lo as Lo
import warnings

warnings.warn(
    "ooodev.inst.lo.Lo class is deprecated. Use ooodev.loader.Lo instead.",
    DeprecationWarning,
    stacklevel=2,
)
