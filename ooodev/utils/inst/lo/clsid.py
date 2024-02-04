"""
This module is DEPRECATED since version ``0.26.0``.
It is no longer recommended for use and may be removed in the future.

Use :py:mod:`ooodev.loader.inst.clsid` instead.
"""

from ooodev.loader.inst.clsid import CLSID as CLSID
import warnings

warnings.warn(
    "ooodev.inst.lo.clsid module is deprecated. Use ooodev.loader.inst.clsid instead.",
    DeprecationWarning,
    stacklevel=2,
)
