"""
This module is DEPRECATED since version ``0.26.0``.
It is no longer recommended for use and may be removed in the future.

Use :py:mod:`ooodev.loader.inst.doc_type` instead.
"""

from ooodev.loader.inst.doc_type import DocType as DocType
from ooodev.loader.inst.doc_type import DocTypeStr as DocTypeStr
import warnings

warnings.warn(
    "ooodev.inst.lo.doc_type module is deprecated. Use ooodev.loader.inst.doc_type instead.",
    DeprecationWarning,
    stacklevel=2,
)
