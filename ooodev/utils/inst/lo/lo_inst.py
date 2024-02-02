"""
This module is DEPRECATED since version ``0.26.0``.
It is no longer recommended for use and may be removed in the future.

Use :ref:`ooodev.loader.inst.LoInst` instead.
"""

from ooodev.loader.inst.lo_inst import LoInst as LoInst
import warnings

warnings.warn(
    "ooodev.utils.inst.lo.lo_inst  module is deprecated. Use ooodev.loader.inst.lo_inst instead.",
    DeprecationWarning,
    stacklevel=2,
)
