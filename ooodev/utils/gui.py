"""
This module is DEPRECATED since version ``0.40.0``.
It is no longer recommended for use and may be removed in the future.

Use :py:class:`ooodev.gui.GUI` instead.
"""

from ooodev.gui.gui import GUI as GUI
import warnings

warnings.warn(
    "ooodev.util.gui class is deprecated. Use from ooodev.gui import GUI",
    DeprecationWarning,
    stacklevel=2,
)
