"""
This module is DEPRECATED since version 0.13 It is no longer recommended for use and may be removed in the future.

Use :ref:`ooodev.dialog.Dialogs` instead.
"""

# Use the following instead:
# from ooodev.dialog import Dialogs

from ooodev.dialog.dialogs import Dialogs as Dialogs
import warnings

warnings.warn(
    "ooodev.utils.dialogs module is deprecated. Use ooodev.dialog instead.", DeprecationWarning, stacklevel=2
)
