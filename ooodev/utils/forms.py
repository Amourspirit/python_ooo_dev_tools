"""
This module is DEPRECATED since version 0.13.8 It is no longer recommended for use and may be removed in the future.

Use :ref:`ooodev.form.Forms` instead.
"""

# Use the following instead:
# from ooodev.form import Forms

from ooo.dyn.form.form_component_type import FormComponentType as FormComponentType
from ooodev.form.forms import Forms as Forms
from ooodev.utils.kind.form_component_kind import FormComponentKind as FormComponentKind

import warnings

warnings.warn("ooodev.utils.forms module is deprecated. Use ooodev.form instead.", DeprecationWarning, stacklevel=2)
