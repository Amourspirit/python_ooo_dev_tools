from __future__ import annotations
from typing import TYPE_CHECKING

from com.sun.star.form import XForm
from ooodev.adapter.form.form_component_partial import FormComponentPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class FormPartial(FormComponentPartial):
    """
    Partial Class for XForm.

    This interface does not really provide an own functionality, it is only for easier runtime identification of form components.
    """

    def __init__(self, component: XForm, interface: UnoInterface | None = XForm) -> None:
        """
        Constructor

        Args:
            component (XForm): UNO Component that implements ``com.sun.star.container.XForm``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XForm``.
        """
        FormComponentPartial.__init__(self, component, interface)
