from __future__ import annotations
from typing import TYPE_CHECKING

from com.sun.star.form import XFormComponent
from ooodev.adapter.container.child_partial import ChildPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class FormComponentPartial(ChildPartial):
    """
    Partial Class for XFormComponent.

    Describes a component which may be part of a form.
    """

    def __init__(self, component: XFormComponent, interface: UnoInterface | None = XFormComponent) -> None:
        """
        Constructor

        Args:
            component (XFormComponent): UNO Component that implements ``com.sun.star.container.XFormComponent``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XFormComponent``.
        """
        ChildPartial.__init__(self, component, interface)
