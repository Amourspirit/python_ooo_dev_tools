from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from com.sun.star.form import XFormComponent
from ooodev.adapter.container.child_partial import ChildPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class FormsComponentPartial(ChildPartial):
    """
    Partial class for Forms Component.
    """

    def __init__(self, component: XFormComponent, interface: UnoInterface | None = XFormComponent) -> None:
        """
        Constructor

        Args:
            component (XFormComponent): UNO Component that implements ``com.sun.star.form.XFormComponent`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XFormComponent``.
        """
        ChildPartial.__init__(self, component=component, interface=interface)
