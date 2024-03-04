from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.form import XForms

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.container.container_partial import ContainerPartial
from ooodev.adapter.container.child_partial import ChildPartial
from ooodev.adapter.container.name_container_partial import NameContainerPartial
from ooodev.adapter.container.index_container_partial import IndexContainerPartial
from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial
from ooodev.adapter.util.cloneable_partial import CloneablePartial
from ooodev.adapter.lang.component_partial import ComponentPartial


if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class FormsPartial(
    ContainerPartial,
    NameContainerPartial,
    IndexContainerPartial,
    EnumerationAccessPartial,
    ChildPartial,
    CloneablePartial,
    ComponentPartial,
):
    def __init__(self, component: XForms, interface: UnoInterface | None = XForms) -> None:
        """
        Constructor

        Args:
            component (XForms): UNO Component that implements ``com.sun.star.form.XForms`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XForms``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        ContainerPartial.__init__(self, component, interface)
        NameContainerPartial.__init__(self, component, interface)
        IndexContainerPartial.__init__(self, component, interface)
        EnumerationAccessPartial.__init__(self, component, interface)
        CloneablePartial.__init__(self, component, interface)
        ChildPartial.__init__(self, component, interface)
        ComponentPartial.__init__(self, component, interface)
