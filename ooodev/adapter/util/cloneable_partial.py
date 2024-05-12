from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.util import XCloneable

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class CloneablePartial:
    """
    Partial Class XCloneable.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XCloneable, interface: UnoInterface | None = XCloneable) -> None:
        """
        Constructor

        Args:
            component (XCloneable): UNO Component that implements ``com.sun.star.util.XCloneable`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XCloneable``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XCloneable
    def create_clone(self) -> XCloneable:
        """
        Creates a clone of the object.

        Returns:
            XCloneable: The clone.
        """
        return self.__component.createClone()

    # endregion XCloneable
