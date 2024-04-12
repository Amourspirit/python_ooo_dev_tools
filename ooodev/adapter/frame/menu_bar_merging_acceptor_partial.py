from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.frame import XMenuBarMergingAcceptor

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.container import XIndexAccess
    from ooodev.utils.type_var import UnoInterface


class MenuBarMergingAcceptorPartial:
    """
    Partial class for XMenuBarMergingAcceptor.
    """

    def __init__(
        self, component: XMenuBarMergingAcceptor, interface: UnoInterface | None = XMenuBarMergingAcceptor
    ) -> None:
        """
        Constructor

        Args:
            component (XMenuBarMergingAcceptor): UNO Component that implements ``com.sun.star.frame.XMenuBarMergingAcceptor`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XMenuBarMergingAcceptor``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XMenuBarMergingAcceptor
    def remove_merged_menu_bar(self) -> None:
        """
        removes a previously set merged menu bar and sets a previously created menu bar back.
        """
        self.__component.removeMergedMenuBar()

    def set_merged_menu_bar(self, merged_menu_bar: XIndexAccess) -> bool:
        """
        allows to set a merged menu bar.

        This function is normally used to provide inplace editing where functions from two application parts, container application and embedded object, are available to the user simultaneously. A menu bar which is set by this method has a higher priority than others created by com.sun.star.frame.XLayoutManager interface. Settings of a merged menu bar cannot be retrieved.
        """
        return self.__component.setMergedMenuBar(merged_menu_bar)

    # endregion XMenuBarMergingAcceptor
