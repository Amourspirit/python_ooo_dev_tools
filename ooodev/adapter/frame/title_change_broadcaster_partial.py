from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple

import uno
from ooo.dyn.frame.command_group import CommandGroupEnum
from com.sun.star.frame import XTitleChangeBroadcaster

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo


if TYPE_CHECKING:
    from com.sun.star.frame import XTitleChangeListener
    from ooodev.utils.type_var import UnoInterface


class TitleChangeBroadcasterPartial:
    """
    Partial class for XTitleChangeBroadcaster.
    """

    def __init__(
        self, component: XTitleChangeBroadcaster, interface: UnoInterface | None = XTitleChangeBroadcaster
    ) -> None:
        """
        Constructor

        Args:
            component (XTitleChangeBroadcaster): UNO Component that implements ``com.sun.star.frame.XTitleChangeBroadcaster`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTitleChangeBroadcaster``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTitleChangeBroadcaster
    def add_title_change_listener(self, listener: XTitleChangeListener) -> None:
        """
        Add a listener.
        """
        self.__component.addTitleChangeListener(listener)

    def remove_title_change_listener(self, listener: XTitleChangeListener) -> None:
        """
        Remove a listener.
        """
        self.__component.removeTitleChangeListener(listener)

    # endregion XTitleChangeBroadcaster
