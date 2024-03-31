from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XSystemChildFactory
from ooo.dyn.lang.system_dependent import SystemDependentEnum
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import XWindowPeer
    from ooodev.utils.type_var import UnoInterface


class SystemChildFactoryPartial:
    """
    Partial class for XSystemChildFactory.
    """

    def __init__(self, component: XSystemChildFactory, interface: UnoInterface | None = XSystemChildFactory) -> None:
        """
        Constructor

        Args:
            component (XSystemChildFactory): UNO Component that implements ``com.sun.star.awt.XSystemChildFactory`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XSystemChildFactory``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XSystemChildFactory
    def create_system_child(
        self, parent: Any, system_type: int | SystemDependentEnum, *process_id: int
    ) -> XWindowPeer:
        """
        Creates a system child window.

        You must check the machine ID and the process ID.WIN32: HWND.WIN16: HWND.

        JAVA: global reference to a java.awt.Component object provided from the JNI-API.

        MAC: (NSView*) pointer.

        Args:
            parent (Any): a system-specific handle to a window.
            system_type (int, SystemDependentEnum): The system type.
            process_id (int): Ignored. One or more process ID.

        Returns:
            XWindowPeer: the created system child window.

        Hint:
            - ``SystemDependentEnum`` is an enum and can be imported from ``ooo.dyn.lang.system_dependent``.
        """
        return self.__component.createSystemChild(parent, process_id, system_type)  # type: ignore

    # endregion XSystemChildFactory
