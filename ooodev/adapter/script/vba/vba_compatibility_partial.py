from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.script.vba import XVBACompatibility

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.script.vba import XVBAScriptListener
    from ooodev.utils.type_var import UnoInterface


class VBACompatibilityPartial:
    """
    Partial class for XVBACompatibility.
    """

    def __init__(self, component: XVBACompatibility, interface: UnoInterface | None = XVBACompatibility) -> None:
        """
        Constructor

        Args:
            component (XVBACompatibility): UNO Component that implements ``com.sun.star.script.vba.XVBACompatibility`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XVBACompatibility``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XVBACompatibility
    def add_vba_script_listener(self, listener: XVBAScriptListener) -> None:
        """
        Add a listener to be notified when a VBA script event occurs.
        """
        self.__component.addVBAScriptListener(listener)

    def broadcast_vba_script_event(self, identifier: int, module_name: str) -> None:
        """
        Broadcast a VBA script event to all listeners.
        """
        self.__component.broadcastVBAScriptEvent(identifier, module_name)

    def remove_vba_script_listener(self, listener: XVBAScriptListener) -> None:
        """
        Remove a listener from the list of listeners.
        """
        self.__component.removeVBAScriptListener(listener)

    @property
    def project_name(self) -> str:
        """
        Gets/Sets the name of the project.
        """
        return self.__component.ProjectName

    @project_name.setter
    def project_name(self, value: str) -> None:

        self.__component.ProjectName = value

    @property
    def running_vba_scripts(self) -> int:
        """
        Gets the number of running VBA scripts.
        """
        return self.__component.RunningVBAScripts

    @property
    def vba_compatibility_mode(self) -> bool:
        """
        Gets/Sets the VBA compatibility mode.
        """
        return self.__component.VBACompatibilityMode

    @vba_compatibility_mode.setter
    def vba_compatibility_mode(self, value: bool) -> None:
        self.__component.VBACompatibilityMode = value

    # endregion XVBACompatibility
