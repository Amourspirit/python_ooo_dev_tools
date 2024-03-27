from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.awt import XControlContainer

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import XControl
    from ooodev.utils.type_var import UnoInterface


class ControlContainerPartial:
    """
    Partial class for XControlContainer.
    """

    def __init__(self, component: XControlContainer, interface: UnoInterface | None = XControlContainer) -> None:
        """
        Constructor

        Args:
            component (XControlContainer): UNO Component that implements ``com.sun.star.awt.XControlContainer`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XControlContainer``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XControlContainer
    def add_control(self, name: str, control: XControl) -> None:
        """
        Adds the given control with the specified name to the container.
        """
        self.__component.addControl(name, control)

    def get_control(self, name: str) -> XControl:
        """
        returns the control with the specified name.
        """
        return self.__component.getControl(name)

    def get_controls(self) -> Tuple[XControl, ...]:
        """
        returns all controls.
        """
        return self.__component.getControls()

    def remove_control(self, control: XControl) -> None:
        """
        removes the given control from the container.
        """
        self.__component.removeControl(control)

    def set_status_text(self, status_text: str) -> None:
        """
        Sets the status text in the status bar of the container.
        """
        self.__component.setStatusText(status_text)

    # endregion XControlContainer
