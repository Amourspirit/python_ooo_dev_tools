from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.awt import XControlContainer

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import XControl
    from ooodev.utils.type_var import UnoInterface


class ControlContainerPartial:
    """
    Partial class for XControlContainer.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XControlContainer, interface: UnoInterface | None = XControlContainer) -> None:
        """
        Constructor

        Args:
            component (XControlContainer): UNO Component that implements ``com.sun.star.container.XControlContainer`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XControlContainer``.
        """

        self.__interface = interface
        self.__validate(component)
        self.__component = component

    def __validate(self, component: Any) -> None:
        """
        Validates the component.

        Args:
            component (Any): The component to be validated.
        """
        if self.__interface is None:
            return
        if not mLo.Lo.is_uno_interfaces(component, self.__interface):
            raise mEx.MissingInterfaceError(self.__interface)

    # region XControlContainer
    def get_control(self, name: str) -> XControl:
        """
        Returns the control with the specified name.
        """
        return self.__component.getControl(name)

    def get_controls(self) -> Tuple[XControl, ...]:
        """
        Returns all controls.
        """
        return self.__component.getControls()

    def remove_control(self, ctl: XControl) -> None:
        """
        Removes the given control from the container.
        """
        self.__component.removeControl(ctl)

    def set_status_text(self, text: str) -> None:
        """
        Sets the status text in the status bar of the container.
        """
        self.__component.setStatusText(text)

    # endregion XControlContainer
