from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.chart import X3DDefaultSetter

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class X3dDefaultSetterPartial:
    """
    Partial class for X3DDefaultSetter.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: X3DDefaultSetter, interface: UnoInterface | None = X3DDefaultSetter) -> None:
        """
        Constructor

        Args:
            component (X3DDefaultSetter): UNO Component that implements ``com.sun.star.chart.X3DDefaultSetter`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``X3DDefaultSetter``.
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

    # region X3DDefaultSetter

    def set3d_settings_to_default(self) -> None:
        """
        The result may depend on the current chart type and the current shade mode.
        """
        return self.__component.set3DSettingsToDefault()

    def set_default_illumination(self) -> None:
        """
        set suitable defaults for the illumination of the current 3D chart.

        The result may dependent on other 3D settings as rotation or shade mode. It may depend on the current chart type also.
        """
        return self.__component.setDefaultIllumination()

    def set_default_rotation(self) -> None:
        """
        sets a suitable default for the rotation of the current 3D chart.

        The result may depend on the current chart type.
        """
        return self.__component.setDefaultRotation()

    # endregion X3DDefaultSetter
