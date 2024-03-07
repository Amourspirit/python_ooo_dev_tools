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

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

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
