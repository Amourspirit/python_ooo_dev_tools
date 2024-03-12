from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import contextlib
import uno

from ooodev.events.events import Events
from ooodev.utils.data_type.intensity import Intensity
from ooodev.units.unit_mm100 import UnitMM100

if TYPE_CHECKING:
    from com.sun.star.drawing import ShadowProperties
    from ooodev.events.args.key_val_args import KeyValArgs
    from ooodev.utils.color import Color
    from ooodev.units.unit_obj import UnitT


class ShadowPropertiesPartial:
    """
    Service Class

    This is a set of properties to describe the style for rendering a shadow.

    See Also:
        `API ShadowProperties <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1ShadowProperties.html>`_
    """

    def __init__(self, component: ShadowProperties) -> None:
        """
        Constructor

        Args:
            component (ShadowProperties): UNO Component that implements ``com.sun.star.drawing.ShadowProperties`` interface.
        """
        self.__event_provider = Events(self)
        self.__props = {}
        self.__component = component

        def on_comp_struct_changed(src: Any, event_args: KeyValArgs) -> None:
            prop_name = str(event_args.event_data["prop_name"])
            if hasattr(self.__component, prop_name):
                setattr(self.__component, prop_name, event_args.source.component)

        self.__fn_on_comp_struct_changed = on_comp_struct_changed
        # pylint: disable=no-member
        self.__event_provider.subscribe_event("com_sun_star_awt_Gradient_changed", self.__fn_on_comp_struct_changed)
        self.__event_provider.subscribe_event("com_sun_star_drawing_Hatch_changed", self.__fn_on_comp_struct_changed)

    # region ShadowProperties
    @property
    def shadow(self) -> bool:
        """
        Gets/Sets - enables/disables the shadow of a Shape.

        The other shadow properties are only applied if this is set to TRUE.
        """
        return self.__component.Shadow

    @shadow.setter
    def shadow(self, value: bool) -> None:
        self.__component.Shadow = value

    @property
    def shadow_blur(self) -> UnitMM100 | None:
        """
        This defines the degree of blur of the shadow in points.

        When setting value can be an int or ``UnitT``.

        **optional**

        Returns:
            UnitMM100 | None: UnitMM100 or None if property is not available.
        """
        with contextlib.suppress(AttributeError):
            return UnitMM100(self.__component.ShadowBlur)
        return None

    @shadow_blur.setter
    def shadow_blur(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        with contextlib.suppress(AttributeError):
            self.__component.ShadowBlur = val.value

    @property
    def shadow_color(self) -> Color:
        """
        This is the color of the shadow of this Shape.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        return self.__component.ShadowColor  # type: ignore

    @shadow_color.setter
    def shadow_color(self, value: Color) -> None:
        self.__component.ShadowColor = value  # type: ignore

    @property
    def shadow_transparence(self) -> Intensity:
        """
        This defines the degree of transparence of the shadow in percent.

        Returns:
            Intensity: Intensity
        """
        return Intensity(self.__component.ShadowTransparence)

    @shadow_transparence.setter
    def shadow_transparence(self, value: int | Intensity) -> None:
        self.__component.ShadowTransparence = Intensity(int(value)).value

    @property
    def shadow_x_distance(self) -> UnitMM100:
        """
        This is the horizontal distance of the left edge of the Shape to the shadow.

        When setting value can be an int or ``UnitT``.

        Returns:
            UnitMM100: UnitMM100
        """
        return UnitMM100(self.__component.ShadowXDistance)

    @shadow_x_distance.setter
    def shadow_x_distance(self, value: int | UnitT) -> None:
        self.__component.ShadowXDistance = UnitMM100.from_unit_val(value).value

    @property
    def shadow_y_distance(self) -> UnitMM100:
        """
        This is the vertical distance of the top edge of the Shape to the shadow.

        When setting value can be an int or ``UnitT``.

        Returns:
            UnitMM100: UnitMM100
        """
        return UnitMM100(self.__component.ShadowYDistance)

    @shadow_y_distance.setter
    def shadow_y_distance(self, value: int | UnitT) -> None:
        self.__component.ShadowYDistance = UnitMM100.from_unit_val(value).value

    # endregion ShadowProperties
