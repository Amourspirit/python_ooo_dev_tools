from __future__ import annotations
from typing import Any, TYPE_CHECKING
import contextlib
import uno

from ooodev.units.unit_mm100 import UnitMM100
from ooodev.units.angle100 import Angle100

if TYPE_CHECKING:
    from com.sun.star.drawing import MeasureProperties
    from ooo.dyn.drawing.measure_text_horz_pos import MeasureTextHorzPos
    from ooo.dyn.drawing.measure_text_vert_pos import MeasureTextVertPos
    from ooo.dyn.drawing.measure_kind import MeasureKind
    from ooodev.units.unit_obj import UnitT
    from ooodev.units.angle_t import AngleT


class MeasurePropertiesPartial:
    """
    Partial class for MeasureProperties Service.

    See Also:
        `API MeasureProperties <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1MeasureProperties.html>`_
    """

    def __init__(self, component: MeasureProperties) -> None:
        """
        Constructor

        Args:
            component (MeasureProperties): UNO Component that implements ``com.sun.star.drawing.MeasureProperties`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``MeasureProperties``.
        """
        self.__component = component

    # region MeasureProperties
    @property
    def measure_below_reference_edge(self) -> bool:
        """
        Gets/Sets - If this property is ``True``, the measure is drawn below the reference edge instead of above it.
        """
        return self.__component.MeasureBelowReferenceEdge

    @measure_below_reference_edge.setter
    def measure_below_reference_edge(self, value: bool) -> None:
        self.__component.MeasureBelowReferenceEdge = value

    @property
    def measure_decimal_places(self) -> int | None:
        """
        Gets/Sets the number of decimal places that is used to format the measure value.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.MeasureDecimalPlaces
        return None

    @measure_decimal_places.setter
    def measure_decimal_places(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.MeasureDecimalPlaces = value

    @property
    def measure_help_line1_length(self) -> UnitMM100:
        """
        Gets/Sets the length of the first help line.

        When setting the value can be an integer or a ``UnitT`` object.

        Returns:
            UnitMM100: The length of the first help line.
        """
        return UnitMM100(self.__component.MeasureHelpLine1Length)

    @measure_help_line1_length.setter
    def measure_help_line1_length(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.MeasureHelpLine1Length = val.value

    @property
    def measure_help_line2_length(self) -> UnitMM100:
        """
        Gets/Sets the length of the second help line.

        When setting the value can be an integer or a ``UnitT`` object.

        Returns:
            UnitMM100: The length of the second help line.
        """
        return UnitMM100(self.__component.MeasureHelpLine2Length)

    @measure_help_line2_length.setter
    def measure_help_line2_length(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.MeasureHelpLine2Length = val.value

    @property
    def measure_help_line_distance(self) -> UnitMM100:
        """
        Gets/Sets the distance from the measure line to the start of the help lines.

        When setting the value can be an integer or a ``UnitT`` object.

        Returns:
            UnitMM100: The distance from the measure line to the start of the help lines.
        """
        return UnitMM100(self.__component.MeasureHelpLineDistance)

    @measure_help_line_distance.setter
    def measure_help_line_distance(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.MeasureHelpLineDistance = val.value

    @property
    def measure_help_line_overhang(self) -> UnitMM100:
        """
        Gets/Sets the overhang of the two help lines.

        When setting the value can be an integer or a ``UnitT`` object.

        Returns:
            UnitMM100: The overhang of the two help lines.
        """
        return UnitMM100(self.__component.MeasureHelpLineOverhang)

    @measure_help_line_overhang.setter
    def measure_help_line_overhang(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.MeasureHelpLineOverhang = val.value

    @property
    def MeasureKind(self) -> MeasureKind:
        """
        Gets/Sets the MeasureKind.

        Returns:
            MeasureKind: The MeasureKind.

        Hint:
            - ``MeasureKind`` can be imported from ``ooo.dyn.drawing.measure_kind``.
        """
        return self.__component.MeasureKind  # type: ignore

    @MeasureKind.setter
    def MeasureKind(self, value: MeasureKind) -> None:
        self.__component.MeasureKind = value  # type: ignore

    @property
    def measure_line_distance(self) -> UnitMM100:
        """
        Gets/Sets the distance from the reference edge to the measure line.

        When setting the value can be an integer or a ``UnitT`` object.

        Returns:
            UnitMM100: The distance from the reference edge to the measure line.
        """
        return UnitMM100(self.__component.MeasureLineDistance)

    @measure_line_distance.setter
    def measure_line_distance(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.MeasureLineDistance = val.value

    @property
    def measure_overhang(self) -> UnitMM100:
        """
        Gets/Sets the overhang of the reference line over the help lines.

        When setting the value can be an integer or a ``UnitT`` object.

        Returns:
            UnitMM100: The overhang of the reference line over the help lines.
        """
        return UnitMM100(self.__component.MeasureOverhang)

    @measure_overhang.setter
    def measure_overhang(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.MeasureOverhang = val.value

    @property
    def measure_show_unit(self) -> bool:
        """
        Gets/Sets - If this is ``True``, the unit of measure is shown in the measure text.
        """
        return self.__component.MeasureShowUnit

    @measure_show_unit.setter
    def measure_show_unit(self, value: bool) -> None:
        self.__component.MeasureShowUnit = value

    @property
    def measure_text_auto_angle(self) -> bool:
        """
        Gets/Sets - If this is ``True``, the angle of the measure is set automatically.
        """
        return self.__component.MeasureTextAutoAngle

    @measure_text_auto_angle.setter
    def measure_text_auto_angle(self, value: bool) -> None:
        self.__component.MeasureTextAutoAngle = value

    @property
    def measure_text_auto_angle_view(self) -> Angle100:
        """
        This is the automatic angle.

        When setting the value can be an integer or an ``AngleT`` object.

        Returns:
            Angle100: The automatic angle.
        """
        return Angle100(self.__component.MeasureTextAutoAngleView)

    @measure_text_auto_angle_view.setter
    def measure_text_auto_angle_view(self, value: int | AngleT) -> None:
        val = Angle100.from_unit_val(value)
        self.__component.MeasureTextAutoAngleView = val.value

    @property
    def measure_text_fixed_angle(self) -> Angle100:
        """
        Gets/Sets the fixed angle.

        When setting the value can be an integer or an ``AngleT`` object.

        Return:
            Angle100: The fixed angle.
        """
        return Angle100(self.__component.MeasureTextFixedAngle)

    @measure_text_fixed_angle.setter
    def measure_text_fixed_angle(self, value: int | AngleT) -> None:
        val = Angle100.from_unit_val(value)
        self.__component.MeasureTextFixedAngle = val.value

    @property
    def measure_text_horizontal_position(self) -> MeasureTextHorzPos:
        """
        Gets/Sets the horizontal position of the measure text.

        Returns:
            MeasureTextHorzPos: The horizontal position of the measure text.

        Hint:
            - ``MeasureTextHorzPos`` can be imported from ``ooo.dyn.drawing.measure_text_horz_pos``.
        """
        return self.__component.MeasureTextHorizontalPosition  # type: ignore

    @measure_text_horizontal_position.setter
    def measure_text_horizontal_position(self, value: MeasureTextHorzPos) -> None:
        self.__component.MeasureTextHorizontalPosition = value  # type: ignore

    @property
    def measure_text_is_fixed_angle(self) -> bool:
        """
        Gets/Sets - If this value is ``True``, the measure has a fixed angle.
        """
        return self.__component.MeasureTextIsFixedAngle

    @measure_text_is_fixed_angle.setter
    def measure_text_is_fixed_angle(self, value: bool) -> None:
        self.__component.MeasureTextIsFixedAngle = value

    @property
    def measure_text_rotate90(self) -> bool:
        """
        Gets/Sets - If this value is ``True``, the text is rotated 90 degrees.
        """
        return self.__component.MeasureTextRotate90

    @measure_text_rotate90.setter
    def measure_text_rotate90(self, value: bool) -> None:
        self.__component.MeasureTextRotate90 = value

    @property
    def measure_text_upside_down(self) -> bool:
        """
        Gets/Sets - If this value is ``True``, the text is printed upside down.
        """
        return self.__component.MeasureTextUpsideDown

    @measure_text_upside_down.setter
    def measure_text_upside_down(self, value: bool) -> None:
        self.__component.MeasureTextUpsideDown = value

    @property
    def measure_text_vertical_position(self) -> MeasureTextVertPos:
        """
        Gets/Sets the vertical position of the text.

        Returns:
            MeasureTextVertPos: The vertical position of the text.

        Hint:
            - ``MeasureTextVertPos`` can be imported from ``ooo.dyn.drawing.measure_text_vert_pos``.
        """
        return self.__component.MeasureTextVerticalPosition  # type: ignore

    @measure_text_vertical_position.setter
    def measure_text_vertical_position(self, value: MeasureTextVertPos) -> None:
        self.__component.MeasureTextVerticalPosition = value  # type: ignore

    # endregion MeasureProperties
