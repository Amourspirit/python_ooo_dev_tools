from __future__ import annotations
from typing import TYPE_CHECKING
import contextlib
import uno

from ooodev.utils import info as mInfo
from ooodev.adapter.container.index_replace_comp import IndexReplaceComp
from ooodev.adapter.text.text_columns_comp import TextColumnsComp
from ooodev.units.unit_px import UnitPX
from ooodev.units.unit_mm100 import UnitMM100

if TYPE_CHECKING:
    from com.sun.star.drawing import TextProperties
    from com.sun.star.container import XIndexReplace
    from com.sun.star.text import XTextColumns
    from ooo.dyn.drawing.text_animation_direction import TextAnimationDirection
    from ooo.dyn.drawing.text_animation_kind import TextAnimationKind
    from ooo.dyn.drawing.text_fit_to_size_type import TextFitToSizeType
    from ooo.dyn.drawing.text_horizontal_adjust import TextHorizontalAdjust
    from ooo.dyn.drawing.text_vertical_adjust import TextVerticalAdjust
    from ooo.dyn.text.writing_mode import WritingMode
    from ooodev.units.unit_obj import UnitT


class TextPropertiesPartial:
    """
    Partial class for TextProperties.

    See Also:
        `API TextProperties <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1TextProperties.html>`_
    """

    def __init__(self, component: TextProperties) -> None:
        """
        Constructor

        Args:
            component (TextProperties): UNO Component that implements ``com.sun.star.drawing.TextProperties`` service.
        """
        self.__component = component

    # region TextProperties
    @property
    def is_numbering(self) -> bool | None:
        """
        Gets/Sets - If this is ``True``, numbering is ON for the text of this Shape.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.IsNumbering
        return None

    @is_numbering.setter
    def is_numbering(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.IsNumbering = value

    @property
    def numbering_rules(self) -> IndexReplaceComp | None:
        """
        Gets/Sets the numbering levels.

        The different rules accessible with this ``com.sun.star.container.XIndexReplace`` interface are sequences of property values as described in the service ``com.sun.star.style.NumberingRule``.

        **optional**

        Returns:
            IndexReplaceComp | None: ``IndexReplaceComp`` object or ``None`` if not available.
        """
        if not hasattr(self.__component, "NumberingRules"):
            return None
        rules = self.__component.NumberingRules
        return None if rules is None else IndexReplaceComp(rules)

    @numbering_rules.setter
    def numbering_rules(self, value: XIndexReplace | IndexReplaceComp) -> None:
        if not hasattr(self.__component, "NumberingRules"):
            return
        if mInfo.Info.is_instance(value, IndexReplaceComp):
            self.__component.NumberingRules = value.component
        else:
            self.__component.NumberingRules = value  # type: ignore

    @property
    def text_animation_amount(self) -> UnitPX:
        """
        This is the number of pixels the text is moved in each animation step.

        When setting this property, you can use either an integer or a ``UnitT`` object.

        Returns:
            UnitPX: The number of pixels the text is moved in each animation step.

        Hint:
            - ``UnitPX`` can be imported from ``ooodev.units``.
        """
        return UnitPX(self.__component.TextAnimationAmount)

    @text_animation_amount.setter
    def text_animation_amount(self, value: int | UnitT) -> None:
        val = UnitPX.from_unit_val(value)
        self.__component.TextAnimationAmount = round(val.value)

    @property
    def text_animation_count(self) -> int:
        """
        This number defines how many times the text animation is repeated.

        If this is set to zero, the repeat is endless.
        """
        return self.__component.TextAnimationCount

    @text_animation_count.setter
    def text_animation_count(self, value: int) -> None:
        self.__component.TextAnimationCount = value

    @property
    def text_animation_delay(self) -> int:
        """
        This is the delay in thousandths of a second between each of the animation steps.
        """
        return self.__component.TextAnimationDelay

    @text_animation_delay.setter
    def text_animation_delay(self, value: int) -> None:
        self.__component.TextAnimationDelay = value

    @property
    def text_animation_direction(self) -> TextAnimationDirection:
        """
        Gets/Sets - This enumeration defines the direction in which the text moves.

        Returns:
            TextAnimationDirection: The direction in which the text moves.

        Hint:
            - ``TextAnimationDirection`` can be imported from ``ooo.dyn.drawing.text_animation_direction``.
        """
        return self.__component.TextAnimationDirection  # type: ignore

    @text_animation_direction.setter
    def text_animation_direction(self, value: TextAnimationDirection) -> None:
        self.__component.TextAnimationDirection = value  # type: ignore

    @property
    def text_animation_kind(self) -> TextAnimationKind:
        """
        Gets/Sets - This enumeration defines the type of animation.

        Returns:
            TextAnimationKind: The type of animation.

        Hint:
            - ``TextAnimationKind`` can be imported from ``ooo.dyn.drawing.text_animation_kind``.
        """
        return self.__component.TextAnimationKind  # type: ignore

    @text_animation_kind.setter
    def text_animation_kind(self, value: TextAnimationKind) -> None:
        self.__component.TextAnimationKind = value  # type: ignore

    @property
    def text_animation_start_inside(self) -> bool:
        """
        Gets/Sets - If this value is ``True``, the text is visible at the start of the animation.
        """
        return self.__component.TextAnimationStartInside

    @text_animation_start_inside.setter
    def text_animation_start_inside(self, value: bool) -> None:
        self.__component.TextAnimationStartInside = value

    @property
    def text_animation_stop_inside(self) -> bool:
        """
        Gets/Sets - If this value is ``True``, the text is visible at the end of the animation.
        """
        return self.__component.TextAnimationStopInside

    @text_animation_stop_inside.setter
    def text_animation_stop_inside(self, value: bool) -> None:
        self.__component.TextAnimationStopInside = value

    @property
    def text_auto_grow_height(self) -> bool:
        """
        Gets/Sets -- If this value is ``True``, the height of the Shape is automatically expanded/shrunk when text is added to or removed from the Shape.
        """
        return self.__component.TextAutoGrowHeight

    @text_auto_grow_height.setter
    def text_auto_grow_height(self, value: bool) -> None:
        self.__component.TextAutoGrowHeight = value

    @property
    def text_auto_grow_width(self) -> bool:
        """
        Gets/Sets - If this value is ``True``, the width of the Shape is automatically expanded/shrunk when text is added to or removed from the Shape.
        """
        return self.__component.TextAutoGrowWidth

    @text_auto_grow_width.setter
    def text_auto_grow_width(self, value: bool) -> None:
        self.__component.TextAutoGrowWidth = value

    @property
    def text_columns(self) -> TextColumnsComp | None:
        """
        Column layout properties for the text.

        **optional**

        **since**

            LibreOffice 7.2

        Returns:
            TextColumnsComp | None: ``TextColumnsComp`` object or ``None`` if not available.
        """
        if not hasattr(self.__component, "TextColumns"):
            return None
        cols = self.__component.TextColumns
        return None if cols is None else TextColumnsComp(cols)

    @text_columns.setter
    def text_columns(self, value: XTextColumns | TextColumnsComp) -> None:
        if not hasattr(self.__component, "TextColumns"):
            return
        if mInfo.Info.is_instance(value, TextColumnsComp):
            self.__component.TextColumns = value.component
        else:
            self.__component.TextColumns = value  # type: ignore

    @property
    def text_contour_frame(self) -> bool:
        """
        Gets/Sets - If this value is ``True``, the left edge of every line of text is aligned with the left edge of this Shape.
        """
        return self.__component.TextContourFrame

    @text_contour_frame.setter
    def text_contour_frame(self, value: bool) -> None:
        self.__component.TextContourFrame = value

    @property
    def text_fit_to_size(self) -> TextFitToSizeType:
        """
        Gets/Sets - With this set to a value other than ``NONE``, the text inside of the Shape is stretched or scaled to fit into the Shape.

        Returns:
            TextFitToSizeType: The type of text fit to size.

        Hint:
            - ``TextFitToSizeType`` can be imported from ``ooo.dyn.drawing.text_fit_to_size_type``.
        """
        return self.__component.TextFitToSize  # type: ignore

    @text_fit_to_size.setter
    def text_fit_to_size(self, value: TextFitToSizeType) -> None:
        self.__component.TextFitToSize = value  # type: ignore

    @property
    def text_horizontal_adjust(self) -> TextHorizontalAdjust:
        """
        Gets/Sets - Adjusts the horizontal position of the text inside of the Shape.

        Returns:
            TextHorizontalAdjust: Horizontal position of the text inside of the Shape.

        Hint:
            - ``TextHorizontalAdjust`` can be imported from ``ooo.dyn.drawing.text_horizontal_adjust``.
        """
        return self.__component.TextHorizontalAdjust  # type: ignore

    @text_horizontal_adjust.setter
    def text_horizontal_adjust(self, value: TextHorizontalAdjust) -> None:
        self.__component.TextHorizontalAdjust = value  # type: ignore

    @property
    def text_left_distance(self) -> UnitMM100:
        """
        This is the distance from the left edge of the Shape to the left edge of the text.

        This is only useful if ``Text.TextHorizontalAdjust`` is ``BLOCK`` or ``STRETCH`` or if ``Text.TextFitSize`` is ``True``.

        When setting this property, you can use either an integer or a ``UnitT`` object.

        Returns:
            UnitMM100: The distance from the left edge of the Shape to the left edge of the text.

        Hint:
            - ``UnitMM100`` can be imported from ``ooodev.units``.
        """
        return UnitMM100(self.__component.TextLeftDistance)

    @text_left_distance.setter
    def text_left_distance(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.TextLeftDistance = val.value

    @property
    def text_lower_distance(self) -> UnitMM100:
        """
        Gets/Sets the distance from the lower edge of the Shape to the lower edge of the text.

        This is only useful if ``Text.TextHorizontalAdjust`` is ``BLOCK`` or ``STRETCH`` or if ``Text.TextFitSize`` is ``True``.

        When setting this property, you can use either an integer or a ``UnitT`` object.

        Returns:
            UnitMM100: The distance from the lower edge of the Shape to the lower edge of the text.

        Hint:
            - ``UnitMM100`` can be imported from ``ooodev.units``.
        """
        return UnitMM100(self.__component.TextLowerDistance)

    @text_lower_distance.setter
    def text_lower_distance(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.TextLowerDistance = val.value

    @property
    def text_maximum_frame_height(self) -> UnitMM100:
        """
        Gets/Sets - With this property you can set the maximum height for a shape with text.

        On edit, the auto grow feature will not grow the object higher than the value of this property.

        When setting this property, you can use either an integer or a ``UnitT`` object.

        Returns:
            UnitMM100: The maximum height for a shape with text.

        Hint:
            - ``UnitMM100`` can be imported from ``ooodev.units``.
        """
        return UnitMM100(self.__component.TextMaximumFrameHeight)

    @text_maximum_frame_height.setter
    def text_maximum_frame_height(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.TextMaximumFrameHeight = val.value

    @property
    def text_maximum_frame_width(self) -> UnitMM100:
        """
        Gets/Sets - With this property you can set the maximum width for a shape with text.

        On edit, the auto grow feature will not grow the objects wider than the value of this property.

        When setting this property, you can use either an integer or a ``UnitT`` object.

        Returns:
            UnitMM100: The maximum width for a shape with text.

        Hint:
            - ``UnitMM100`` can be imported from ``ooodev.units``.
        """
        return UnitMM100(self.__component.TextMaximumFrameWidth)

    @text_maximum_frame_width.setter
    def text_maximum_frame_width(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.TextMaximumFrameWidth = val.value

    @property
    def text_minimum_frame_height(self) -> UnitMM100:
        """
        Gets/Sets with this property you can set the minimum height for a shape with text.

        On edit, the auto grow feature will not shrink the objects height smaller than the value of this property.

        When setting this property, you can use either an integer or a ``UnitT`` object.

        Returns:
            UnitMM100: The minimum height for a shape with text.

        Hint:
            - ``UnitMM100`` can be imported from ``ooodev.units``.
        """
        return UnitMM100(self.__component.TextMinimumFrameHeight)

    @text_minimum_frame_height.setter
    def text_minimum_frame_height(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.TextMinimumFrameHeight = val.value

    @property
    def text_minimum_frame_width(self) -> UnitMM100:
        """
        Gets/Sets - With this property you can set the minimum width for a shape with text.

        On edit, the auto grow feature will not shrink the object width smaller than the value of this property.

        When setting this property, you can use either an integer or a ``UnitT`` object.

        Returns:
            UnitMM100: The minimum width for a shape with text.

        Hint:
            - ``UnitMM100`` can be imported from ``ooodev.units``.
        """
        return UnitMM100(self.__component.TextMinimumFrameWidth)

    @text_minimum_frame_width.setter
    def text_minimum_frame_width(self, value: int) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.TextMinimumFrameWidth = val.value

    @property
    def text_right_distance(self) -> UnitMM100:
        """
        Gets/Sets the distance from the right edge of the Shape to the right edge of the text.

        This is only useful if ``Text.TextHorizontalAdjust`` is ``BLOCK`` or ``STRETCH`` or if ``Text.TextFitSize`` is ``True``.

        When setting this property, you can use either an integer or a ``UnitT`` object.

        Returns:
            UnitMM100: The distance from the right edge of the Shape to the right edge of the text.

        Hint:
            - ``UnitMM100`` can be imported from ``ooodev.units``.
        """
        return UnitMM100(self.__component.TextRightDistance)

    @text_right_distance.setter
    def text_right_distance(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.TextRightDistance = val.value

    @property
    def text_upper_distance(self) -> UnitMM100:
        """
        This is the distance from the upper edge of the Shape to the upper edge of the text.

        This is only useful if ``Text.TextHorizontalAdjust`` is ``BLOCK`` or ``STRETCH`` or if ``Text.TextFitSize`` is ``True``.

        When setting this property, you can use either an integer or a ``UnitT`` object.

        Returns:
            UnitMM100: The distance from the upper edge of the Shape to the upper edge of the text.

        Hint:
            - ``UnitMM100`` can be imported from ``ooodev.units``.
        """
        return UnitMM100(self.__component.TextUpperDistance)

    @text_upper_distance.setter
    def text_upper_distance(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.TextUpperDistance = val.value

    @property
    def text_vertical_adjust(self) -> TextVerticalAdjust:
        """
        Gets/Sets - This enum adjusts the vertical position of the text inside of the Shape.

        Returns:
            TextVerticalAdjust: Vertical position of the text inside of the Shape.

        Hint:
            - ``TextVerticalAdjust`` can be imported from ``ooo.dyn.drawing.text_vertical_adjust``.
        """
        return self.__component.TextVerticalAdjust  # type: ignore

    @text_vertical_adjust.setter
    def text_vertical_adjust(self, value: TextVerticalAdjust) -> None:
        self.__component.TextVerticalAdjust = value  # type: ignore

    @property
    def text_writing_mode(self) -> WritingMode:
        """
        This value selects the writing mode for the text.

        Returns:
            WritingMode: Writing mode for the text.

        Hint:
            - ``WritingMode`` can be imported from ``ooo.dyn.text.writing_mode``.
        """
        return self.__component.TextWritingMode  # type: ignore

    @text_writing_mode.setter
    def text_writing_mode(self, value: WritingMode) -> None:
        self.__component.TextWritingMode = value  # type: ignore

    # endregion TextProperties
