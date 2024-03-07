from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooo.dyn.awt.gradient import Gradient
from ooo.dyn.awt.gradient_style import GradientStyle
from ooodev.adapter.struct_base import StructBase
from ooodev.units.angle10 import Angle10
from ooodev.utils.data_type.intensity import Intensity

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT
    from ooodev.utils.color import Color
    from ooodev.units.angle_t import AngleT


class GradientStructComp(StructBase[Gradient]):
    """
    Gradient Struct.

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_awt_Gradient_changing``.
    The event raised after the property is changed is called ``com_sun_star_awt_Gradient_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: Gradient, prop_name: str, event_provider: EventsT | None = None) -> None:
        """
        Constructor

        Args:
            component (Gradient): Gradient.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT, optional): Event Provider.
        """
        super().__init__(component=component, prop_name=prop_name, event_provider=event_provider)

    # region Overrides
    def _get_on_changing_event_name(self) -> str:
        return "com_sun_star_awt_Gradient_changing"

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_awt_Gradient_changed"

    def _copy(self, src: Gradient | None = None) -> Gradient:
        if src is None:
            src = self.component
        return Gradient(
            Style=src.Style,
            StartColor=src.StartColor,
            EndColor=src.EndColor,
            Angle=src.Angle,
            Border=src.Border,
            XOffset=src.XOffset,
            YOffset=src.YOffset,
            StartIntensity=src.StartIntensity,
            EndIntensity=src.EndIntensity,
            StepCount=src.StepCount,
        )

    # endregion Overrides

    # region Properties
    @property
    def style(self) -> GradientStyle:
        """
        Gets/Sets the style of the gradient.

        Returns:
            GradientStyle: Gradient Style.

        Hint:
            - ``GradientStyle`` can be imported from ``ooodev.awt.gradient_style``
        """
        return self.component.Style  # type: ignore

    @style.setter
    def style(self, value: GradientStyle) -> None:
        old_value = self.component.Style
        if old_value != value:
            event_args = self._trigger_cancel_event("Style", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def start_color(self) -> Color:
        """
        Gets/Sets the color at the start point of the gradient.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        return self.component.StartColor  # type: ignore

    @start_color.setter
    def start_color(self, value: Color) -> None:
        old_value = self.component.StartColor
        if old_value != value:
            event_args = self._trigger_cancel_event("StartColor", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def end_color(self) -> Color:
        """
        Gets/Sets the color at the end point of the gradient.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        return self.component.EndColor  # type: ignore

    @end_color.setter
    def end_color(self, value: Color) -> None:
        old_value = self.component.EndColor
        if old_value != value:
            event_args = self._trigger_cancel_event("EndColor", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def angle(self) -> Angle10:
        """
        Gets/Sets angle of the gradient in ``1/10`` degree.

        When setting the value can be set with an ``int`` in ``1/10`` degrees or an ``AngleT`` instance.

        Returns:
            Angle10: Angle in ``1/10`` degree.

        Hint:
            - ``Angle10`` can be imported from ``ooodev.units``.
        """
        return Angle10(self.component.Angle)

    @angle.setter
    def angle(self, value: int | AngleT) -> None:
        old_value = self.component.Angle
        val = Angle10.from_unit_val(value)
        new_value = val.value
        if old_value != new_value:
            event_args = self._trigger_cancel_event("Angle", old_value, new_value)
            self._trigger_done_event(event_args)

    @property
    def border(self) -> int:
        """
        Gets/Sets per cent of the total width where just the start color is used.

        Must be from ``0`` to ``100``.
        """
        return self.component.Border

    @border.setter
    def border(self, value: int) -> None:
        val = Intensity(value).value
        old_value = self.component.Border
        if old_value != val:
            event_args = self._trigger_cancel_event("Border", old_value, val)
            self._trigger_done_event(event_args)

    @property
    def x_offset(self) -> int:
        """
        Gets/Sets the X-coordinate, where the gradient begins.

        This is effectively the center of the ``RADIAL``, ``ELLIPTICAL``, ``SQUARE`` and ``RECT`` style gradients.

        Must be from ``0`` to ``100``.
        """
        return self.component.XOffset

    @x_offset.setter
    def x_offset(self, value: int) -> None:
        val = Intensity(value).value
        old_value = self.component.XOffset
        if old_value != val:
            event_args = self._trigger_cancel_event("XOffset", old_value, val)
            self._trigger_done_event(event_args)

    @property
    def y_offset(self) -> int:
        """
        Gets/Sets the Y-coordinate, where the gradient begins.

        Must be from ``0`` to ``100``.
        """
        ...

    @y_offset.setter
    def y_offset(self, value: int) -> None:
        val = Intensity(value).value
        old_value = self.component.YOffset
        if old_value != val:
            event_args = self._trigger_cancel_event("YOffset", old_value, val)
            self._trigger_done_event(event_args)

    @property
    def start_intensity(self) -> Intensity:
        """
        Specifies the intensity at the start point of the gradient.

        What that means is undefined.

        Returns:
            Intensity: Intensity at the start point of the gradient.

        Hint:
            - ``Intensity`` can be imported from ``ooodev.utils.data_type.intensity``
        """
        return Intensity(self.component.StartIntensity)

    @start_intensity.setter
    def start_intensity(self, value: int | Intensity) -> None:
        val = Intensity(int(value)).value
        old_value = self.component.StartIntensity
        if old_value != val:
            event_args = self._trigger_cancel_event("StartIntensity", old_value, val)
            self._trigger_done_event(event_args)

    @property
    def end_intensity(self) -> Intensity:
        """
        Gets/Sets the intensity at the end point of the gradient.

        Returns:
            Intensity: Intensity at the end point of the gradient.

        Hint:
            - ``Intensity`` can be imported from ``ooodev.utils.data_type.intensity``
        """
        return Intensity(self.component.EndIntensity)

    @end_intensity.setter
    def end_intensity(self, value: int | Intensity) -> None:
        val = Intensity(int(value)).value
        old_value = self.component.EndIntensity
        if old_value != val:
            event_args = self._trigger_cancel_event("EndIntensity", old_value, val)
            self._trigger_done_event(event_args)

    @property
    def step_count(self) -> int:
        """
        Gets/Sets the number of steps of change color.
        """
        # step_count must be between 3 and 256 when not automatic in paragraph gradient
        return self.component.StepCount

    @step_count.setter
    def step_count(self, value: int) -> None:
        old_value = self.component.StepCount
        if old_value != value:
            event_args = self._trigger_cancel_event("StepCount", old_value, value)
            self._trigger_done_event(event_args)

    # endregion Properties
