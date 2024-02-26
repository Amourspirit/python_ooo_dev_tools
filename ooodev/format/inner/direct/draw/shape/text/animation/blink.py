from __future__ import annotations
from typing import cast, TYPE_CHECKING, overload, TypeVar, Type, Any, Set
import uno
from ooo.dyn.drawing.text_animation_kind import TextAnimationKind

from ooodev.exceptions import ex as mEx
from ooodev.units.unit_px import UnitPX
from ooodev.units.unit_mm import UnitMM
from ooodev.utils import props as mProps
from ooodev.format.inner.direct.draw.shape.text.animation.no_effect import NoEffect

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

_TBlink = TypeVar("_TBlink", bound="Blink")


class Blink(NoEffect):
    """
    This class represents the text Animation of an object that supports ``com.sun.star.drawing.TextProperties``.
    """

    def __init__(
        self,
        start_inside: bool | None = None,
        visible_on_exit: bool | None = None,
        increment: float | UnitT | None = None,
        delay: int | None = None,
    ) -> None:
        """
        Constructor.

        Args:
            start_inside: bool, optional): Start inside.
                If this value is ``True``, the text is visible at the start of the animation.
            visible_on_exit (bool, optional): Visible on exit.
                If this value is ``True``, the text is visible at the end of the animation.
            increment (float, UnitT, optional): Increment value.
                A positive float represents MM units, a negative value represents pixels,
                or a ``UnitT`` instance can be passed in.
            delay (int, optional): Delay in milliseconds.
                This is the delay in thousandths of a second (milliseconds) between each of the animation steps.
                A value of 0 means automatic delay.
        """
        super().__init__()
        self.prop_start_inside = start_inside
        self.prop_visible_on_exit = visible_on_exit
        self.prop_increment = increment
        self.prop_delay = delay
        self._set_animation()

    # region Overridden Methods
    def _set_animation(self) -> None:
        self._set("TextAnimationKind", TextAnimationKind.BLINK)

    def _get_uno_props(self) -> Set[str]:
        props = super()._get_uno_props()
        props = props.union(
            {"TextAnimationStartInside", "TextAnimationStopInside", "TextAnimationAmount", "TextAnimationDelay"}
        )
        return props

    # endregion Overridden Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TBlink], obj: object) -> _TBlink: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TBlink], obj: object, **kwargs) -> _TBlink: ...

    @classmethod
    def from_obj(cls: Type[_TBlink], obj: Any, **kwargs) -> _TBlink:
        """
        Creates a new instance from ``obj``.

        Args:
            obj (Any): UNO Shape object.

        Returns:
            Spacing: New instance.
        """
        inst = cls(**kwargs)

        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError("Object is not supported for conversion to Line Properties")

        props = inst._get_uno_props()

        def set_property(prop: str):
            value = mProps.Props.get(obj, prop, None)
            if value is not None:
                inst._set(prop, value)

        for prop in props:
            set_property(prop)

        # set the animation kind possible overriding the value set by the properties
        inst._set_animation()
        return inst

    # endregion from_obj()

    # region Properties
    @property
    def prop_start_inside(self) -> bool | None:
        """If this value is ``True``, the text is visible at the start of the animation."""
        return self._get("TextAnimationStartInside")

    @prop_start_inside.setter
    def prop_start_inside(self, value: bool | None) -> None:
        if value is None:
            self._remove("TextAnimationStartInside")
            return
        self._set("TextAnimationStartInside", value)

    @property
    def prop_visible_on_exit(self) -> bool | None:
        """If this value is ``True``, the text is visible at the end of the animation."""
        return self._get("TextAnimationStopInside")

    @prop_visible_on_exit.setter
    def prop_visible_on_exit(self, value: bool | None) -> None:
        if value is None:
            self._remove("TextAnimationStopInside")
            return
        self._set("TextAnimationStopInside", value)

    @property
    def prop_increment(self) -> UnitMM | None:
        """
        Gets or sets This is the number of pixels the text is moved in each animation step.

        When setting the value, it can be a positive float in MM units
        or a negative integer in pixels or a ``UnitT`` instance.
        """
        pv = cast(int, self._get("TextAnimationAmount"))
        if pv is None:
            return None
        if pv < 0:
            # negative value is pixels
            return UnitMM.from_px(abs(pv))

        return UnitMM.from_mm100(pv)

    @prop_increment.setter
    def prop_increment(self, value: float | UnitT | None) -> None:
        # if value is negative then it is pixels and needs to be stored as negative integer.
        # if the value is positive float or int then is is mm and should be stored as positive integer for 1/100th mm units.
        if value is None:
            self._remove("TextAnimationAmount")
            return
        if hasattr(value, "get_value_mm"):
            # UnitT object
            if isinstance(value, UnitPX):
                # store as pixels in negative integer
                val = -abs(value.value)
                self._set("TextAnimationAmount", val)
            else:
                val = cast(UnitT, value)
                self._set("TextAnimationAmount", val.get_value_mm100())
        else:
            val = cast(float, value)
            if val < 0:
                # store as pixels in negative integer
                self._set("TextAnimationAmount", round(val))
            else:
                self._set("TextAnimationAmount", UnitMM(val).get_value_mm100())

    @property
    def prop_delay(self) -> int | None:
        """
        This is the delay in thousandths (milliseconds) of a second between each of the animation steps.

        A value of 0 means automatic delay.
        """
        return self._get("TextAnimationDelay")

    @prop_delay.setter
    def prop_delay(self, value: int | None) -> None:
        if value is None:
            self._remove("TextAnimationDelay")
            return
        if value < 0:
            value = 0
        self._set("TextAnimationDelay", value)

    # endregion Properties
