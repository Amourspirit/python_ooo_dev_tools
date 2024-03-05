from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler
from ooodev.format.inner import style_factory
from ooodev.events.partial.events_partial import EventsPartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.write.char.font.font_position_t import FontPositionT
    from ooodev.format.writer.direct.char.font import CharSpacingKind
    from ooodev.format.writer.direct.char.font import FontScriptKind
    from ooodev.utils.data_type.intensity import Intensity
    from ooodev.units.angle import Angle
    from ooodev.units.unit_obj import UnitT


class FontPositionPartial:
    """
    Partial class for Char Font Position.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_font_position",
            after_event="after_style_font_position",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

    def style_font_position(
        self,
        script_kind: FontScriptKind | None = None,
        raise_lower: int | None = None,
        rel_size: int | Intensity | None = None,
        rotation: int | Angle | None = None,
        scale: int | None = None,
        fit: bool | None = None,
        spacing: CharSpacingKind | float | UnitT | None = None,
        pair: bool | None = None,
    ) -> FontPositionT | None:
        """
        Style Axis Line.

        Args:
            script_kind (FontScriptKind, optional): Specifies Superscript/Subscript option.
            raise_lower (int, optional): Specifies raise or Lower as percent value. Set to a value of 0 for automatic.
            rel_size (int, Intensity, optional): Specifies relative Font Size as percent value.
                Set this value to ``0`` for automatic; Otherwise value from ``1`` to ``100``.
            rotation (int, Angle, optional): Specifies the rotation of a character in degrees. Depending on the
                implementation only certain values may be allowed.
            scale (int, optional): Specifies scale width as percent value. Min value is ``1``.
            fit (bool, optional): Specifies if rotation is fit to line.
            spacing (CharSpacingKind, float, UnitT, optional): Specifies character spacing in ``pt`` (point) units
                or :ref:`proto_unit_obj`.
            pair (bool, optional): Specifies pair kerning.

        Raises:
            CancelEventError: If the event ``before_style_number_number`` is cancelled and not handled.

        Returns:
            FontPositionT | None: Font Only instance or ``None`` if cancelled.

        Hint:
            - ``Angle`` can be imported from ``ooodev.units``
            - ``CharSpacingKind`` can be imported from ``ooodev.format.writer.direct.char.font``
            - ``FontScriptKind`` can be imported from ``ooodev.format.writer.direct.char.font``
            - ``Intensity`` can be imported from ``ooodev.utils.data_type.intensity``
        """
        factory = style_factory.font_position_factory
        kwargs = {}
        if script_kind is not None:
            kwargs["script_kind"] = script_kind
        if raise_lower is not None:
            kwargs["raise_lower"] = raise_lower
        if rel_size is not None:
            kwargs["rel_size"] = rel_size
        if rotation is not None:
            kwargs["rotation"] = rotation
        if scale is not None:
            kwargs["scale"] = scale
        if fit is not None:
            kwargs["fit"] = fit
        if spacing is not None:
            kwargs["spacing"] = spacing
        if pair is not None:
            kwargs["pair"] = pair
        return self.__styler.style(factory=factory, **kwargs)

    def style_font_position_get(self) -> FontPositionT | None:
        """
        Gets the Font Position Style.

        Raises:
            CancelEventError: If the event ``before_style_number_number_get`` is cancelled and not handled.

        Returns:
            FontPositionT | None: Font Position style or ``None`` if cancelled.
        """
        return self.__styler.style_get(factory=style_factory.font_position_factory)
