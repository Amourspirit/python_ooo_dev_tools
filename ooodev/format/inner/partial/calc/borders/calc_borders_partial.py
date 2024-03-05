from __future__ import annotations
from typing import Any, cast, Dict, TYPE_CHECKING

from ooodev.mock import mock_g
from ooodev.format.inner.style_factory import calc_borders_factory
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler


if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT
    from ooodev.utils.color import Color
    from ooodev.events.args.cancel_event_args import CancelEventArgs
    from ooodev.format.inner.direct.structs.side import Side
    from ooodev.format.inner.direct.calc.border.padding import Padding
    from ooodev.format.inner.direct.calc.border.shadow import Shadow
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.inner.direct.structs.side import BorderLineKind
    from ooodev.format.inner.direct.structs.side import LineSize
    from ooodev.format.proto.calc.borders.borders_t import BordersT


class CalcBordersPartial:
    """
    Partial class for Calc Borders.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_calc_borders",
            after_event="after_style_calc_borders",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

    def style_borders_clear(self) -> None:
        """
        Clear Borders Styles.
        """

        def on_style(src: Any, events: CancelEventArgs) -> None:
            # this will stop the style from being applied.
            # not critical but style is being applied later via style.empty.update()
            events.event_data["cancel_apply"] = True

        styler = self.__styler
        factory = calc_borders_factory
        if isinstance(self, EventsPartial):
            # on_style will automatically removed when it is out of scope.
            self.subscribe_event(event_name=styler.before_event_name, callback=on_style)
        style = cast("BordersT", styler.style(factory=factory))
        if style is not None:
            empty = style.empty
            if not empty.has_update_obj():
                empty.set_update_obj(style.get_update_obj())
            style.empty.update()

    def style_borders_default(self) -> BordersT | None:
        """
        Style default border.

        Returns:
            BordersT | None: Border instance or ``None`` if cancelled.
        """

        def on_style(src: Any, events: CancelEventArgs) -> None:
            # this will stop the style from being applied.
            # not critical but style is being applied later via style.empty.update()
            events.event_data["cancel_apply"] = True

        styler = self.__styler
        factory = calc_borders_factory
        if isinstance(self, EventsPartial):
            # on_style will automatically removed when it is out of scope.
            self.subscribe_event(event_name=styler.before_event_name, callback=on_style)
        style = cast("BordersT", styler.style(factory=factory))
        if style is None:
            return None

        default = style.default
        if not default.has_update_obj():
            default.set_update_obj(style.get_update_obj())
        default.update()
        return default

    def style_borders(
        self,
        *,
        right: Side | None = None,
        left: Side | None = None,
        top: Side | None = None,
        bottom: Side | None = None,
        border_side: Side | None = None,
        vertical: Side | None = None,
        horizontal: Side | None = None,
        distance: float | UnitT | None = None,
        diagonal_down: Side | None = None,
        diagonal_up: Side | None = None,
        shadow: Shadow | None = None,
        padding: Padding | None = None,
    ) -> BordersT | None:
        """
        Style Borders.

        Args:
            left (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the left edge.
            right (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the right edge.
            top (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the top edge.
            bottom (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the bottom edge.
            border_side (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the top, bottom, left, right edges. If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
            horizontal (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style of horizontal lines for the inner part of a cell range.
            vertical (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style of vertical lines for the inner part of a cell range.
            distance (float, UnitT, optional): Contains the distance between the lines and other contents in ``mm`` units or :ref:`proto_unit_obj`.
            diagonal_down (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style from top-left to bottom-right diagonal.
            diagonal_up (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style from bottom-left to top-right diagonal.
            shadow (~ooodev.format.inner.direct.calc.border.shadow.Shadow, optional): Cell Shadow.
            padding (padding, optional): Cell padding.

        Raises:
            CancelEventError: If the event ``before_style_calc_borders`` is cancelled and not handled.

        Returns:
            BordersT | None: Border instance or ``None`` if cancelled.

        Hint:
            - ``Side``, ``Shadow`` and ``Padding`` can be imported from ``ooodev.format.calc.direct.cell.borders``
            - ``BorderLineKind`` can be imported from ``ooodev.format.calc.direct.cell.borders``
            - ``LineSize`` can be imported from ``ooodev.format.calc.direct.cell.borders``
        """
        styler = self.__styler
        factory = calc_borders_factory
        kwargs: Dict[str, Any] = {
            "right": right,
            "left": left,
            "top": top,
            "bottom": bottom,
            "border_side": border_side,
            "vertical": vertical,
            "horizontal": horizontal,
            "distance": distance,
            "diagonal_down": diagonal_down,
            "diagonal_up": diagonal_up,
            "shadow": shadow,
            "padding": padding,
        }
        return styler.style(factory=factory, **kwargs)

    def style_borders_sides(
        self,
        *,
        line: BorderLineKind | None = None,
        color: Color | None = None,
        width: LineSize | float | UnitT | None = None,
        distance: float | UnitT | None = None,
        shadow: Shadow | None = None,
        padding: Padding | None = None,
        hori: bool = False,
        vert: bool = False,
    ) -> BordersT | None:
        """
        Style All border to specified line properties.
        This method is a subset of ``style_borders()`` method for convenience.

        Args:
            line (BorderLineStyleEnum, optional): Line Style of the border. Default ``BorderLineKind.SOLID``.
            color (:py:data:`~.utils.color.Color`, optional): Color of the border. Default ``StandardColor.BLACK``
            width (LineSize, float, UnitT, optional): Contains the width in of a single line or the width of outer part of a double line (in ``pt`` units) or :ref:`proto_unit_obj`. If this value is zero, no line is drawn. Default ``LineSize.THIN``
            distance (float, UnitT, optional): Contains the distance between the lines and other contents in ``mm`` units or :ref:`proto_unit_obj`.
            shadow (~ooodev.format.inner.direct.calc.border.shadow.Shadow, optional): Cell Shadow.
            padding (padding, optional): Cell padding.
            hori (bool, optional): If ``True`` then horizontal lines are also styled.
            vert (bool, optional): If ``True`` then vertical lines are also styled.

        Raises:
            CancelEventError: If the event ``before_style_calc_borders`` is cancelled and not handled.

        Returns:
            BordersT | None: Font Effects instance or ``None`` if cancelled.

        Hint:
            - ``Shadow`` and ``Padding`` can be imported from ``ooodev.format.calc.direct.cell.borders``
            - ``BorderLineKind`` can be imported from ``ooodev.format.calc.direct.cell.borders``
            - ``LineSize`` can be imported from ``ooodev.format.calc.direct.cell.borders``
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.format.inner.direct.structs.side import Side
        from ooodev.format.inner.direct.structs.side import BorderLineKind
        from ooodev.format.inner.direct.structs.side import LineSize

        styler = self.__styler

        side_args = {}
        if line is None:
            side_args["line"] = BorderLineKind.SOLID
        else:
            side_args["line"] = line
        if color is None:
            side_args["color"] = 0  # black
        else:
            side_args["color"] = color
        if width is None:
            side_args["width"] = LineSize.THIN
        else:
            side_args["width"] = width

        border_side = Side(**side_args)

        factory = calc_borders_factory
        kwargs: Dict[str, Any] = {
            "border_side": border_side,
            "distance": distance,
            "shadow": shadow,
            "padding": padding,
        }
        if hori:
            kwargs["horizontal"] = border_side
        if vert:
            kwargs["vertical"] = border_side
        return styler.style(factory=factory, **kwargs)

    def style_borders_get(self) -> BordersT | None:
        """
        Gets the Borders Style.

        Raises:
            CancelEventError: If the event ``before_style_calc_borders_get`` is cancelled and not handled.

        Returns:
            BordersT | None: border style or ``None`` if cancelled.
        """
        return self.__styler.style_get(factory=calc_borders_factory)


if mock_g.FULL_IMPORT:
    from ooodev.format.inner.direct.structs.side import Side
    from ooodev.format.inner.direct.structs.side import BorderLineKind
    from ooodev.format.inner.direct.structs.side import LineSize
