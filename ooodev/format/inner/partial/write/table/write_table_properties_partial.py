from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING
import contextlib
import uno

from ooodev.mock import mock_g
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.style_named_event import StyleNameEvent
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial

if TYPE_CHECKING:
    from ooodev.format.inner.direct.write.table.props.table_properties import (
        TblAbsUnit,
        TblRelUnit,
        TableAlignKind,
        TableProperties,
    )


class WriteTablePropertiesPartial:
    """
    Partial class for Write Table Properties.
    """

    def __init__(self, component: Any) -> None:
        self.__component = component

    def style_table_props(
        self,
        *,
        width: TblAbsUnit | TblRelUnit | None = None,
        left: TblAbsUnit | TblRelUnit | None = None,
        right: TblAbsUnit | TblRelUnit | None = None,
        above: TblAbsUnit | TblRelUnit | None = None,
        below: TblAbsUnit | TblRelUnit | None = None,
        align: TableAlignKind | None = None,
        relative: bool = False,
    ) -> TableProperties | None:
        """
        Style Write Table Properties.

        Args:
            width (TblAbsUnit, TblRelUnit, optional): Specifies table Width.
            left (TblAbsUnit, TblRelUnit, optional): Specifies table Left.
            right (TblAbsUnit, TblRelUnit, optional): Specifies table Right.
            above (TblAbsUnit, TblRelUnit, optional): Specifies table spacing above.
            below (TblAbsUnit, TblRelUnit, optional): Specifies table spacing below.
            align (TableAlignKind, optional): Specifies table alignment.
            relative (bool, optional): Specifies if table horizontal values are in percentages or ``mm`` units.


        Raises:
            CancelEventError: If the event ``before_style_table_properties`` is cancelled and not handled.

        Returns:
            TableProperties | None: Table Properties Style instance or ``None`` if cancelled.

        Hint:
            - ``TblAbsUnit`` can be imported from ``ooodev.format.inner.direct.write.table.props.table_properties``
            - ``TblRelUnit`` can be imported from ``ooodev.format.inner.direct.write.table.props.table_properties``
            - ``TableAlignKind`` can be imported from ``ooodev.format.inner.direct.write.table.props.table_properties``

        """
        # pylint: disable=import-outside-toplevel
        # from ooodev.format.inner.direct.write.table.props import table_properties as tp
        from ooodev.format.inner.direct.write.table.props.table_properties import TableProperties

        comp = self.__component
        has_events = False
        cargs = None
        # name = self.__table.name
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_table_props.__qualname__)
            event_data: Dict[str, Any] = {
                "width": width,
                "left": left,
                "right": right,
                "above": above,
                "below": below,
                "align": align,
                "relative": relative,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event(StyleNameEvent.STYLE_APPLYING, cargs)
            self.trigger_event("before_style_table_properties", cargs)
            if cargs.cancel is True:
                if cargs.handled is True:
                    return None
                cargs.set("initial_event", "before_style_table_properties")
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style has been cancelled.")
                else:
                    return None

            width = cargs.event_data.get("width", width)
            left = cargs.event_data.get("left", left)
            right = cargs.event_data.get("right", right)
            above = cargs.event_data.get("above", above)
            below = cargs.event_data.get("below", below)
            align = cargs.event_data.get("align", align)
            relative = cargs.event_data.get("relative", relative)

            comp = cargs.event_data.get("this_component", comp)

        fe = TableProperties(
            width=width,
            left=left,
            right=right,
            above=above,
            below=below,
            align=align,
            relative=relative,
        )

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        if isinstance(self, LoInstPropsPartial):
            lo_inst = self.lo_inst
        else:
            lo_inst = mLo.Lo.current_lo
        with LoContext(lo_inst):
            fe.apply(comp)

        fe.set_update_obj(comp)
        if cargs is not None:
            # pylint: disable=no-member
            event_args = EventArgs.from_args(cargs)
            event_args.event_data["styler_object"] = fe
            self.trigger_event("after_style_table_properties", event_args)  # type: ignore
            self.trigger_event(StyleNameEvent.STYLE_APPLIED, event_args)  # type: ignore
        return fe

    def style_table_props_get(self) -> TableProperties | None:
        """
        Gets the Table Properties Style.

        Returns:
            TableProperties | None: Table Properties Style instance or ``None`` if not available.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.format.inner.direct.write.table.props.table_properties import TableProperties

        with contextlib.suppress(Exception):
            return TableProperties.from_obj(self.__component)
        return None


if mock_g.FULL_IMPORT:
    from ooodev.format.inner.direct.write.table.props.table_properties import TableAlignKind, TableProperties
