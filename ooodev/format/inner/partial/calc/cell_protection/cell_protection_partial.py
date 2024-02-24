from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING
import uno

from ooodev.mock import mock_g
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.style_named_event import StyleNameEvent
from ooodev.exceptions import ex as mEx
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial

if TYPE_CHECKING:
    from ooodev.format.inner.direct.calc.cell_protection.cell_protection import CellProtection


class CellProtectionPartial:
    """
    Partial class for Cell Protection.
    """

    def __init__(self, component: Any) -> None:
        self.__component = component

    def style_protection(
        self, hide_all: bool = False, protected: bool = False, hide_formula: bool = False, hide_print: bool = False
    ) -> CellProtection | None:
        """
        Style cell protection.

        AArgs:
            hide_all (bool, optional): Specifies if all is hidden. Defaults to ``False``.
            protected (bool, optional): Specifies protected value. Defaults to ``False``.
            hide_formula (bool, optional): Specifies if the formula is hidden. Defaults to ``False``.
            hide_print (bool, optional): Specifies if the cell are to be omitted during print. Defaults to ``False``.

        Raises:
            CancelEventError: If the event ``before_style_cell_protection`` is cancelled and not handled.

        Returns:
            CellProtection | None: Attribute Options Style instance or ``None`` if cancelled.

        Hint:
            ``PlacementKind``, ``SeparatorKind`` and ``CellProtection`` can be imported from ``ooodev.format.inner.direct.chart2.series.data_labels.data_labels.attrib_options``

        """
        # pylint: disable=import-outside-toplevel
        from ooodev.format.inner.direct.calc.cell_protection.cell_protection import CellProtection

        comp = self.__component
        has_events = False
        cancel_apply = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_protection.__qualname__)
            event_data: Dict[str, Any] = {
                "hide_all": hide_all,
                "protected": protected,
                "hide_formula": hide_formula,
                "hide_print": hide_print,
                "this_component": comp,
                "cancel_apply": cancel_apply,
            }
            cargs.event_data = event_data
            self.trigger_event(StyleNameEvent.STYLE_APPLYING, cargs)
            self.trigger_event("before_style_cell_protection", cargs)
            if cargs.cancel is True:
                if cargs.handled is True:
                    return None
                cargs.set("initial_event", "before_style_cell_protection")
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style has been cancelled.")
                else:
                    return None

            hide_all = cargs.event_data.get("hide_all", hide_all)
            protected = cargs.event_data.get("protected", protected)
            hide_formula = cargs.event_data.get("hide_formula", hide_formula)
            hide_print = cargs.event_data.get("hide_print", hide_print)
            cancel_apply = cargs.event_data.pop("cancel_apply", cancel_apply)
            comp = cargs.event_data.pop("this_component", comp)

        fe = CellProtection(hide_all=hide_all, protected=protected, hide_formula=hide_formula, hide_print=hide_print)

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore
        if cancel_apply is False:
            if isinstance(self, LoInstPropsPartial):
                with LoContext(inst=self.lo_inst):
                    fe.apply(comp)
            else:
                fe.apply(comp)
        fe.set_update_obj(comp)
        if cargs is not None:
            # pylint: disable=no-member
            event_args = EventArgs.from_args(cargs)
            event_args.event_data["styler_object"] = fe
            self.trigger_event("after_style_cell_protection", event_args)  # type: ignore
            self.trigger_event(StyleNameEvent.STYLE_APPLIED, event_args)  # type: ignore
        return fe

    def style_protection_get(self) -> CellProtection | None:
        """
        Gets the cell protection Style.

        Raises:
            CancelEventError: If the event ``before_style_cell_protection_get`` is cancelled and not handled.

        Returns:
            PercentFormat | None: Number percent style or ``None`` if cancelled.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.format.inner.direct.calc.cell_protection.cell_protection import CellProtection

        # pylint: disable=no-member
        comp = self.__component
        cargs = None
        if isinstance(self, EventsPartial):
            cargs = CancelEventArgs(self.style_protection_get.__qualname__)
            event_data: Dict[str, Any] = {
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_cell_protection_get", cargs)
            if cargs.cancel is True:
                if cargs.handled is not False:
                    return None
                cargs.set("initial_event", "before_style_cell_protection_get")
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style get has been cancelled.")
                else:
                    return None
            comp = cargs.event_data.get("this_component", comp)

        try:
            if isinstance(self, LoInstPropsPartial):
                with LoContext(inst=self.lo_inst):
                    style = CellProtection.from_obj(comp)
            else:
                style = CellProtection.from_obj(comp)
        except mEx.DisabledMethodError:
            return None

        style.set_update_obj(comp)
        return style


if mock_g.FULL_IMPORT:
    from ooodev.format.inner.direct.calc.cell_protection.cell_protection import CellProtection
