from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING
import uno

from ooodev.mock import mock_g
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.exceptions import ex as mEx

if TYPE_CHECKING:
    from ooodev.format.inner.direct.chart2.series.data_labels.data_labels.attrib_options import PlacementKind
    from ooodev.format.inner.direct.chart2.series.data_labels.data_labels.attrib_options import SeparatorKind
    from ooodev.format.inner.direct.chart2.series.data_labels.data_labels.attrib_options import AttribOptions
else:
    PlacementKind = Any
    SeparatorKind = Any
    AttribOptions = Any


class Chart2DataLabelAttribOptPartial:
    """
    Partial class for Chart2 Legend Position.
    """

    def __init__(self, component: Any) -> None:
        self.__component = component

    def style_attribute_options(
        self, placement: PlacementKind | None = None, separator: SeparatorKind | None = None
    ) -> AttribOptions | None:
        """
        Style Chart2 Data Label Attribute Options.

        Args:
            placement (PlacementKind, optional): Specifies the placement of data labels relative to the objects.
            separator (SeparatorKind, optional): Specifies the separator between multiple text strings for the same object.

        Raises:
            CancelEventError: If the event ``before_style_chart2_data_label_attribute_opt`` is cancelled and not handled.

        Returns:
            AttribOptions | None: Attribute Options Style instance or ``None`` if cancelled.

        Hint:
            ``PlacementKind``, ``SeparatorKind`` and ``AttribOptions`` can be imported from ``ooodev.format.inner.direct.chart2.series.data_labels.data_labels.attrib_options``

        """
        # pylint: disable=import-outside-toplevel
        from ooodev.format.inner.direct.chart2.series.data_labels.data_labels.attrib_options import AttribOptions

        comp = self.__component
        has_events = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_attribute_options.__qualname__)
            event_data: Dict[str, Any] = {
                "placement": placement,
                "separator": separator,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_chart2_data_label_attribute_opt", cargs)
            if cargs.cancel is True:
                if cargs.handled is True:
                    return None
                cargs.set("initial_event", "before_style_chart2_data_label_attribute_opt")
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style has been cancelled.")
                else:
                    return None
            placement = cargs.event_data.get("placement", placement)
            separator = cargs.event_data.get("separator", separator)
            comp = cargs.event_data.get("this_component", comp)

        fe = AttribOptions(
            placement=placement,
            separator=separator,
        )

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event("after_style_chart2_data_label_attribute_opt", EventArgs.from_args(cargs))  # type: ignore
        return fe


if mock_g.FULL_IMPORT:
    from ooodev.format.inner.direct.chart2.series.data_labels.data_labels.attrib_options import AttribOptions
