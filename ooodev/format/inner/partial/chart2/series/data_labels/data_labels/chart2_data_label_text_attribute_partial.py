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
    from ooodev.format.inner.direct.chart2.series.data_labels.data_labels.text_attribs import TextAttribs


class Chart2DataLabelTextAttributePartial:
    """
    Partial class for Chart2 Data Labels Text Attribute.
    """

    def __init__(self, component: Any) -> None:
        self.__component = component

    def style_text_attributes(
        self,
        *,
        show_number: bool = False,
        show_number_in_percent: bool = False,
        show_category_name: bool = False,
        show_legend_symbol: bool = False,
        show_custom_label: bool = False,
        show_series_name: bool = False,
        auto_text_wrap: bool | None = None,
    ) -> TextAttribs | None:
        """
        Style Chart2 Data Series Text Attributes.

        Args:
            show_number (bool, optional): if ``True``, the value that is represented by a data point is displayed next to it. Defaults to ``False``.
            show_number_in_percent (bool, optional): Only effective, if ``ShowNumber`` is ``True``.
                If this member is also ``True``, the numbers are displayed as percentages of a category.
                That means, if a data point is the first one of a series, the percentage is calculated by using the first data points of all available series.
                Defaults to ``False``.
            show_category_name (bool, optional): Specifies the caption contains the category name of the category to which a data point belongs. Defaults to ``False``.
            show_legend_symbol (bool, optional): Specifies the symbol of data series is additionally displayed in the caption.
                Since LibreOffice ``7.1``. Defaults to ``False``.
            show_custom_label (bool, optional): Specifies the caption contains a custom label text, which belongs to a data point label. Defaults to ``False``.
            show_series_name (bool, optional): Specifies the name of the data series is additionally displayed in the caption.
                Since LibreOffice ``7.2``. Defaults to ``False``.
            auto_text_wrap (bool, optional): Specifies the text is automatically wrapped, if the text is too long to fit in the available space.

        Raises:
            CancelEventError: If the event ``before_style_chart2_data_label_text_attribs`` is cancelled and not handled.

        Returns:
            AttribOptions | None: Text Attribute Style instance or ``None`` if cancelled.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.format.inner.direct.chart2.series.data_labels.data_labels.text_attribs import TextAttribs

        comp = self.__component
        has_events = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_text_attributes.__qualname__)
            event_data: Dict[str, Any] = {
                "show_number": show_number,
                "show_number_in_percent": show_number_in_percent,
                "show_category_name": show_category_name,
                "show_legend_symbol": show_legend_symbol,
                "show_custom_label": show_custom_label,
                "show_series_name": show_series_name,
                "auto_text_wrap": auto_text_wrap,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_chart2_data_label_text_attribs", cargs)
            if cargs.cancel is True:
                if cargs.handled is not False:
                    return None
                cargs.set("initial_event", "before_style_chart2_data_label_text_attribs")
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style has been cancelled.")
                else:
                    return None
            show_number = cargs.event_data.get("show_number", show_number)
            show_number_in_percent = cargs.event_data.get("show_number_in_percent", show_number_in_percent)
            show_category_name = cargs.event_data.get("show_category_name", show_category_name)
            show_legend_symbol = cargs.event_data.get("show_legend_symbol", show_legend_symbol)
            show_custom_label = cargs.event_data.get("show_custom_label", show_custom_label)
            show_series_name = cargs.event_data.get("show_series_name", show_series_name)
            auto_text_wrap = cargs.event_data.get("auto_text_wrap", auto_text_wrap)
            comp = cargs.event_data.get("this_component", comp)

        fe = TextAttribs(
            show_number=show_number,
            show_number_in_percent=show_number_in_percent,
            show_category_name=show_category_name,
            show_legend_symbol=show_legend_symbol,
            show_custom_label=show_custom_label,
            show_series_name=show_series_name,
            auto_text_wrap=auto_text_wrap,
        )

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event("after_style_chart2_data_label_text_attribs", EventArgs.from_args(cargs))  # type: ignore
        return fe


if mock_g.FULL_IMPORT:
    from ooodev.format.inner.direct.chart2.series.data_labels.data_labels.text_attribs import TextAttribs
