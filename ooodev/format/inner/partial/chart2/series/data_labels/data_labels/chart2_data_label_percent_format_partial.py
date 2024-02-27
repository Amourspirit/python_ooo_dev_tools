from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING
import uno

from ooodev.mock import mock_g
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.exceptions import ex as mEx
from ooodev.calc.chart2.partial.chart_doc_prop_partial import ChartDocPropPartial

if TYPE_CHECKING:
    from ooo.dyn.util.number_format import NumberFormatEnum
    from ooo.dyn.lang.locale import Locale
    from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
    from ooodev.format.inner.direct.chart2.series.data_labels.data_labels.percent_format import PercentFormat


class Chart2DataLabelPercentFormatPartial:
    """
    Partial class for Chart2 Data Labels Percent Format.
    """

    def __init__(self, component: Any) -> None:
        self.__component = component

    def style_numbers_percent(
        self,
        *,
        source_format: bool = True,
        num_format: NumberFormatEnum | int = 0,
        num_format_index: NumberFormatIndexEnum | int = -1,
        lang_locale: Locale | None = None,
    ) -> PercentFormat | None:
        """
        Style Chart2 Data Series Text Attributes.

        Args:
            chart_doc (XChartDocument): Chart document.
            source_format (bool, optional): Specifies whether the number format should be linked to the source format. Defaults to ``True``.
            num_format (NumberFormatEnum, int, optional): specifies the number format. Defaults to ``0``.
            num_format_index (NumberFormatIndexEnum | int, optional): Specifies the number format index. Defaults to ``-1``.
            lang_locale (Locale, optional): Specifies the language locale. Defaults to ``None``.

        Raises:
            CancelEventError: If the event ``before_style_chart2_data_label_format_percent`` is cancelled and not handled.

        Returns:
            PercentFormat | None: Percent Format Style instance or ``None`` if cancelled.

        Hint:
            - ``NumberFormatEnum`` can be imported from ``ooo.dyn.util.number_format``
            - ``NumberFormatIndexEnum`` can be imported from ``ooo.dyn.i18n.number_format_index``
            - ``PercentFormat`` can be imported from ``ooodev.format.inner.direct.chart2.series.data_labels.data_labels.percent_format``
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.format.inner.direct.chart2.series.data_labels.data_labels.percent_format import PercentFormat

        if not isinstance(self, ChartDocPropPartial):
            raise TypeError("The class must be inherited from ChartDocPropPartial to use this method.")
        # pylint: disable=no-member
        doc = self.chart_doc.component
        comp = self.__component
        has_events = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_numbers_percent.__qualname__)
            event_data: Dict[str, Any] = {
                "source_format": source_format,
                "num_format": num_format,
                "num_format_index": num_format_index,
                "lang_locale": lang_locale,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_chart2_data_label_format_percent", cargs)
            if cargs.cancel is True:
                if cargs.handled is not False:
                    return None
                cargs.set("initial_event", "before_style_chart2_data_label_format_percent")
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style has been cancelled.")
                else:
                    return None
            source_format = cargs.event_data.get("source_format", source_format)
            num_format = cargs.event_data.get("num_format", num_format)
            num_format_index = cargs.event_data.get("num_format_index", num_format_index)
            lang_locale = cargs.event_data.get("lang_locale", lang_locale)
            comp = cargs.event_data.get("this_component", comp)

        fe = PercentFormat(
            chart_doc=doc,
            source_format=source_format,
            num_format=num_format,
            num_format_index=num_format_index,
            lang_locale=lang_locale,
        )

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event("after_style_chart2_data_label_format_percent", EventArgs.from_args(cargs))  # type: ignore
        return fe

    def style_numbers_percent_get(self) -> PercentFormat | None:
        """
        Gets the number percent Style.

        Raises:
            CancelEventError: If the event ``before_style_chart2_data_label_format_percent_get`` is cancelled and not handled.

        Returns:
            PercentFormat | None: Number percent style or ``None`` if cancelled.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.format.inner.direct.chart2.series.data_labels.data_labels.percent_format import PercentFormat

        if not isinstance(self, ChartDocPropPartial):
            raise TypeError("The class must be inherited from ChartDocPropPartial to use this method.")
        # pylint: disable=no-member
        doc = self.chart_doc.component
        comp = self.__component
        cargs = None
        if isinstance(self, EventsPartial):
            cargs = CancelEventArgs(self.style_numbers_percent_get.__qualname__)
            event_data: Dict[str, Any] = {
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_chart2_data_label_format_percent_get", cargs)
            if cargs.cancel is True:
                if cargs.handled is not False:
                    return None
                cargs.set("initial_event", "before_style_chart2_data_label_format_percent_get")
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style get has been cancelled.")
                else:
                    return None
            comp = cargs.event_data.get("this_component", comp)

        try:
            style = PercentFormat.from_obj(chart_doc=doc, obj=comp)
        except mEx.DisabledMethodError:
            return None

        style.set_update_obj(comp)
        return style


if mock_g.FULL_IMPORT:
    from ooodev.format.inner.direct.chart2.series.data_labels.data_labels.percent_format import PercentFormat
