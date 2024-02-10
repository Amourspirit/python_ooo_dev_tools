from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from ooodev.events.partial.events_partial import EventsPartial
from ooodev.loader import lo as mLo
from ooodev.format.inner.partial.factory_styler import FactoryStyler
from ooodev.format.inner import style_factory

if TYPE_CHECKING:
    from com.sun.star.chart2 import XChartDocument
    from ooo.dyn.lang.locale import Locale
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.chart2.numbers.numbers_t import NumbersT
    from ooo.dyn.util.number_format import NumberFormatEnum
    from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
else:
    XChartDocument = Any
    NumbersT = Any
    LoInst = Any
    NumberFormatEnum = Any
    NumberFormatIndexEnum = Any
    Locale = Any


class NumbersNumbersPartial:
    """
    Partial class for Numbers Numbers.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__styler = FactoryStyler(factory_name=factory_name, component=component, lo_inst=lo_inst)
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)
        self.__styler.after_event_name = "after_style_number_number"
        self.__styler.before_event_name = "before_style_number_number"

    def _NumbersNumbersPartial_get_chart_doc(self) -> XChartDocument:
        raise NotImplementedError

    def style_numbers_numbers(
        self,
        source_format: bool = True,
        num_format: NumberFormatEnum | int = 0,
        num_format_index: NumberFormatIndexEnum | int = -1,
        lang_locale: Locale | None = None,
    ) -> NumbersT | None:
        """
        Style Axis Line.

        Args:
            num_format (NumberFormatEnum, int, optional): Type of a number format.
                Use this to select a default format. Defaults to 0 (General Format).
                Only used if ``num_format_index`` is ``-1`` (omitted).
            num_format_index (NumberFormatIndexEnum, int, optional): Index of a number format.
                The enumeration values represent the built-in number formats. Defaults to ``-1``.
            lang_locale (Locale, optional): Locale of the number format. Defaults to ``None`` which used current Locale.

        Raises:
            CancelEventError: If the event ``after_style_number_number`` is cancelled and not handled.

        Returns:
            LinePropertiesT | None: Font Only instance or ``None`` if cancelled.
        """
        factory = style_factory.chart2_series_data_labels_numbers_factory
        kwargs = {
            "chart_doc": self._NumbersNumbersPartial_get_chart_doc(),
            "source_format": source_format,
            "num_format": num_format,
            "num_format_index": num_format_index,
            "lang_locale": lang_locale,
        }
        return self.__styler.style(factory=factory, **kwargs)
