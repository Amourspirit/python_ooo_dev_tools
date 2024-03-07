from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING
import uno

from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler
from ooodev.format.inner import style_factory
from ooodev.calc.chart2.partial.chart_doc_prop_partial import ChartDocPropPartial
from ooodev.events.partial.events_partial import EventsPartial

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
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_number_number",
            after_event="after_style_number_number",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

    def _NumbersNumbersPartial_get_chart_doc(self) -> XChartDocument:
        if isinstance(self, ChartDocPropPartial):
            return self.chart_doc.component
        raise NotImplementedError

    def style_numbers_numbers(
        self,
        source_format: bool = True,
        num_format: NumberFormatEnum | int = 0,
        num_format_index: NumberFormatIndexEnum | int = -1,
        lang_locale: Locale | None = None,
    ) -> NumbersT | None:
        """
        Style Numbers.

        Args:
            num_format (NumberFormatEnum, int, optional): Type of a number format.
                Use this to select a default format. Defaults to 0 (General Format).
                Only used if ``num_format_index`` is ``-1`` (omitted).
            num_format_index (NumberFormatIndexEnum, int, optional): Index of a number format.
                The enumeration values represent the built-in number formats. Defaults to ``-1``.
            lang_locale (Locale, optional): Locale of the number format. Defaults to ``None`` which used current Locale.

        Raises:
            CancelEventError: If the event ``before_style_number_number`` is cancelled and not handled.

        Returns:
            NumbersT | None: Number Style instance or ``None`` if cancelled.

        Hint:
            - ``Locale`` can be imported from ``ooo.dyn.lang.locale``
            - ``NumberFormatEnum`` can be imported from ``ooo.dyn.util.number_format``
            - ``NumberFormatIndexEnum`` can be imported from ``ooo.dyn.i18n.number_format_index``
        """
        factory = style_factory.chart2_series_data_labels_numbers_factory
        kwargs = {
            "chart_doc": self._NumbersNumbersPartial_get_chart_doc(),
            "source_format": source_format,
        }
        if lang_locale is not None:
            kwargs["lang_locale"] = lang_locale
        if num_format_index != -1:
            kwargs["num_format_index"] = num_format_index
        if num_format != 0:
            kwargs["num_format"] = num_format
        return self.__styler.style(factory=factory, **kwargs)

    def style_numbers_numbers_get(self) -> NumbersT | None:
        """
        Gets the Numbers Style.

        Raises:
            CancelEventError: If the event ``before_style_number_number_get`` is cancelled and not handled.

        Returns:
            NumbersT | None: Numbers style or ``None`` if cancelled.
        """
        return self.__styler.style_get(
            factory=style_factory.chart2_series_data_labels_numbers_factory,
            chart_doc=self._NumbersNumbersPartial_get_chart_doc(),
        )

    def style_numbers_numbers_get_from_index(self, idx: int, locale: Locale | None = None) -> NumbersT | None:
        """
        Gets the Numbers Style.

        Raises:
            CancelEventError: If the event ``before_style_number_number_from_index`` is cancelled and not handled.

        Returns:
            NumbersT | None: Numbers style or ``None`` if cancelled.
        """
        styler = self.__styler
        kwargs: Dict[str, Any] = {"index": idx}
        if locale is not None:
            kwargs["lang_locale"] = locale
        return styler.style_get(
            factory=style_factory.chart2_series_data_labels_numbers_factory,
            chart_doc=self._NumbersNumbersPartial_get_chart_doc(),
            call_method_name="from_index",
            event_name_suffix="_from_index",
            obj_arg_name="",
            **kwargs,
        )

    def style_numbers_numbers_get_from_str(
        self, nf_str: str, locale: Locale | None = None, auto_add: bool = False, source_format: bool = False
    ) -> NumbersT | None:
        """
        Gets the Numbers Style.

        Args:
            nf_str (str): Format string.
            lang_locale (Locale, optional): Locale. Defaults to ``None``.
            auto_add (bool, optional): If True, format string will be added to document if not found. Defaults to ``False``.
            source_format (bool, optional): If ``True``, the number format will be linked to the source format. Defaults to ``False``.

        Raises:
            CancelEventError: If the event ``before_style_number_number_from_index`` is cancelled and not handled.

        Returns:
            NumbersT | None: Number style or ``None`` if cancelled.
        """
        styler = self.__styler
        kwargs: Dict[str, Any] = {"nf_str": nf_str, "auto_add": auto_add, "source_format": source_format}
        if locale is not None:
            kwargs["lang_locale"] = locale
        return styler.style_get(
            factory=style_factory.chart2_series_data_labels_numbers_factory,
            chart_doc=self._NumbersNumbersPartial_get_chart_doc(),
            call_method_name="from_str",
            event_name_suffix="_from_index",
            **kwargs,
        )
