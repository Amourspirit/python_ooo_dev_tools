from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING
import uno

from ooodev.events.partial.events_partial import EventsPartial
from ooodev.loader import lo as mLo
from ooodev.format.inner.partial.factory_styler import FactoryStyler
from ooodev.format.inner.style_factory import numbers_numbers_factory
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial

if TYPE_CHECKING:
    from ooo.dyn.lang.locale import Locale
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.calc.numbers.numbers_t import NumbersT
    from ooo.dyn.util.number_format import NumberFormatEnum
    from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
else:
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

    def style_numbers_numbers(
        self,
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
            CancelEventError: If the event ``before_style_number_number`` is cancelled and not handled.

        Returns:
            LinePropertiesT | None: Font Only instance or ``None`` if cancelled.

        Hint:
            - ``NumberFormatEnum`` can be imported from ``ooo.dyn.util.number_format``
            - ``NumberFormatIndexEnum`` can be imported from ``ooo.dyn.i18n.number_format_index``
            - ``Locale`` can be imported from ``ooo.dyn.lang.locale``
        """
        factory = numbers_numbers_factory
        kwargs = {
            "num_format": num_format,
            "num_format_index": num_format_index,
            "lang_locale": lang_locale,
            "component": self.__styler.component,
        }
        return self.__styler.style(factory=factory, **kwargs)

    def style_numbers_numbers_get(self) -> NumbersT | None:
        """
        Gets the Numbers Style.

        Raises:
            CancelEventError: If the event ``before_style_number_number_get`` is cancelled and not handled.

        Returns:
            TransparencyT | None: Area transparency style or ``None`` if cancelled.
        """
        if isinstance(self, CalcDocPropPartial):
            return self.__styler.style_get(
                factory=numbers_numbers_factory,
                component=self.calc_doc.component,
            )
        return self.__styler.style_get(factory=numbers_numbers_factory)

    def style_numbers_numbers_get_from_index(self, idx: int, locale: Locale | None = None) -> NumbersT | None:
        """
        Gets the Numbers Style.

        Raises:
            CancelEventError: If the event ``before_style_number_number_from_index`` is cancelled and not handled.

        Returns:
            TransparencyT | None: Area transparency style or ``None`` if cancelled.
        """
        kwargs: Dict[str, Any] = {"index": idx}
        if locale is not None:
            kwargs["lang_locale"] = locale
        if isinstance(self, CalcDocPropPartial):
            kwargs["component"] = self.calc_doc.component
        return self.__styler.style_get(
            factory=numbers_numbers_factory,
            call_method_name="from_index",
            event_name_suffix="_from_index",
            obj_arg_name="",
            **kwargs,
        )

    def style_numbers_numbers_get_from_str(
        self, nf_str: str, locale: Locale | None = None, auto_add: bool = False
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
            TransparencyT | None: Area transparency style or ``None`` if cancelled.
        """
        kwargs: Dict[str, Any] = {"nf_str": nf_str, "auto_add": auto_add}
        if locale is not None:
            kwargs["lang_locale"] = locale
        if isinstance(self, CalcDocPropPartial):
            kwargs["component"] = self.calc_doc.component
        return self.__styler.style_get(
            factory=numbers_numbers_factory,
            call_method_name="from_str",
            event_name_suffix="_from_index",
            **kwargs,
        )
