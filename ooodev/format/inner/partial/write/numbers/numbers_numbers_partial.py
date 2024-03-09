from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING
import uno

from ooodev.events.partial.events_partial import EventsPartial
from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler
from ooodev.format.inner.style_factory import numbers_numbers_factory
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial

if TYPE_CHECKING:
    from ooo.dyn.lang.locale import Locale
    from ooo.dyn.util.number_format import NumberFormatEnum
    from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.calc.numbers.numbers_t import NumbersT
    from ooodev.events.args.cancel_event_args import CancelEventArgs
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
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_number_number",
            after_event="after_style_number_number",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

    def style_numbers_numbers(
        self,
        num_format: NumberFormatEnum | int = 0,
        num_format_index: NumberFormatIndexEnum | int = -1,
        lang_locale: Locale | None = None,
    ) -> NumbersT | None:
        """
        Style numbers numbers.

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
            NumbersT | None: Style Numbers instance or ``None`` if cancelled.

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

    def __get_style_numbers(self) -> NumbersT | None:
        def on_style(src: Any, events: CancelEventArgs) -> None:
            # this will stop the style from being applied.
            # not critical but style is being applied later via style.empty.update()
            events.event_data["cancel_apply"] = True

        styler = self.__styler
        factory = numbers_numbers_factory
        if isinstance(self, EventsPartial):
            self.subscribe_event(event_name=styler.before_event_name, callback=on_style)
        return styler.style(factory=factory)

    def style_numbers_general(self) -> NumbersT | None:
        """
        Style numbers general.
        """

        style = self.__get_style_numbers()
        if style is None:
            return None
        new_style = style.number
        if not new_style.has_update_obj():
            new_style.set_update_obj(style.get_update_obj())
        new_style.update()
        return new_style

    def style_numbers_currency(self) -> NumbersT | None:
        """
        Style numbers currency.
        """

        style = self.__get_style_numbers()
        if style is None:
            return None
        new_style = style.currency
        if not new_style.has_update_obj():
            new_style.set_update_obj(style.get_update_obj())
        new_style.update()
        return new_style

    def style_numbers_date(self) -> NumbersT | None:
        """
        Style numbers date.
        """

        style = self.__get_style_numbers()
        if style is None:
            return None
        new_style = style.date
        if not new_style.has_update_obj():
            new_style.set_update_obj(style.get_update_obj())
        new_style.update()
        return new_style

    def style_numbers_time(self) -> NumbersT | None:
        """
        Style numbers time.
        """

        style = self.__get_style_numbers()
        if style is None:
            return None
        new_style = style.time
        if not new_style.has_update_obj():
            new_style.set_update_obj(style.get_update_obj())
        new_style.update()
        return new_style

    def style_numbers_number(self) -> NumbersT | None:
        """
        Style numbers number.
        """

        style = self.__get_style_numbers()
        if style is None:
            return None
        new_style = style.number
        if not new_style.has_update_obj():
            new_style.set_update_obj(style.get_update_obj())
        new_style.update()
        return new_style

    def style_numbers_scientific(self) -> NumbersT | None:
        """
        Style numbers scientific.
        """

        style = self.__get_style_numbers()
        if style is None:
            return None
        new_style = style.scientific
        if not new_style.has_update_obj():
            new_style.set_update_obj(style.get_update_obj())
        new_style.update()
        return new_style

    def style_numbers_fraction(self) -> NumbersT | None:
        """
        Style numbers fraction.
        """

        style = self.__get_style_numbers()
        if style is None:
            return None
        new_style = style.fraction
        if not new_style.has_update_obj():
            new_style.set_update_obj(style.get_update_obj())
        new_style.update()
        return new_style

    def style_numbers_percent(self) -> NumbersT | None:
        """
        Style numbers percent.
        """

        style = self.__get_style_numbers()
        if style is None:
            return None
        new_style = style.percent
        if not new_style.has_update_obj():
            new_style.set_update_obj(style.get_update_obj())
        new_style.update()
        return new_style

    def style_numbers_datetime(self) -> NumbersT | None:
        """
        Style numbers datetime.
        """

        style = self.__get_style_numbers()
        if style is None:
            return None
        new_style = style.datetime
        if not new_style.has_update_obj():
            new_style.set_update_obj(style.get_update_obj())
        new_style.update()
        return new_style

    def style_numbers_boolean(self) -> NumbersT | None:
        """
        Style numbers boolean.
        """

        style = self.__get_style_numbers()
        if style is None:
            return None
        new_style = style.boolean
        if not new_style.has_update_obj():
            new_style.set_update_obj(style.get_update_obj())
        new_style.update()
        return new_style

    def style_numbers_numbers_get(self) -> NumbersT | None:
        """
        Gets the Numbers Style.

        Raises:
            CancelEventError: If the event ``before_style_number_number_get`` is cancelled and not handled.

        Returns:
            NumbersT | None: Numbers style or ``None`` if cancelled.
        """
        styler = self.__styler
        if isinstance(self, WriteDocPropPartial):
            return styler.style_get(
                factory=numbers_numbers_factory,
                component=self.write_doc.component,
            )
        return styler.style_get(factory=numbers_numbers_factory)

    def style_numbers_numbers_get_from_index(self, idx: int, locale: Locale | None = None) -> NumbersT | None:
        """
        Gets the Numbers Style.

        Raises:
            CancelEventError: If the event ``before_style_number_number_from_index`` is cancelled and not handled.

        Returns:
            NumbersT | None: Numbers style or ``None`` if cancelled.

        Hint:
            - ``Locale`` can be imported from ``ooo.dyn.lang.locale``
        """
        styler = self.__styler
        kwargs: Dict[str, Any] = {"index": idx}
        if locale is not None:
            kwargs["lang_locale"] = locale
        if isinstance(self, WriteDocPropPartial):
            kwargs["component"] = self.write_doc.component
        return styler.style_get(
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
            NumbersT | None: Numbers style or ``None`` if cancelled.

        Hint:
            - ``Locale`` can be imported from ``ooo.dyn.lang.locale``
        """
        styler = self.__styler
        kwargs: Dict[str, Any] = {"nf_str": nf_str, "auto_add": auto_add}
        if locale is not None:
            kwargs["lang_locale"] = locale
        if isinstance(self, WriteDocPropPartial):
            kwargs["component"] = self.write_doc.component
        return styler.style_get(
            factory=numbers_numbers_factory,
            call_method_name="from_str",
            event_name_suffix="_from_index",
            **kwargs,
        )
