from __future__ import annotations
from typing import Any, Tuple, Type, TypeVar, overload, TYPE_CHECKING
import uno

from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
from ooo.dyn.lang.locale import Locale
from ooo.dyn.util.number_format import NumberFormatEnum

from ooodev.format.inner.direct.calc.numbers.numbers import Numbers as CalcNumbers
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.exceptions import ex as mEx
from ooodev.utils import props as mProps

if TYPE_CHECKING:
    from com.sun.star.chart2 import XChartDocument
else:
    XChartDocument = Any

_TNumbers = TypeVar(name="_TNumbers", bound="Numbers")


class Numbers(CalcNumbers):
    """
    Chart Numbers format.

    .. versionadded:: 0.9.4
    """

    def __init__(
        self,
        chart_doc: XChartDocument,
        *,
        num_format: NumberFormatEnum | int = 0,
        num_format_index: NumberFormatIndexEnum | int = -1,
        lang_locale: Locale | None = None,
    ) -> None:
        """
        Constructor

        Args:
            chart_doc (XChartDocument): Chart document.
            num_format (NumberFormatEnum, int, optional): Type of a number format.
                Use this to select a default format. Defaults to 0 (General Format).
                Only used if ``num_format_index`` is ``-1`` (omitted).
            num_format_index (NumberFormatIndexEnum | int, optional): Index of a number format.
                The enumeration values represent the built-in number formats. Defaults to ``-1``.
            lang_locale (Locale, optional): Locale of the number format. Defaults to ``None`` which used current Locale.
        """
        self._chart_doc = chart_doc
        super().__init__(
            num_format=num_format,
            num_format_index=num_format_index,
            lang_locale=lang_locale,
            component=self._chart_doc,
        )

    # region Overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.chart2.RegressionEquation", "com.sun.star.chart2.Axis")
        return self._supported_services_values

    # region Copy()
    @overload
    def copy(self: _TNumbers) -> _TNumbers: ...

    @overload
    def copy(self: _TNumbers, **kwargs) -> _TNumbers: ...

    def copy(self: _TNumbers, **kwargs) -> _TNumbers:
        """
        Creates a copy of the instance.

        Returns:
            Numbers: Copy of the instance.
        """
        inst = self.__class__(chart_doc=self._chart_doc, **kwargs)
        inst._format_key_prop = self.prop_format_key
        return inst

    # endregion Copy()
    # endregion Overrides

    # region Static Methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TNumbers], chart_doc: XChartDocument, obj: object) -> _TNumbers: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TNumbers], chart_doc: XChartDocument, obj: object, **kwargs) -> _TNumbers: ...

    @classmethod
    def from_obj(cls: Type[_TNumbers], chart_doc: XChartDocument, obj: object, **kwargs) -> _TNumbers:
        """
        Gets instance from object

        Args:
            chart_doc (XChartDocument): Chart document.
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Numbers: Instance that represents numbers format.
        """
        nu = cls(chart_doc=chart_doc, **kwargs)
        if not nu._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        # Get the number format index key of the cell's properties
        nf = int(mProps.Props.get(obj, nu._get_property_name(), -1))
        if nf == -1:
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        locale = mProps.Props.get(obj, "CharLocale", None)
        inst = cls(chart_doc=chart_doc, lang_locale=locale, **kwargs)
        inst._format_key_prop = nf
        return inst

    # endregion from_obj()
    # region from string
    @classmethod
    def from_str(
        cls: Type[_TNumbers],
        chart_doc: XChartDocument,
        nf_str: str,
        lang_locale: Locale | None = None,
        auto_add: bool = False,
        **kwargs,
    ) -> _TNumbers:
        """
        Gets instance from format string

        Args:
            chart_doc (XChartDocument): Chart document.
            nf_str (str): Format string.
            lang_locale (Locale, optional): Locale. Defaults to ``None``.
            auto_add (bool, optional): If True, format string will be added to document if not found. Defaults to ``False``.

        Returns:
            Numbers: Instance that represents numbers format.
        """
        num_calc = CalcNumbers.from_str(
            nf_str=nf_str, lang_locale=lang_locale, auto_add=auto_add, component=chart_doc, **kwargs
        )

        inst = cls(chart_doc=chart_doc, **kwargs)

        inst._format_key_prop = num_calc.prop_format_key
        return inst

    # endregion from string

    # region from index
    @classmethod
    def from_index(
        cls: Type[_TNumbers], chart_doc: XChartDocument, index: int, lang_locale: Locale | None = None, **kwargs
    ) -> _TNumbers:
        """
        Gets instance from number format index. This is the index that is assigned to the ``NumberFormat`` property of an object such as a cell.

        Args:
            chart_doc (XChartDocument): Chart document.
            index (int): Format (``NumberFormat``) index.
            lang_locale (Locale, optional): Locale. Defaults to ``None``.


        Returns:
            Numbers: Instance that represents numbers format.
        """
        inst = cls(chart_doc=chart_doc, lang_locale=lang_locale, **kwargs)
        inst._format_key_prop = index
        return inst

    # endregion from index

    # endregion Static Methods

    # region Properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.UNKNOWN
        return self._format_kind_prop

    # endregion Properties
