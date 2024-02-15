from __future__ import annotations
from typing import Any, Dict, Tuple, overload
import uno
from com.sun.star.chart2 import XChartDocument

from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
from ooo.dyn.lang.locale import Locale
from ooo.dyn.util.number_format import NumberFormatEnum

from ooodev.exceptions import ex as mEx
from ooodev.format.inner.direct.chart2.chart.numbers.numbers import Numbers as ChartNumbers
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps


class NumberFormat(ChartNumbers):
    """
    Chart Data Series, Data Labels Number Format.

    .. seealso::

        - :ref:`help_chart2_format_direct_series_labels_data_labels`
    """

    def __init__(
        self,
        chart_doc: XChartDocument,
        *,
        source_format: bool = True,
        num_format: NumberFormatEnum | int = 0,
        num_format_index: NumberFormatIndexEnum | int = -1,
        lang_locale: Locale | None = None,
    ) -> None:
        """
        Constructor

        Args:
            chart_doc (XChartDocument): Chart document.
            source_format (bool, optional): Specifies whether the number format should be linked to the source format. Defaults to ``True``.
            num_format (NumberFormatEnum, int, optional): specifies the number format. Defaults to ``0``.
            num_format_index (NumberFormatIndexEnum | int, optional): Specifies the number format index. Defaults to ``-1``.
            lang_locale (Locale, optional): Specifies the language locale. Defaults to ``None``.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_series_labels_data_labels`
            - `API NumberFormat <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1util_1_1NumberFormat.html>`__
            - `API NumberFormatIndex <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1i18n_1_1NumberFormatIndex.html>`__
        """
        # LinkNumberFormatToSource
        super().__init__(
            chart_doc=chart_doc, num_format=num_format, num_format_index=num_format_index, lang_locale=lang_locale
        )
        self._source_format = source_format

    # region Overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.chart2.DataPoint", "com.sun.star.chart2.DataSeries")
        return self._supported_services_values

    # region apply()
    @overload
    def apply(self, obj: object) -> None: ...

    @overload
    def apply(self, obj: object, **kwargs: Any) -> None: ...

    def apply(self, obj: object, **kwargs: Any) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO Object that styles are to be applied.
            kwargs (Any, optional): Expandable list of key value pairs that may be used in child classes.

        Keyword Args:
            format_key (Any, optional): NumberFormat key, overrides ``prop_format_key`` property value.

        Returns:
            None:
        """
        # obj is expected to be a data series or data point
        if not self._is_valid_obj(obj):
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property. Invalid object")
            return
        dv = kwargs.pop("override_dv", {})
        props: Dict[str, Any] = {"verify": False}

        if self._source_format:
            # if source format is true then the number format should be set to None.
            # This causes the chart to use the source format.
            props["format_key"] = None
            dv["LinkNumberFormatToSource"] = True
        else:
            dv["LinkNumberFormatToSource"] = False
        props["override_dv"] = dv

        super().apply(obj, **props)

    # endregion apply()

    # region Copy()
    @overload
    def copy(self) -> NumberFormat: ...

    @overload
    def copy(self, **kwargs) -> NumberFormat: ...

    def copy(self, **kwargs) -> NumberFormat:
        """
        Creates a copy of the instance.

        Returns:
            Numbers: Copy of the instance.
        """
        # pylint: disable=protected-access
        inst = super().copy(**kwargs)
        inst._source_format = self._source_format
        return inst

    # endregion Copy()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: object) -> NumberFormat: ...

    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: object, **kwargs) -> NumberFormat: ...

    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: object, **kwargs) -> NumberFormat:
        """
        Gets instance from object

        Args:
            chart_doc (XChartDocument): Chart document.
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            NumberFormat: Instance that represents numbers format.
        """
        # pylint: disable=protected-access
        nu = cls(chart_doc=chart_doc, source_format=False, **kwargs)
        if not nu._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        # Get the number format index key of the cell's properties
        nf = int(mProps.Props.get(obj, nu._get_property_name(), -1))
        if nf == -1:
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        locale = mProps.Props.get(obj, "CharLocale", None)
        src_format = bool(mProps.Props.get(obj, "LinkNumberFormatToSource", True))
        inst = cls(chart_doc=chart_doc, source_format=src_format, lang_locale=locale, **kwargs)
        inst._format_key_prop = nf
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()

    # region from string
    @classmethod
    def from_str(
        cls,
        chart_doc: XChartDocument,
        nf_str: str,
        lang_locale: Locale | None = None,
        auto_add: bool = False,
        **kwargs,
    ) -> NumberFormat:
        """
        Gets instance from format string

        Args:
            chart_doc (XChartDocument): Chart document.
            nf_str (str): Format string.
            lang_locale (Locale, optional): Locale. Defaults to ``None``.
            auto_add (bool, optional): If True, format string will be added to document if not found. Defaults to ``False``.

        Keyword Args:
            source_format (bool, optional): If ``True``, the number format will be linked to the source format. Defaults to ``False``.

        Returns:
            NumberFormat: Instance that represents numbers format.
        """
        # pylint: disable=protected-access
        num_chart = ChartNumbers.from_str(
            chart_doc=chart_doc, nf_str=nf_str, lang_locale=lang_locale, auto_add=auto_add, **kwargs
        )
        source_format = kwargs.pop("source_format", False)

        inst = cls(chart_doc=chart_doc, source_format=source_format, **kwargs)

        inst._format_key_prop = num_chart.prop_format_key
        return inst

    # endregion from string

    # region from index
    @classmethod
    def from_index(
        cls, chart_doc: XChartDocument, index: int, lang_locale: Locale | None = None, **kwargs
    ) -> NumberFormat:
        """
        Gets instance from number format index. This is the index that is assigned to the ``NumberFormat`` property of an object such as a cell.

        Args:
            chart_doc (XChartDocument): Chart document.
            index (int): Format (``NumberFormat``) index.
            lang_locale (Locale, optional): Locale. Defaults to ``None``.

        Keyword Args:
            source_format (bool, optional): If ``True``, the number format will be linked to the source format. Defaults to ``False``.

        Returns:
            NumberFormat: Instance that represents numbers format.
        """
        # pylint: disable=protected-access
        source_format = kwargs.pop("source_format", False)
        inst = cls(chart_doc=chart_doc, source_format=source_format, lang_locale=lang_locale, **kwargs)
        inst._format_key_prop = index
        return inst

    # endregion from index

    # endregion Overrides

    # region Methods
    def _source_format_prop_name(self) -> str:
        return "LinkNumberFormatToSource"

    # endregion Methods
