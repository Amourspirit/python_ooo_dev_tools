from __future__ import annotations
import contextlib
from typing import Any, Dict, Tuple, overload
import uno
from com.sun.star.chart2 import XChartDocument

from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
from ooo.dyn.lang.locale import Locale
from ooo.dyn.util.number_format import NumberFormatEnum
from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs

from ooodev.exceptions import ex as mEx
from ooodev.format.inner.direct.chart2.chart.numbers.numbers import Numbers as ChartNumbers
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps


class PercentFormat(ChartNumbers):
    """
    Chart Data Series, Data Labels Percent Format.

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

    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "PercentageNumberFormat"
        return self._property_name

    # region apply()
    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs: Any) -> None: ...

    def apply(self, obj: Any, **kwargs: Any) -> None:
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
        override_dv = {"LinkNumberFormatToSource": self._source_format}
        props: Dict[str, Any] = {"verify": False, "override_dv": override_dv}
        if self._source_format:
            # if source format is true then the number format should be set to None.
            # This causes the chart to use the source format.
            props["format_key"] = None

        super().apply(obj, **props)

    # endregion apply()

    # region Copy()
    @overload
    def copy(self) -> PercentFormat: ...

    @overload
    def copy(self, **kwargs) -> PercentFormat: ...

    def copy(self, **kwargs) -> PercentFormat:
        """
        Creates a copy of the instance.

        Returns:
            Numbers: Copy of the instance.
        """
        inst = super().copy(**kwargs)
        inst._source_format = self._source_format
        return inst

    # endregion Copy()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: Any) -> PercentFormat: ...

    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: Any, **kwargs) -> PercentFormat: ...

    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: Any, **kwargs) -> PercentFormat:
        """
        Gets instance from object

        Args:
            chart_doc (XChartDocument): Chart document.
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            PercentFormat: Instance that represents numbers format.
        """
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
    ) -> PercentFormat:
        """
        Gets instance from format string

        Args:
            chart_doc (XChartDocument): Chart document.
            nf_str (str): Format string.
            lang_locale (Locale, optional): Locale. Defaults to ``None``.
            auto_add (bool, optional): If ``True``, format string will be added to document if not found. Defaults to ``False``.

        Keyword Args:
            source_format (bool, optional): If ``True``, the number format will be linked to the source format. Defaults to ``False``.

        Returns:
            PercentFormat: Instance that represents numbers format.
        """
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
    ) -> PercentFormat:
        """
        Gets instance from number format index. This is the index that is assigned to the ``PercentageNumberFormat`` property of an object such as a cell.

        Args:
            chart_doc (XChartDocument): Chart document.
            index (int): Format (``PercentageNumberFormat``) index.
            lang_locale (Locale, optional): Locale. Defaults to ``None``.

        Keyword Args:
            source_format (bool, optional): If ``True``, the number format will be linked to the source format. Defaults to ``False``.

        Returns:
            PercentFormat: Instance that represents numbers format.
        """
        source_format = kwargs.pop("source_format", False)
        inst = cls(chart_doc=chart_doc, source_format=source_format, lang_locale=lang_locale, **kwargs)
        inst._format_key_prop = index
        return inst

    # endregion from index

    # endregion Overrides

    # region on events
    def on_property_set_error(self, source: Any, event_args: KeyValCancelArgs) -> None:
        if event_args.key == self._get_property_name() and event_args.value is None:
            with contextlib.suppress(Exception):
                # in theory this should work, but it doesn't.
                # As of LibreOffice 7.5 it still throws an exception.
                # This seems to be due to XPropertySet not being properly implemented on data points and data series.
                # currently this event is raised only for data point and not for data series.
                # Either way, the error is caught and only printed to console.
                mProps.Props.set_default(event_args.event_data, self._get_property_name())
                event_args.handled = True
        super().on_property_set_error(source, event_args)

    # endregion on events
