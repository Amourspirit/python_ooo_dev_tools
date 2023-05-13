from __future__ import annotations
from typing import Tuple
import uno
from com.sun.star.chart2 import XChartDocument
from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
from ooo.dyn.lang.locale import Locale
from ooo.dyn.util.number_format import NumberFormatEnum


from ooodev.format.inner.direct.chart2.series.data_labels.data_labels.number_format import (
    NumberFormat as DataLabelsNumberFormat,
)


class Numbers(DataLabelsNumberFormat):
    """
    Chart Axis Numbers format.

    .. seealso::

        - :ref:`help_chart2_format_direct_axis_numbers`

    .. versionadded:: 0.9.4
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
            - :ref:`help_chart2_format_direct_axis_numbers`
            - `API NumberFormat <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1util_1_1NumberFormat.html>`__
            - `API NumberFormatIndex <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1i18n_1_1NumberFormatIndex.html>`__
        """
        super().__init__(
            chart_doc=chart_doc,
            source_format=source_format,
            num_format=num_format,
            num_format_index=num_format_index,
            lang_locale=lang_locale,
        )

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.chart2.Axis",)
        return self._supported_services_values
