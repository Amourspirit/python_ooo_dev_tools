from __future__ import annotations
from typing import Type, Tuple
import uno

from ooodev.format.inner.direct.structs.data_point_label_struct import DataPointLabelStruct
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleMulti


class TextAttribs(StyleMulti):
    """
    Chart Data Series, Data Labels Text Attributes.

    Any properties starting with ``prop_`` set or get current instance values.

    .. seealso::

        - :ref:`help_chart2_format_direct_series_labels_data_labels`
    """

    def __init__(
        self,
        show_number: bool = False,
        show_number_in_percent: bool = False,
        show_category_name: bool = False,
        show_legend_symbol: bool = False,
        show_custom_label: bool = False,
        show_series_name: bool = False,
        auto_text_wrap: bool | None = None,
    ) -> None:
        """
        Constructor.

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

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_series_labels_data_labels`
        """
        super().__init__()
        if auto_text_wrap is not None:
            self.prop_auto_text_wrap = auto_text_wrap

        struct = self._get_struct_type()(
            show_number=show_number,
            show_number_in_percent=show_number_in_percent,
            show_category_name=show_category_name,
            show_legend_symbol=show_legend_symbol,
            show_custom_label=show_custom_label,
            show_series_name=show_series_name,
        )
        self._set_style("dp_struct", struct)  # type: ignore

    # region Internal Methods
    def _get_struct_type(self) -> Type[DataPointLabelStruct]:
        return DataPointLabelStruct

    # endregion Internal Methods

    # region Overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.chart2.DataSeries",
                "com.sun.star.chart2.DataPointProperties",
            )
        return self._supported_services_values

    def copy(self, **kwargs) -> TextAttribs:
        """Copies the style"""
        return self.__class__(
            show_number=self.prop_show_number,
            show_number_in_percent=self.prop_show_number_in_percent,
            show_category_name=self.prop_show_category_name,
            show_legend_symbol=self.prop_show_legend_symbol,
            show_custom_label=self.prop_show_custom_label,
            show_series_name=self.prop_show_series_name,
            auto_text_wrap=self.prop_auto_text_wrap,
            **kwargs,
        )

    # endregion Overrides

    # region style properties
    @property
    def show_number(self) -> TextAttribs:
        """Gets a copy of the style with ``prop_show_number`` set to ``True``"""
        cp = self.copy()
        cp.prop_show_number = True
        return cp

    @property
    def show_number_in_percent(self) -> TextAttribs:
        """Gets a copy of the style with ``prop_show_number_in_percent`` set to ``True``"""
        cp = self.copy()
        cp.prop_show_number_in_percent = True
        return cp

    @property
    def show_category_name(self) -> TextAttribs:
        """Gets a copy of the style with ``prop_show_category_name`` set to ``True``"""
        cp = self.copy()
        cp.prop_show_category_name = True
        return cp

    @property
    def show_legend_symbol(self) -> TextAttribs:
        """Gets a copy of the style with ``prop_show_legend_symbol`` set to ``True``"""
        cp = self.copy()
        cp.prop_show_legend_symbol = True
        return cp

    @property
    def show_custom_label(self) -> TextAttribs:
        """Gets a copy of the style with ``prop_show_custom_label`` set to ``True``"""
        cp = self.copy()
        cp.prop_show_custom_label = True
        return cp

    @property
    def show_series_name(self) -> TextAttribs:
        """Gets a copy of the style with ``prop_show_series_name`` set to ``True``"""
        cp = self.copy()
        cp.prop_show_series_name = True
        return cp

    @property
    def auto_text_wrap(self) -> TextAttribs:
        """Gets a copy of the style with ``prop_auto_text_wrap`` set to ``True``"""
        cp = self.copy()
        cp.prop_auto_text_wrap = True
        return cp

    # endregion style properties

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.UNKNOWN
        return self._format_kind_prop

    @property
    def _dp_struct(self) -> DataPointLabelStruct:
        """Gets the struct of the style"""
        try:
            return self._dp_struct_prop  # type: ignore
        except AttributeError:
            self._dp_struct_prop = self._get_style("dp_struct")
        return self._dp_struct_prop  # type: ignore

    @property
    def prop_auto_text_wrap(self) -> bool | None:
        """Gets or set if the text is automatically wrapped, if the text is too long to fit in the available space."""
        return self._get("TextWordWrap")

    @prop_auto_text_wrap.setter
    def prop_auto_text_wrap(self, value: bool | None) -> None:
        if value is None:
            self._remove("TextWordWrap")
            return
        self._set("TextWordWrap", bool(value))

    # region Data Point Label Struct Properties
    @property
    def prop_show_number(self) -> bool:
        """Gets or set if the number is additionally displayed in the caption."""
        return self._dp_struct.prop_show_number

    @prop_show_number.setter
    def prop_show_number(self, value: bool) -> None:
        self._dp_struct.prop_show_number = value

    @property
    def prop_show_number_in_percent(self) -> bool:
        """
        Only effective, if ``ShowNumber`` is ``True``.

        If this member is also ``True``, the numbers are displayed as percentages of a category
        """
        return self._dp_struct.prop_show_number_in_percent

    @prop_show_number_in_percent.setter
    def prop_show_number_in_percent(self, value: bool) -> None:
        self._dp_struct.prop_show_number_in_percent = value

    @property
    def prop_show_category_name(self) -> bool:
        """
        Gets or set if the caption contains the category name of the category to which a data point belongs
        """
        return self._dp_struct.prop_show_number_in_percent

    @prop_show_category_name.setter
    def prop_show_category_name(self, value: bool) -> None:
        self._dp_struct.prop_show_number_in_percent = value

    @property
    def prop_show_legend_symbol(self) -> bool:
        """
        Gets or set if the legend symbol is additionally displayed in the caption.
        """
        return self._dp_struct.prop_show_legend_symbol

    @prop_show_legend_symbol.setter
    def prop_show_legend_symbol(self, value: bool) -> None:
        self._dp_struct.prop_show_legend_symbol = value

    @property
    def prop_show_custom_label(self) -> bool:
        """
        Gets or set if a custom label is additionally displayed in the caption.

        Since LibreOffice ``7.1``
        """
        return self._dp_struct.prop_show_custom_label

    @prop_show_custom_label.setter
    def prop_show_custom_label(self, value: bool) -> None:
        self._dp_struct.prop_show_custom_label = value

    @property
    def prop_show_series_name(self) -> bool:
        """
        Gets or set if the caption contains the name of the series to which a data point belongs.

        Since LibreOffice ``7.2``
        """
        return self._dp_struct.prop_show_series_name

    @prop_show_series_name.setter
    def prop_show_series_name(self, value: bool) -> None:
        self._dp_struct.prop_show_series_name = value

    # endregion Data Point Label Struct Properties

    # endregion Properties
