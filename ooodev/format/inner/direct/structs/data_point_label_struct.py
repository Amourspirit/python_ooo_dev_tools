# region Import
from __future__ import annotations
import contextlib
from typing import Any, Tuple, Type, cast, overload, TypeVar
import uno
from ooo.dyn.chart2.data_point_label import DataPointLabel

from ooodev.exceptions import ex as mEx
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.common.props.struct_data_point_label_props import StructDataPointLabelProps
from ooodev.format.inner.direct.structs.struct_base import StructBase

# endregion Import

_TDataPointLabelStruct = TypeVar(name="_TDataPointLabelStruct", bound="DataPointLabelStruct")


class DataPointLabelStruct(StructBase):
    """
    DataPointLabel struct.

    .. versionadded:: 0.9.4
    """

    # region init

    def __init__(
        self,
        show_number: bool = False,
        show_number_in_percent: bool = False,
        show_category_name: bool = False,
        show_legend_symbol: bool = False,
        show_custom_label: bool = False,
        show_series_name: bool = False,
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
        """
        super().__init__()
        self.prop_show_number = show_number
        self.prop_show_number_in_percent = show_number_in_percent
        self.prop_show_category_name = show_category_name
        self.prop_show_legend_symbol = show_legend_symbol
        self.prop_show_custom_label = show_custom_label
        self.prop_show_series_name = show_series_name

    # endregion init

    # region internal methods
    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "Label"
        return self._property_name

    # endregion internal methods

    # region dunder methods
    def __eq__(self, oth: object) -> bool:
        obj2 = None
        if isinstance(oth, DataPointLabelStruct):
            obj2 = oth.get_uno_struct()
        if getattr(oth, "typeName", None) == "com.sun.star.chart2.DataPointLabel":
            obj2 = cast(DataPointLabel, oth)
        if obj2:
            obj1 = self.get_uno_struct()
            for prop in self._props:
                with contextlib.suppress(AttributeError):
                    if getattr(obj1, prop) != getattr(obj2, prop):
                        return False
            return True
        return NotImplemented

    # endregion dunder methods

    # region methods
    def get_uno_struct(self) -> DataPointLabel:
        """
        Gets UNO ``DataPointLabel`` from instance.

        Returns:
            DataPointLabel: ``DataPointLabel`` instance
        """
        inst = DataPointLabel()
        for prop in self._props:
            with contextlib.suppress(AttributeError):
                setattr(inst, prop, self._get(prop))
        return inst

    # endregion methods

    # region overrides methods

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.chart2.DataSeries",
                "com.sun.star.chart2.DataPointProperties",
            )
        return self._supported_services_values

    # region apply()

    @overload
    def apply(self, obj: Any) -> None:  # type: ignore
        ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies tab properties to ``obj``

        Args:
            obj (object): UNO object.

        Returns:
            None:
        """
        name = self._get_property_name()
        if not name:
            return
        if not mProps.Props.has(obj, name):
            self._print_not_valid_srv("apply")
            return

        struct = self.get_uno_struct()
        props = {name: struct}
        super().apply(obj=obj, override_dv=props)

    # endregion apply()

    # endregion overrides methods

    # region static methods

    # region from_uno_struct()
    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TDataPointLabelStruct], value: DataPointLabel) -> _TDataPointLabelStruct: ...

    @overload
    @classmethod
    def from_uno_struct(
        cls: Type[_TDataPointLabelStruct], value: DataPointLabel, **kwargs
    ) -> _TDataPointLabelStruct: ...

    @classmethod
    def from_uno_struct(cls: Type[_TDataPointLabelStruct], value: DataPointLabel, **kwargs) -> _TDataPointLabelStruct:
        """
        Converts a ``DataPointLabel`` instance to a ``PointStruct``.

        Args:
            value (DataPointLabel): UNO ``DataPointLabel``.

        Returns:
            DataPointLabelStruct: ``PointStruct`` set with ``DataPointLabel`` properties.
        """
        inst = cls(**kwargs)
        for prop in inst._props:
            with contextlib.suppress(AttributeError):
                inst._set(prop, getattr(value, prop))
        return inst

    # endregion from_uno_struct()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TDataPointLabelStruct], obj: Any) -> _TDataPointLabelStruct: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TDataPointLabelStruct], obj: Any, **kwargs) -> _TDataPointLabelStruct: ...

    @classmethod
    def from_obj(cls: Type[_TDataPointLabelStruct], obj: Any, **kwargs) -> _TDataPointLabelStruct:
        # sourcery skip: raise-from-previous-error
        """
        Gets instance from object

        Args:
            obj (object): UNO object

        Raises:
            PropertyNotFoundError: If ``obj`` does not have required property

        Returns:
            DataPointLabelStruct: ``DataPointLabelStruct`` instance.
        """
        # this nu is only used to get Property Name
        nu = cls(**kwargs)
        prop_name = nu._get_property_name()
        if not prop_name:
            raise ValueError("from_obj() Internal Property Name is empty")

        try:
            point = cast(DataPointLabel, mProps.Props.get(obj, prop_name))
        except mEx.PropertyNotFoundError:
            raise mEx.PropertyNotFoundError(prop_name, f"from_obj() obj as no {prop_name} property")

        return cls.from_uno_struct(point, **kwargs)

    # endregion from_obj()

    # endregion static methods

    # region Properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.STRUCT
        return self._format_kind_prop

    @property
    def prop_show_number(self) -> bool:
        """Gets or set if the number is additionally displayed in the caption."""
        return self._get(self._props.show_number)

    @prop_show_number.setter
    def prop_show_number(self, value: bool) -> None:
        self._set(self._props.show_number, value)

    @property
    def prop_show_number_in_percent(self) -> bool:
        """
        Only effective, if ``ShowNumber`` is ``True``.

        If this member is also ``True``, the numbers are displayed as percentages of a category
        """
        return self._get(self._props.show_number_in_percent)

    @prop_show_number_in_percent.setter
    def prop_show_number_in_percent(self, value: bool) -> None:
        self._set(self._props.show_number_in_percent, value)

    @property
    def prop_show_category_name(self) -> bool:
        """
        Gets or set if the caption contains the category name of the category to which a data point belongs
        """
        return self._get(self._props.show_category_name)

    @prop_show_category_name.setter
    def prop_show_category_name(self, value: bool) -> None:
        self._set(self._props.show_category_name, value)

    @property
    def prop_show_legend_symbol(self) -> bool:
        """
        Gets or set if the legend symbol is additionally displayed in the caption.
        """
        return self._get(self._props.show_legend_symbol)

    @prop_show_legend_symbol.setter
    def prop_show_legend_symbol(self, value: bool) -> None:
        self._set(self._props.show_legend_symbol, value)

    @property
    def prop_show_custom_label(self) -> bool:
        """
        Gets or set if a custom label is additionally displayed in the caption.

        Since LibreOffice ``7.1``
        """
        return self._get(self._props.show_custom_label)

    @prop_show_custom_label.setter
    def prop_show_custom_label(self, value: bool) -> None:
        self._set(self._props.show_custom_label, value)

    @property
    def prop_show_series_name(self) -> bool:
        """
        Gets or set if the caption contains the name of the series to which a data point belongs.

        Since LibreOffice ``7.2``
        """
        return self._get(self._props.show_series_name)

    @prop_show_series_name.setter
    def prop_show_series_name(self, value: bool) -> None:
        self._set(self._props.show_series_name, value)

    @property
    def _props(self) -> StructDataPointLabelProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = StructDataPointLabelProps(
                show_number="ShowNumber",
                show_number_in_percent="ShowNumberInPercent",
                show_category_name="ShowCategoryName",
                show_legend_symbol="ShowLegendSymbol",
                show_custom_label="ShowCustomLabel",
                show_series_name="ShowSeriesName",
            )
        return self._props_internal_attributes

    # endregion Properties
