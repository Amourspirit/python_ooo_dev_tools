from __future__ import annotations
from typing import cast, TYPE_CHECKING, overload, TypeVar, Type, Any, Tuple
import contextlib
import uno
from com.sun.star.text import XTextColumns

from ooodev.exceptions import ex as mEx
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.loader import lo as mLo
from ooodev.units.unit_mm import UnitMM
from ooodev.utils import props as mProps

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT
    from com.sun.star.text import TextColumns as UnoTextColumns

_TTextColumns = TypeVar("_TTextColumns", bound="TextColumns")


class TextColumns(StyleBase):
    """
    This class represents the text columns of a UNO object that supports ``com.sun.star.drawing.TextProperties`` service.
    """

    def __init__(self, col_count: int | None = None, spacing: float | UnitT | None = None) -> None:
        """
        Constructor.

        Args:
            count (int, optional): Number of columns. Defaults to None.
            spacing (float, UnitT, optional): Spacing between columns in MM units or ``UnitT``. Defaults to None.
        """
        super().__init__()
        self._spacing = None
        self._col_count = col_count
        self.prop_spacing = spacing

    def _create_columns(self) -> "UnoTextColumns | None":
        col_count = self.prop_col_count
        if col_count is None or col_count < 1:
            return None
        columns = cast(
            "UnoTextColumns", mLo.Lo.create_instance_msf(XTextColumns, "com.sun.star.text.TextColumns", raise_err=True)
        )
        if self._spacing is not None:
            columns.AutomaticDistance = self._spacing
        columns.setColumnCount(col_count)
        return columns

    def _get_unit_mm_100(self, value: float | UnitT | None) -> int | None:
        if value is None:
            return None
        with contextlib.suppress(AttributeError):
            return value.get_value_mm100()  # type: ignore
        return UnitMM(value).get_value_mm100()  # type: ignore

    # region Overridden Methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.drawing.TextProperties",)
        return self._supported_services_values

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies the properties to the given object.

        Args:
            obj (Any): UNO Shape object.

        Raises:
            mEx.NotSupportedError: Object is not supported for conversion to Line Properties
        """
        if not self._is_valid_obj(obj):
            raise mEx.NotSupportedError("Object is not supported for conversion to Line Properties")

        columns = self._create_columns()
        mProps.Props.set(obj, TextColumns=columns)

    def copy(self) -> TextColumns:
        """
        Creates a copy of this instance.

        Returns:
            TextColumns: Copy of this instance.
        """
        return TextColumns(self.prop_col_count, self.prop_spacing)

    # endregion Overridden Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TTextColumns], obj: object) -> _TTextColumns: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TTextColumns], obj: object, **kwargs) -> _TTextColumns: ...

    @classmethod
    def from_obj(cls: Type[_TTextColumns], obj: Any, **kwargs) -> _TTextColumns:
        """
        Creates a new instance from ``obj``.

        Args:
            obj (Any): UNO Shape object.

        Returns:
            Spacing: New instance.
        """
        inst = cls(**kwargs)

        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError("Object is not supported for conversion to Line Properties")

        text_cols = cast("UnoTextColumns", mProps.Props.get(obj, "TextColumns", None))
        if text_cols is not None:
            inst.prop_col_count = text_cols.getColumnCount()
            inst.prop_spacing = text_cols.AutomaticDistance
        return inst

    # endregion from_obj()

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.SHAPE
        return self._format_kind_prop

    @property
    def prop_col_count(self) -> int | None:
        return self._col_count

    @prop_col_count.setter
    def prop_col_count(self, value: int | None) -> None:
        self._col_count = value

    @property
    def prop_spacing(self) -> UnitMM | None:
        if self._spacing is None:
            return None
        return UnitMM.from_mm100(self._spacing)

    @prop_spacing.setter
    def prop_spacing(self, value: float | UnitT | None) -> None:
        self._spacing = self._get_unit_mm_100(value)

    # endregion Properties
