from __future__ import annotations
from typing import Any, Tuple, overload
import uno
from ooo.dyn.awt.size import Size as UnoSize

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.utils import lo as mLo
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_base import StyleBase
from ooodev.units import UnitObj, UnitConvert, UnitMM
from ooodev.format.inner.direct.structs.size_struct import SizeStruct


class Size(StyleBase):
    """
    Size of a shape.

    .. versionadded:: 0.9.4
    """

    def __init__(
        self,
        width: float | UnitObj,
        height: float | UnitObj,
    ) -> None:
        """
        Constructor

        Args:
            width (float | UnitObj): Specifies the width of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            height (float | UnitObj): Specifies the height of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
        """
        super().__init__()
        # self._chart_doc = chart_doc
        try:
            self._width = width.get_value_mm100()
        except AttributeError:
            self._width = UnitConvert.convert_mm_mm100(width)
        try:
            self._height = height.get_value_mm100()
        except AttributeError:
            self._height = UnitConvert.convert_mm_mm100(height)

    def _get_property_name(self) -> str:
        return "Size"

    # region Overridden Methods
    def apply(self, obj: object, **kwargs) -> None:
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
        struct = UnoSize(Width=self._width, Height=self._height)
        props = {name: struct}
        super().apply(obj=obj, override_dv=props)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.drawing.Shape",)
        return self._supported_services_values

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # region copy()
    @overload
    def copy(self) -> Size:
        ...

    @overload
    def copy(self, **kwargs) -> Size:
        ...

    def copy(self, **kwargs) -> Size:
        """
        Copy the current instance.

        Returns:
            Position: The copied instance.
        """
        cp = super().copy(**kwargs)
        cp._width = self._width
        cp._height = self._height
        return cp

    # endregion copy()
    # endregion Overridden Methods

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
    def prop_width(self) -> UnitMM:
        """Gets or sets the width of the shape (in ``mm`` units)."""
        return UnitMM.from_mm100(self._width)

    @prop_width.setter
    def prop_width(self, value: float | UnitObj) -> None:
        try:
            self._width = value.get_value_mm100()
        except AttributeError:
            self._width = UnitConvert.convert_mm_mm100(value)

    @property
    def prop_height(self) -> UnitMM:
        """Gets or sets the height of the shape (in ``mm`` units)."""
        return UnitMM.from_mm100(self._height)

    @prop_height.setter
    def prop_height(self, value: float | UnitObj) -> None:
        try:
            self._height = value.get_value_mm100()
        except AttributeError:
            self._height = UnitConvert.convert_mm_mm100(value)

    # endregion Properties
