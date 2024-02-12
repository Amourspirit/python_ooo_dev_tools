from __future__ import annotations
from typing import Any, Tuple, overload, TYPE_CHECKING
import uno
from ooo.dyn.awt.size import Size as UnoSize

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.loader import lo as mLo
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_base import StyleBase
from ooodev.units import UnitConvert, UnitMM

if TYPE_CHECKING:
    from ooodev.units import UnitT


class Size(StyleBase):
    """
    Size of a shape.

    .. versionadded:: 0.9.4
    """

    def __init__(
        self,
        width: float | UnitT,
        height: float | UnitT,
    ) -> None:
        """
        Constructor

        Args:
            width (float | UnitT): Specifies the width of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            height (float | UnitT): Specifies the height of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.

        Returns:
            None:
        """
        super().__init__()
        # self._chart_doc = chart_doc
        try:
            self._width = width.get_value_mm100()  # type: ignore
        except AttributeError:
            self._width = UnitConvert.convert_mm_mm100(width)  # type: ignore
        try:
            self._height = height.get_value_mm100()  # type: ignore
        except AttributeError:
            self._height = UnitConvert.convert_mm_mm100(height)  # type: ignore

    def _get_property_name(self) -> str:
        return "Size"

    # region Overridden Methods
    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies properties to ``obj``

        Args:
            obj (Any): UNO object.

        Returns:
            None:
        """
        name = self._get_property_name()
        if not name:
            return
        props = kwargs.pop("override_dv", {})
        update_dv = bool(kwargs.pop("update_dv", True))
        if update_dv:
            struct = UnoSize(Width=self._width, Height=self._height)
            props.update({name: struct})
        if props:
            super().apply(obj=obj, override_dv=props)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.drawing.Shape",)
        return self._supported_services_values

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # region copy()
    @overload
    def copy(self) -> Size: ...

    @overload
    def copy(self, **kwargs) -> Size: ...

    def copy(self, **kwargs) -> Size:
        """
        Copy the current instance.

        Returns:
            Position: The copied instance.
        """
        cp = super().copy(width=0, height=0, **kwargs)
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
    def prop_width(self, value: float | UnitT) -> None:
        try:
            self._width = value.get_value_mm100()  # type: ignore
        except AttributeError:
            self._width = UnitConvert.convert_mm_mm100(value)  # type: ignore

    @property
    def prop_height(self) -> UnitMM:
        """Gets or sets the height of the shape (in ``mm`` units)."""
        return UnitMM.from_mm100(self._height)

    @prop_height.setter
    def prop_height(self, value: float | UnitT) -> None:
        try:
            self._height = value.get_value_mm100()  # type: ignore
        except AttributeError:
            self._height = UnitConvert.convert_mm_mm100(value)  # type: ignore

    # endregion Properties
