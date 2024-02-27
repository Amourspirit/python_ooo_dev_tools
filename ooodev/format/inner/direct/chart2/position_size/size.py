from __future__ import annotations
from typing import Any, cast, Tuple, overload, TYPE_CHECKING
import uno  # pylint: disable=unused-import
from ooo.dyn.awt.size import Size as UnoSize

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.loader import lo as mLo
from ooodev.exceptions import ex as mEx
from ooodev.utils import props as mProps
from ooodev.format.inner.style_base import StyleBase
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_mm import UnitMM
from ooodev.units.unit_mm100 import UnitMM100

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT
    from typing_extensions import Self
else:
    Self = Any


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
    def _container_get_service_name(self) -> str:
        # keep type checker happy.
        raise NotImplementedError

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
            Size: The copied instance.
        """
        # pylint: disable=protected-access
        cp = super().copy(width=0, height=0, **kwargs)
        cp._width = self._width
        cp._height = self._height
        return cp

    # endregion copy()
    # endregion Overridden Methods

    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> Self:
        """
        Creates a new instance from ``obj``.

        Args:
            obj (Any): UNO Shape object.

        Returns:
            Size: New instance.
        """
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> Self:
        """
        Creates a new instance from ``obj``.

        Args:
            obj (Any): UNO Shape object.
            **kwargs: Additional arguments.

        Returns:
            Size: New instance.
        """
        ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> Self:
        """
        Creates a new instance from ``obj``.

        Args:
            obj (Any): UNO Shape object.

        Returns:
            Size: New instance.
        """
        # pylint: disable=protected-access
        inst = cls(width=0, height=0, **kwargs)
        name = inst._get_property_name()
        if not name:
            raise ValueError("No property name to retrieve.")

        sz = cast(UnoSize, mProps.Props.get(obj, name))
        nu = cls(width=UnitMM100(sz.Width), height=UnitMM100(sz.Height), **kwargs)
        nu.set_update_obj(obj)
        return nu

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
