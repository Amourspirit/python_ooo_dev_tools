# region Import
from __future__ import annotations
from typing import cast, overload
from typing import Any, Tuple, Type, TypeVar
from enum import Enum
from ooo.dyn.table.cell_orientation import CellOrientation
from ooo.dyn.table.cell_vert_justify2 import CellVertJustify2

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.units.angle import Angle as Angle
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.format.inner.common.props.cell_text_orientation_props import CellTextOrientationProps

# endregion Import

_TTextOrientation = TypeVar(name="_TTextOrientation", bound="TextOrientation")


class EdgeKind(Enum):
    """Reference Edge Kind"""

    LOWER = CellVertJustify2.BOTTOM
    """Text Extension From Lower Cell Border (default)."""
    UPPER = CellVertJustify2.TOP
    """Text Extension From Upper Cell Border."""
    INSIDE = CellVertJustify2.STANDARD
    """Text Extension Inside Cell."""


class TextOrientation(StyleBase):
    """
    Text Rotation

    .. seealso::

        - :ref:`help_calc_format_direct_cell_alignment`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self, vert_stack: bool | None = None, rotation: int | Angle | None = None, edge: EdgeKind | None = None
    ) -> None:
        """
        Constructor

        Args:
            vert_stack (bool, optional): Specifies if vertical stack is to be used.
            rotation (int, Angle, optional): Specifies if the rotation.
            edge (EdgeKind, optional): Specifies the Reference Edge.

        Returns:
            None:

        Note:
            When ``vert_stack`` is ``True`` other parameters are not used.

        See Also:
            - :ref:`help_calc_format_direct_cell_alignment`
        """
        super().__init__()
        if vert_stack is not None:
            self.prop_vert_stacked = vert_stack
        if rotation is not None:
            self.prop_rotation = rotation
        if edge is not None:
            self.prop_edge = edge

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.CellStyle", "com.sun.star.table.CellProperties")
        return self._supported_services_values

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion Overrides

    # region Static Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TTextOrientation], obj: Any) -> _TTextOrientation: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TTextOrientation], obj: Any, **kwargs) -> _TTextOrientation: ...

    @classmethod
    def from_obj(cls: Type[_TTextOrientation], obj: Any, **kwargs) -> _TTextOrientation:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TextOrientation: Instance that represents text orientation options.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        for prop in inst._props:
            if prop:
                val = mProps.Props.get(obj, prop, None)
                if val is not None:
                    inst._set(prop, val)
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()

    # endregion Static Methods

    # region Properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.CELL
        return self._format_kind_prop

    @property
    def prop_vert_stacked(self) -> bool | None:
        """
        Gets/Sets vertically stacked.
        """
        pv = cast(CellOrientation, self._get(self._props.stacked))
        return None if pv is None else pv == CellOrientation.STACKED

    @prop_vert_stacked.setter
    def prop_vert_stacked(self, value: bool | None):
        if value is None:
            self._remove(self._props.stacked)
            return
        if value:
            self._set(self._props.stacked, CellOrientation.STACKED)
        else:
            self._set(self._props.stacked, CellOrientation.STANDARD)

    @property
    def prop_rotation(self) -> Angle | None:
        """Gets/Sets Vertical flip option"""
        # in 1/100 degree units
        pv = cast(int, self._get(self._props.rotation))
        return None if pv is None else Angle.from_angle100(pv)

    @prop_rotation.setter
    def prop_rotation(self, value: int | Angle | None) -> None:
        if value is None:
            self._remove(self._props.rotation)
            return
        val = Angle(int(value))
        # internally in 1/100 degree units
        self._set(self._props.rotation, val.get_angle100())

    @property
    def prop_edge(self) -> EdgeKind | None:
        """
        Gets/Sets Edge Kind.
        """
        pv = cast(int, self._get(self._props.rotate_ref))
        return None if pv is None else EdgeKind(pv)

    @prop_edge.setter
    def prop_edge(self, value: EdgeKind | None):
        if value is None:
            self._remove(self._props.rotate_ref)
            return
        self._set(self._props.rotate_ref, value.value)

    @property
    def _props(self) -> CellTextOrientationProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = CellTextOrientationProps(
                rotation="RotateAngle", rotate_ref="RotateReference", stacked="Orientation"
            )
        return self._props_internal_attributes

    # endregion Properties
