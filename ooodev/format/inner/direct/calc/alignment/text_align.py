# region Imports
from __future__ import annotations
from typing import Any, Tuple, Type, TypeVar, cast, overload
from enum import Enum
from ooo.dyn.table.cell_vert_justify2 import CellVertJustify2
from ooo.dyn.table.cell_hori_justify import CellHoriJustify

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_pt import UnitPT
from ooodev.units.unit_obj import UnitT
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.format.inner.common.props.cell_text_align_props import CellTextAlignProps

# endregion Imports

_TTextAlign = TypeVar("_TTextAlign", bound="TextAlign")


# region Enums
class VertAlignKind(Enum):
    """Text Align Vertical Align Values."""

    DEFAULT = (0, CellVertJustify2.STANDARD)
    """Default alignment is used."""
    TOP = (0, CellVertJustify2.TOP)
    """Contents are aligned with the upper edge of the cell."""
    MIDDLE = (0, CellVertJustify2.CENTER)
    """Contents are aligned to the vertical middle of the cell."""
    BOTTOM = (0, CellVertJustify2.BOTTOM)
    """Contents are aligned to the lower edge of the cell."""
    JUSTIFIED = (0, CellVertJustify2.BLOCK)
    """Contents are justified to the cell height."""
    DISTRIBUTED = (1, CellVertJustify2.BLOCK)
    """Contents are distributed to the cell height."""

    def get_justify_method(self) -> int:
        """Gets the value that is typically applied to ``VertJustifyMethod`` property."""
        return self.value[0]

    def get_justify_value(self) -> int:
        """Gets the value that is typically applied to ``VertJustify`` property."""
        return self.value[1]


class HoriAlignKind(Enum):
    """Text Align Horizontal Align Values."""

    DEFAULT = (0, CellHoriJustify.STANDARD)
    """Default alignment is used."""
    LEFT = (0, CellHoriJustify.LEFT)
    """Contents are aligned to the left edge of the cell."""
    CENTER = (0, CellHoriJustify.CENTER)
    """Contents are horizontally centered."""
    RIGHT = (0, CellHoriJustify.RIGHT)
    """Contents are aligned to the right edge of the cell."""
    JUSTIFIED = (0, CellHoriJustify.BLOCK)
    """Contents are justified to the cell width."""
    FILLED = (0, CellHoriJustify.REPEAT)
    """Contents are repeated to fill the cell."""
    DISTRIBUTED = (1, CellHoriJustify.BLOCK)
    """Contents are justified to the cell width."""

    def get_justify_method(self) -> int:
        "Gets the value that is typically applied to ``HoriJustifyMethod`` property."
        return self.value[0]

    def get_justify_value(self) -> CellHoriJustify:
        "Gets the enum value that is typically applied to ``HoriJustify`` property."
        return self.value[1]


# endregion Enums


class TextAlign(StyleBase):
    """
    Cell Text Alignment.

    .. seealso::

        - :ref:`help_calc_format_direct_cell_alignment`

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(
        self,
        hori_align: HoriAlignKind | None = None,
        indent: float | UnitT | None = None,
        vert_align: VertAlignKind | None = None,
    ) -> None:
        """
        Constructor

        Args:
            hori_align (HoriAlignKind, optional): Specifies Horizontal Alignment.
            indent: (float, UnitT, optional): Specifies indent in ``pt`` (point) units or :ref:`proto_unit_obj`.
                Only used when ``hori_align`` is set to ``HoriAlignKind.LEFT``
            vert_align (VertAdjustKind, optional): Specifies Vertical Alignment.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_direct_cell_alignment`
        """
        super().__init__()
        if hori_align is not None:
            self.prop_hori_align = hori_align
        if indent is not None:
            self.prop_indent = indent
        if vert_align is not None:
            self.prop_vert_align = vert_align

    # endregion Init

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
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TTextAlign], obj: Any) -> _TTextAlign: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TTextAlign], obj: Any, **kwargs) -> _TTextAlign: ...

    @classmethod
    def from_obj(cls: Type[_TTextAlign], obj: Any, **kwargs) -> _TTextAlign:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TextAlign: Instance that represents Text Alignment.
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

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.CELL
        return self._format_kind_prop

    @property
    def prop_hori_align(self) -> HoriAlignKind | None:
        """Gets/Sets Horizontal align value."""
        justify = cast(int, self._get(self._props.hori_justify))
        meth = cast(int, self._get(self._props.hori_method))
        if justify is None:
            return None
        return None if meth is None else HoriAlignKind((meth, justify))

    @prop_hori_align.setter
    def prop_hori_align(self, value: HoriAlignKind | None) -> None:
        if value is None:
            self._remove(self._props.hori_justify)
            self._remove(self._props.hori_method)
            return
        self._set(self._props.hori_justify, value.get_justify_value())
        self._set(self._props.hori_method, value.get_justify_method())

    @property
    def prop_indent(self) -> UnitPT | None:
        """
        Gets/Sets indent.
        """
        pv = cast(int, self._get(self._props.indent))
        return None if pv is None else UnitPT.from_mm100(pv)

    @prop_indent.setter
    def prop_indent(self, value: float | UnitT | None):
        if value is None:
            self._remove(self._props.indent)
            return
        try:
            val = value.get_value_mm100()  # type: ignore
        except AttributeError:
            # value is rounded in Calc Dialog.
            val = UnitConvert.convert_pt_mm100(round(value))  # type: ignore
        val = max(val, 0)
        self._set(self._props.indent, val)

    @property
    def prop_vert_align(self) -> VertAlignKind | None:
        """Gets/Sets vertical align value."""
        justify = cast(int, self._get(self._props.vert_justify))
        meth = cast(int, self._get(self._props.vert_method))
        if justify is None:
            return None
        return None if meth is None else VertAlignKind((meth, justify))

    @prop_vert_align.setter
    def prop_vert_align(self, value: VertAlignKind | None) -> None:
        if value is None:
            self._remove(self._props.vert_justify)
            self._remove(self._props.vert_method)
            return
        self._set(self._props.vert_justify, value.get_justify_value())
        self._set(self._props.vert_method, value.get_justify_method())

    @property
    def _props(self) -> CellTextAlignProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = CellTextAlignProps(
                hori_justify="HoriJustify",
                hori_method="HoriJustifyMethod",
                vert_justify="VertJustify",
                vert_method="VertJustifyMethod",
                indent="ParaIndent",
            )
        return self._props_internal_attributes
