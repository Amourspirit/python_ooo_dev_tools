# region Imports
from __future__ import annotations
from typing import Any, Tuple, cast, TypeVar, Type, overload

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.format.calc.style.cell.kind.style_cell_kind import StyleCellKind

from ooodev.format.inner.direct.calc.border.padding import Padding
from ooodev.format.inner.direct.calc.border.shadow import Shadow
from ooodev.format.inner.common.props.border_props import BorderProps
from ooodev.format.inner.common.props.cell_style_borders_props import CellStyleBordersProps
from ooodev.format.inner.direct.structs.side import Side
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.format.inner.modify.calc.cell_style_base_multi import CellStyleBaseMulti

# endregion Imports

_TInnerBorders = TypeVar(name="_TInnerBorders", bound="InnerBorders")


class InnerBorders(StyleMulti):
    """
    Style Table Borders used in styles for table cells and ranges.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        *,
        right: Side | None = None,
        left: Side | None = None,
        top: Side | None = None,
        bottom: Side | None = None,
        border_side: Side | None = None,
        diagonal_down: Side | None = None,
        diagonal_up: Side | None = None,
        shadow: Shadow | None = None,
        padding: Padding | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the left edge.
            right (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the right edge.
            top (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the top edge.
            bottom (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the bottom edge.
            border_side (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the top, bottom, left, right edges.
                If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
            diagonal_down (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style from top-left to bottom-right diagonal.
            diagonal_up (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style from bottom-left to top-right diagonal.
            shadow (Shadow, optional): Cell Shadow.
            padding (BorderPadding, optional): Cell padding.
        """
        init_vals = {}

        if shadow is None:
            shadow_fmt = None
        else:
            shadow_fmt = shadow.copy(_cattribs=self._get_shadow_cattribs())

        if diagonal_down is None:
            diag_dn = None
        else:
            diag_dn = diagonal_down.copy(_cattribs=self._get_diagonal_dn_cattribs())

        if diagonal_up is None:
            diag_up = None
        else:
            diag_up = diagonal_up.copy(_cattribs=self._get_diagonal_up_cattribs())

        if padding is None:
            padding_fmt = None
        else:
            padding_fmt = padding.copy(_cattribs=self._get_padding_cattribs())

        bdr_right = None
        bdr_left = None
        bdr_top = None
        bdr_btm = None

        if border_side is None:
            if not right is None:
                bdr_right = right.copy(_cattribs=self._get_bdr_right_cattribs())

            if not left is None:
                bdr_left = left.copy(_cattribs=self._get_bdr_left_cattribs())

            if not top is None:
                bdr_top = top.copy(_cattribs=self._get_bdr_top_cattribs())

            if not bottom is None:
                bdr_btm = bottom.copy(_cattribs=self._get_bdr_btm_cattribs())
        else:
            bdr_right = border_side.copy(_cattribs=self._get_bdr_right_cattribs())
            bdr_left = border_side.copy(_cattribs=self._get_bdr_left_cattribs())
            bdr_top = border_side.copy(_cattribs=self._get_bdr_top_cattribs())
            bdr_btm = border_side.copy(_cattribs=self._get_bdr_btm_cattribs())

        super().__init__(**init_vals)
        if not padding_fmt is None:
            self._set_style("padding", padding_fmt)
        if not shadow_fmt is None:
            self._set_style("shadow", shadow_fmt)
        if not diag_dn is None:
            self._set_style("diag_dn", diag_dn)
        if not diag_up is None:
            self._set_style("diag_up", diag_up)
        if not bdr_right is None:
            self._set_style("bdr_right", bdr_right)
        if not bdr_left is None:
            self._set_style("bdr_left", bdr_left)
        if not bdr_top is None:
            self._set_style("bdr_top", bdr_top)
        if not bdr_btm is None:
            self._set_style("bdr_btm", bdr_btm)

    # endregion init

    # region internal methods
    def _get_shadow_cattribs(self) -> dict:
        return {
            "_property_name": self._props.shadow,
            "_supported_services_values": self._supported_services(),
        }

    def _get_padding_cattribs(self) -> dict:
        return {
            "_props_internal_attributes": BorderProps(
                left=self._props.pad_left,
                top=self._props.pad_top,
                right=self._props.pad_right,
                bottom=self._props.pad_btm,
            ),
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
        }

    def _get_diagonal_up_cattribs(self) -> dict:
        return {
            "_property_name": self._props.diag_up,
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
        }

    def _get_diagonal_dn_cattribs(self) -> dict:
        return {
            "_property_name": self._props.diag_dn,
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
        }

    def _get_bdr_left_cattribs(self) -> dict:
        return {
            "_property_name": self._props.bdr_left,
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
        }

    def _get_bdr_right_cattribs(self) -> dict:
        return {
            "_property_name": self._props.bdr_right,
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
        }

    def _get_bdr_top_cattribs(self) -> dict:
        return {
            "_property_name": self._props.bdr_top,
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
        }

    def _get_bdr_btm_cattribs(self) -> dict:
        return {
            "_property_name": self._props.bdr_btm,
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
        }

    # endregion internal methods

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.CellStyle",)
        return self._supported_services_values

    # region apply()
    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO object.

        Returns:
            None:
        """
        super().apply(obj, **kwargs)

    # endregion apply()

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion Overrides

    # region Static Methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TInnerBorders], obj: Any) -> _TInnerBorders: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TInnerBorders], obj: Any, **kwargs) -> _TInnerBorders: ...

    @classmethod
    def from_obj(cls: Type[_TInnerBorders], obj: Any, **kwargs) -> _TInnerBorders:
        """
        Gets Borders instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedServiceError: If ``obj`` is not supported.

        Returns:
            Borders: Borders that represents ``obj`` borders.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        shadow_fmt = Shadow.from_obj(obj=obj, _cattribs=inst._get_shadow_cattribs())
        inst._set_style("shadow", shadow_fmt)

        padding_fmt = Padding.from_obj(obj=obj, _cattribs=inst._get_padding_cattribs())
        inst._set_style("padding", padding_fmt)

        diag_dn = Side.from_obj(obj=obj, _cattribs=inst._get_diagonal_dn_cattribs())
        inst._set_style("diag_dn", diag_dn)

        diag_up = Side.from_obj(obj=obj, _cattribs=inst._get_diagonal_up_cattribs())
        inst._set_style("diag_up", diag_up)

        bdr_right = Side.from_obj(obj=obj, _cattribs=inst._get_bdr_right_cattribs())
        inst._set_style("bdr_right", bdr_right)

        bdr_left = Side.from_obj(obj=obj, _cattribs=inst._get_bdr_left_cattribs())
        inst._set_style("bdr_left", bdr_left)

        bdr_top = Side.from_obj(obj=obj, _cattribs=inst._get_bdr_top_cattribs())
        inst._set_style("bdr_top", bdr_top)

        bdr_btm = Side.from_obj(obj=obj, _cattribs=inst._get_bdr_btm_cattribs())
        inst._set_style("bdr_btm", bdr_btm)

        return inst

    # endregion from_obj()
    # endregion Static Methods

    # region Style Methods
    def fmt_left(self: _TInnerBorders, value: Side | None) -> _TInnerBorders:
        """
        Gets copy of instance with left set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        cp.prop_left = value
        return cp

    def fmt_right(self: _TInnerBorders, value: Side | None) -> _TInnerBorders:
        """
        Gets copy of instance with right set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        cp.prop_right = value
        return cp

    def fmt_top(self: _TInnerBorders, value: Side | None) -> _TInnerBorders:
        """
        Gets copy of instance with top set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        cp.prop_top = value
        return cp

    def fmt_bottom(self: _TInnerBorders, value: Side | None) -> _TInnerBorders:
        """
        Gets copy of instance with bottom set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        cp.prop_bottom = value
        return cp

    def fmt_diagonal_down(self: _TInnerBorders, value: Side | None) -> _TInnerBorders:
        """
        Gets copy of instance with diagonal down set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        cp.prop_diagonal_dn = value
        return cp

    def fmt_diagonal_up(self: _TInnerBorders, value: Side | None) -> _TInnerBorders:
        """
        Gets copy of instance with diagonal up set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        cp.prop_diagonal_up = value
        return cp

    def fmt_shadow(self: _TInnerBorders, value: Shadow | None) -> _TInnerBorders:
        """
        Gets copy of instance with shadow set or removed

        Args:
            value (Shadow, optional): Shadow value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        cp.prop_shadow = value
        return cp

    def fmt_padding(self: _TInnerBorders, value: Padding | None) -> _TInnerBorders:
        """
        Gets copy of instance with padding set or removed

        Args:
            value (Padding, optional): Padding value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        cp.prop_padding = value
        return cp

    # endregion Style Methods

    # region Properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.STYLE
        return self._format_kind_prop

    @property
    def prop_padding(self) -> Padding | None:
        """Gets/Sets Padding instance"""
        try:
            return self._direct_inner_padding
        except AttributeError:
            self._direct_inner_padding = cast(Padding, self._get_style_inst("padding"))
        return self._direct_inner_padding

    @prop_padding.setter
    def prop_padding(self, value: Padding | None) -> None:
        self._del_attribs("_direct_inner_padding")
        if value is None:
            self._remove_style("padding")
            return
        self._set_style("padding", value.copy(_cattribs=self._get_padding_cattribs()))

    @property
    def prop_shadow(self) -> Shadow | None:
        """Gets/Sets inner shadow instance"""
        try:
            return self._direct_inner_shadow
        except AttributeError:
            self._direct_inner_shadow = cast(Shadow, self._get_style_inst("shadow"))
        return self._direct_inner_shadow

    @prop_shadow.setter
    def prop_shadow(self, value: Shadow | None) -> None:
        self._del_attribs("_direct_inner_shadow")
        if value is None:
            self._remove_style("shadow")
            return
        self._set_style("shadow", value.copy(_cattribs=self._get_shadow_cattribs()))

    @property
    def prop_diagonal_up(self) -> Side | None:
        """Gets/Sets inner Diagonal up instance"""
        try:
            return self._direct_dial_up
        except AttributeError:
            self._direct_dial_up = cast(Side, self._get_style_inst("diag_up"))
        return self._direct_dial_up

    @prop_diagonal_up.setter
    def prop_diagonal_up(self, value: Side | None) -> None:
        self._del_attribs("_direct_dial_up")
        if value is None:
            self._remove_style("diag_up")
            return
        self._set_style("diag_up", value.copy(_cattribs=self._get_diagonal_up_cattribs()))

    @property
    def prop_diagonal_dn(self) -> Side | None:
        """Gets/Sets inner Diagonal down instance"""
        try:
            return self._direct_dial_dn
        except AttributeError:
            self._direct_dial_dn = cast(Side, self._get_style_inst("diag_dn"))
        return self._direct_dial_dn

    @prop_diagonal_dn.setter
    def prop_diagonal_dn(self, value: Side | None) -> None:
        self._del_attribs("_direct_dial_dn")
        if value is None:
            self._remove_style("diag_dn")
            return
        self._set_style("diag_dn", value.copy(_cattribs=self._get_diagonal_dn_cattribs()))

    @property
    def prop_left(self) -> Side | None:
        """Gets/Sets inner Left instance"""
        try:
            return self._direct_left
        except AttributeError:
            self._direct_left = cast(Side, self._get_style_inst("bdr_left"))
        return self._direct_left

    @prop_left.setter
    def prop_left(self, value: Side | None) -> None:
        self._del_attribs("_direct_left")
        if value is None:
            self._remove_style("bdr_left")
            return
        self._set_style("bdr_left", value.copy(_cattribs=self._get_bdr_left_cattribs()))

    @property
    def prop_right(self) -> Side | None:
        """Gets/Sets inner Right instance"""
        try:
            return self._direct_right
        except AttributeError:
            self._direct_right = cast(Side, self._get_style_inst("bdr_right"))
        return self._direct_right

    @prop_right.setter
    def prop_right(self, value: Side | None) -> None:
        self._del_attribs("_direct_right")
        if value is None:
            self._remove_style("bdr_right")
            return
        self._set_style("bdr_right", value.copy(_cattribs=self._get_bdr_right_cattribs()))

    @property
    def prop_top(self) -> Side | None:
        """Gets/Sets inner Top instance"""
        try:
            return self._direct_top
        except AttributeError:
            self._direct_top = cast(Side, self._get_style_inst("bdr_top"))
        return self._direct_top

    @prop_top.setter
    def prop_top(self, value: Side | None) -> None:
        self._del_attribs("_direct_top")
        if value is None:
            self._remove_style("bdr_top")
            return
        self._set_style("bdr_top", value.copy(_cattribs=self._get_bdr_top_cattribs()))

    @property
    def prop_bottom(self) -> Side | None:
        """Gets/Sets inner Bottom instance"""
        try:
            return self._direct_bottom
        except AttributeError:
            self._direct_bottom = cast(Side, self._get_style_inst("bdr_btm"))
        return self._direct_bottom

    @prop_bottom.setter
    def prop_bottom(self, value: Side | None) -> None:
        self._del_attribs("_direct_bottom")
        if value is None:
            self._remove_style("bdr_btm")
            return
        self._set_style("bdr_btm", value.copy(_cattribs=self._get_bdr_btm_cattribs()))

    @property
    def _props(self) -> CellStyleBordersProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = CellStyleBordersProps(
                shadow="ShadowFormat",
                diag_up="DiagonalBLTR2",
                diag_dn="DiagonalTLBR2",
                bdr_left="LeftBorder2",
                bdr_right="RightBorder2",
                bdr_top="TopBorder2",
                bdr_btm="BottomBorder2",
                pad_left="ParaLeftMargin",
                pad_top="ParaTopMargin",
                pad_right="ParaRightMargin",
                pad_btm="ParaBottomMargin",
            )
        return self._props_internal_attributes

    # endregion Properties


class Borders(CellStyleBaseMulti):
    """
    Cell Style Borders.

    .. seealso::

        - :ref:`help_calc_format_modify_cell_borders`

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(
        self,
        *,
        right: Side | None = None,
        left: Side | None = None,
        top: Side | None = None,
        bottom: Side | None = None,
        border_side: Side | None = None,
        diagonal_down: Side | None = None,
        diagonal_up: Side | None = None,
        shadow: Shadow | None = None,
        padding: Padding | None = None,
        style_name: StyleCellKind | str = StyleCellKind.DEFAULT,
        style_family: str = "CellStyles",
    ) -> None:
        """
        Constructor

        Args:
            left (~ooodev.format.inner.direct.structs.side.Side,, optional): Specifies the line style at the left edge.
            right (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the right edge.
            top (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the top edge.
            bottom (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the bottom edge.
            border_side (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the top, bottom, left, right edges.
                If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
            diagonal_down (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style from top-left to bottom-right diagonal.
            diagonal_up (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style from bottom-left to top-right diagonal.
            shadow (Shadow, optional): Cell Shadow.
            padding (BorderPadding, optional): Cell padding.
            style_name (StyleCellKind, str, optional): Specifies the Cell Style that instance applies to.
                Default is Default Cell Style.
            style_family (str, optional): Style family. Default ``CellStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_cell_borders`
        """

        direct = InnerBorders(
            right=right,
            left=left,
            top=top,
            bottom=bottom,
            border_side=border_side,
            diagonal_down=diagonal_down,
            diagonal_up=diagonal_up,
            shadow=shadow,
            padding=padding,
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct)

    # endregion Init

    # region Static Methods
    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleCellKind | str = StyleCellKind.DEFAULT,
        style_family: str = "CellStyles",
    ) -> Borders:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleCellKind, str, optional): Specifies the Cell Style that instance applies to.
                Default is Default Cell Style.
            style_family (str, optional): Style family. Default ``CellStyles``.

        Returns:
            Borders: ``Borders`` instance from style properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerBorders.from_obj(obj=inst.get_style_props(doc))
        inst._set_style("direct", direct)
        return inst

    # endregion Static Methods

    # region Properties
    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleCellKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerBorders:
        """Gets/Sets Inner Borders instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerBorders, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerBorders) -> None:
        if not isinstance(value, InnerBorders):
            raise TypeError(f'Expected type of InnerBorders, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value)

    # endregion Properties
