"""
Module for managing character borders.

.. versionadded:: 0.9.0
"""

# region imports
from __future__ import annotations
from typing import Any, Tuple, overload, cast, Type, TypeVar

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.direct.structs.side import Side
from ooodev.format.inner.direct.write.para.border.padding import Padding as InnerPadding
from ooodev.format.inner.direct.write.para.border.shadow import Shadow as InnerShadow
from ooodev.format.inner.direct.write.para.border.sides import Sides
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps

# endregion imports

_TBorders = TypeVar("_TBorders", bound="Borders")


class Borders(StyleMulti):
    """
    Border used in styles for characters.

    All methods starting with ``fmt_`` can be used to chain together Borders properties.

    .. seealso::

        - :ref:`help_writer_format_direct_para_borders`

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
        all: Side | None = None,
        shadow: InnerShadow | None = None,
        padding: InnerPadding | None = None,
        merge: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (Side, None, optional): Determines the line style at the left edge.
            right (Side, None, optional): Determines the line style at the right edge.
            top (Side, None, optional): Determines the line style at the top edge.
            bottom (Side, None, optional): Determines the line style at the bottom edge.
            all (Side, None, optional): Determines the line style at the top, bottom, left, right edges.
                If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
            shadow (BorderShadow, None, optional): Character Shadow
            padding (BorderPadding, None, optional): Character padding
            merge (bool, None, optional): Merge with next paragraph

        Returns:
            None:

        See Also:

            - :ref:`help_writer_format_direct_para_borders`
        """
        init_vals = {}

        sides = self._get_sides_type()(
            left=left,
            right=right,
            top=top,
            bottom=bottom,
            all=all,
        )

        if merge is not None:
            init_vals["ParaIsConnectBorder"] = merge

        if padding is not None:
            # BorderDistance is set to padding bottom for some reason.
            init_vals["BorderDistance"] = padding._get(padding._props.bottom)

        super().__init__(**init_vals)

        if sides.prop_has_attribs:
            sides._prop_parent = self
            self._set_style("sides", sides, *sides.get_attrs())  # type: ignore
        if padding is not None:
            self._set_style("padding", padding, *padding.get_attrs())
        if shadow is not None:
            self._set_style("shadow", shadow, *shadow.get_attrs())

    # endregion init

    # region format Methods
    def fmt_border_side(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with left, right, top, bottom sides set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        if self.prop_inner_sides is not None:
            sides = self.prop_inner_sides.fmt_border_side(value)
        else:
            sides = self._get_sides_type()(all=value)
        cp._set_style("sides", sides, *sides.get_attrs())  # type: ignore
        return cp

    def fmt_left(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with left set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        if self.prop_inner_sides is not None:
            sides = self.prop_inner_sides.fmt_left(value)
        else:
            sides = self._get_sides_type()(left=value)
        cp._set_style("sides", sides, *sides.get_attrs())  # type: ignore
        return cp

    def fmt_right(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with right set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        if self.prop_inner_sides is not None:
            sides = self.prop_inner_sides.fmt_right(value)
        else:
            sides = self._get_sides_type()(right=value)
        cp._set_style("sides", sides, *sides.get_attrs())  # type: ignore
        return cp

    def fmt_top(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with top set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        if self.prop_inner_sides is not None:
            sides = self.prop_inner_sides.fmt_top(value)
        else:
            sides = self._get_sides_type()(top=value)
        cp._set_style("sides", sides, *sides.get_attrs())  # type: ignore
        return cp

    def fmt_bottom(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with bottom set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        if self.prop_inner_sides is not None:
            sides = self.prop_inner_sides.fmt_bottom(value)
        else:
            sides = self._get_sides_type()(bottom=value)
        cp._set_style("sides", sides, *sides.get_attrs())  # type: ignore
        return cp

    def fmt_shadow(self: _TBorders, value: InnerShadow | None) -> _TBorders:
        """
        Gets copy of instance with shadow set or removed

        Args:
            value (Shadow | None): Shadow value

        Returns:
            Borders: Borders instance
        """
        # pylint: disable=protected-access
        cp = self.copy()
        if value is None:
            cp._remove_style("shadow")
        else:
            cp._set_style("shadow", value, *value.get_attrs())
        return cp

    def fmt_padding(self: _TBorders, value: InnerPadding | None) -> _TBorders:
        """
        Gets copy of instance with padding set or removed

        Args:
            value (BorderPadding | None): Padding value

        Returns:
            Borders: Borders instance
        """
        # pylint: disable=protected-access
        cp = self.copy()
        if value is None:
            cp._remove_style("padding")
        else:
            cp._set_style("padding", value, *value.get_attrs())
        return cp

    # endregion format Methods

    # region methods

    def _get_sides_type(self) -> Type[Sides]:
        return Sides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.CharacterProperties",
                "com.sun.star.style.ParagraphStyle",
                "com.sun.star.text.TextFrame",
            )
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    # region apply()
    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): Object that supports ``com.sun.star.style.CharacterProperties`` service.

        Returns:
            None:
        """

        try:
            super().apply(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TBorders], obj: Any) -> _TBorders: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TBorders], obj: Any, **kwargs) -> _TBorders: ...

    @classmethod
    def from_obj(cls: Type[_TBorders], obj: Any, **kwargs) -> _TBorders:
        """
        Gets instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Raises:
            ServiceNotSupported: If ``obj`` does not support  ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            Borders: ``Borders`` instance that represents the ``obj`` borders.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.ServiceNotSupported(inst._supported_services()[0])
        inst_sides = Sides.from_obj(obj)
        inst_padding = InnerPadding.from_obj(obj)
        inst_shadow = InnerShadow.from_obj(obj)
        inst._set("ParaIsConnectBorder", mProps.Props.get(obj, "ParaIsConnectBorder"))
        inst._set("BorderDistance", mProps.Props.get(obj, "BorderDistance"))
        inst._set_style("sides", inst_sides, *inst_sides.get_attrs())
        inst._set_style("padding", inst_padding, *inst_padding.get_attrs())
        inst._set_style("shadow", inst_shadow, *inst_shadow.get_attrs())
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()

    # endregion methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA | FormatKind.STATIC
        return self._format_kind_prop

    @property
    def prop_inner_sides(self) -> Sides | None:
        """Gets Sides instance"""
        try:
            return self._direct_inner_sides
        except AttributeError:
            self._direct_inner_sides = cast(Sides, self._get_style_inst("sides"))
        return self._direct_inner_sides

    @property
    def prop_inner_padding(self) -> InnerPadding | None:
        """Gets Padding instance"""
        try:
            return self._direct_inner_padding
        except AttributeError:
            self._direct_inner_padding = cast(InnerPadding, self._get_style_inst("padding"))
        return self._direct_inner_padding

    @property
    def prop_inner_shadow(self) -> InnerShadow | None:
        """Gets Shadow instance"""
        try:
            return self._direct_inner_shadow
        except AttributeError:
            self._direct_inner_shadow = cast(InnerShadow, self._get_style_inst("shadow"))
        return self._direct_inner_shadow

    @property
    def default(self: _TBorders) -> _TBorders:  # type: ignore[misc]
        """Gets Default Border."""
        # pylint: disable=protected-access
        try:
            return self._default_inst
        except AttributeError:
            if self.prop_inner_padding is None:
                padding = InnerPadding().default
            else:
                padding = self.prop_inner_padding.default
            if self.prop_inner_shadow is None:
                shadow = InnerShadow().empty
            else:
                shadow = self.prop_inner_shadow.empty
            self._default_inst = self.__class__(all=Side().empty, padding=padding, shadow=shadow, merge=True)
            self._default_inst._is_default_inst = True
        return self._default_inst

    # endregion Properties
