"""
Module for managing character borders.

.. versionadded:: 0.9.0
"""

# region Import
from __future__ import annotations
from typing import Any, Tuple, cast, overload, TypeVar

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.format.inner.kind.border_kind import BorderKind
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.format.inner.direct.structs.side import Side
from ooodev.format.inner.direct.write.char.border.padding import Padding as InnerPadding
from ooodev.format.inner.direct.write.char.border.shadow import Shadow as InnerShadow
from ooodev.format.inner.direct.write.char.border.sides import Sides

# endregion Import

_TBorders = TypeVar("_TBorders", bound="Borders")


class Borders(StyleMulti):
    """
    Border used in styles for characters.

    All methods starting with ``fmt_`` can be used to chain together Borders properties.

    .. seealso::

        - :ref:`help_writer_format_direct_char_borders`

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
    ) -> None:
        """
        Constructor

        Args:
            left (Side | None, optional): Determines the line style at the left edge.
            right (Side | None, optional): Determines the line style at the right edge.
            top (Side | None, optional): Determines the line style at the top edge.
            bottom (Side | None, optional): Determines the line style at the bottom edge.
            all (Side | None, optional): Determines the line style at the top, bottom, left, right edges. If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
            shadow (Shadow | None, optional): Character Shadow
            padding (Padding | None, optional): Character padding

        Returns:
            None:

        See Also:

            - :ref:`help_writer_format_direct_char_borders`
        """
        init_vals = {}

        sides = Sides(
            left=left,
            right=right,
            top=top,
            bottom=bottom,
            all=all,
        )

        super().__init__(**init_vals)

        if sides.prop_has_attribs:
            self._set_style("sides", sides, *sides.get_attrs())  # type: ignore
        if padding is not None:
            self._set_style("padding", padding, *padding.get_attrs())  # type: ignore
        if shadow is not None:
            self._set_style("shadow", shadow, *shadow.get_attrs())

    # endregion init
    def _set_side(self, side: Side | None, pos: BorderKind, inst: Borders | None = None) -> None:
        if inst is None:
            inst = self
        if pos == BorderKind.NONE:
            inst._remove_style("sides")
            return
        sides_info = inst._get_style_inst("sides")

        sides = Sides() if sides_info is None else cast(Sides, sides_info.copy())
        if BorderKind.BOTTOM in pos:
            sides.prop_bottom = side
        if BorderKind.LEFT in pos:
            sides.prop_left = side
        if BorderKind.RIGHT in pos:
            sides.prop_right = side
        if BorderKind.TOP in pos:
            sides.prop_bottom = side
        inst._set_style("sides", sides, *sides.get_attrs())

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
        if value is None:
            self._set_side(side=value, pos=BorderKind.NONE, inst=cp)
            return cp
        self._set_side(side=value, pos=BorderKind.ALL, inst=cp)
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
        if value is None:
            self._set_side(side=value, pos=BorderKind.NONE, inst=cp)
            return cp
        self._set_side(side=value, pos=BorderKind.LEFT, inst=cp)
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
        if value is None:
            self._set_side(side=value, pos=BorderKind.NONE, inst=cp)
            return cp
        self._set_side(side=value, pos=BorderKind.RIGHT, inst=cp)
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
        if value is None:
            self._set_side(side=value, pos=BorderKind.NONE, inst=cp)
            return cp
        self._set_side(side=value, pos=BorderKind.TOP, inst=cp)
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
        if value is None:
            self._set_side(side=value, pos=BorderKind.NONE, inst=cp)
            return cp
        self._set_side(side=value, pos=BorderKind.BOTTOM, inst=cp)
        return cp

    def fmt_shadow(self: _TBorders, value: InnerShadow | None) -> _TBorders:
        """
        Gets copy of instance with shadow set or removed

        Args:
            value (Shadow | None): Shadow value

        Returns:
            Borders: Borders instance
        """
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
            value (Padding | None): Padding value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        if value is None:
            cp._remove_style("padding")
        else:
            cp._set_style("padding", value, *value.get_attrs())
        return cp

    # endregion format Methods

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.CharacterProperties",
                "com.sun.star.style.CharacterStyle",
            )
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    # region apply()
    @overload
    def apply(self, obj: Any) -> None:  # type: ignore
        ...

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
    # endregion methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.CHAR
        return self._format_kind_prop

    @property
    def prop_inner_sides(self) -> Sides | None:
        """Gets Sides Instance"""
        try:
            return self._direct_inner_sides
        except AttributeError:
            self._direct_inner_sides = cast(Sides, self._get_style_inst("sides"))
        return self._direct_inner_sides

    @property
    def prop_inner_padding(self) -> InnerPadding | None:
        """Gets Padding Instance"""
        try:
            return self._direct_inner_padding
        except AttributeError:
            self._direct_inner_padding = cast(InnerPadding, self._get_style_inst("padding"))
        return self._direct_inner_padding

    @property
    def prop_inner_shadow(self) -> InnerShadow | None:
        """Gets Shadow Instance"""
        try:
            return self._direct_inner_shadow
        except AttributeError:
            self._direct_inner_shadow = cast(InnerShadow, self._get_style_inst("shadow"))
        return self._direct_inner_shadow

    @property
    def default(self: _TBorders) -> _TBorders:  # type: ignore[misc]
        """Gets Default Border."""
        # pylint: disable=unexpected-keyword-arg
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

            self._default_inst = self.__class__(
                all=Side().empty,
                padding=padding,
                shadow=shadow,
                _cattribs=self._get_internal_cattribs(),
            )
            self._default_inst._is_default_inst = True
        return self._default_inst

    @property
    def empty(self: _TBorders) -> _TBorders:  # type: ignore[misc]
        """Gets Empty Border. When style is applied formatting is removed."""
        return self.default

    # endregion Properties
