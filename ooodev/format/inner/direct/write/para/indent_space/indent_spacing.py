"""
Module for managing paragraph Indents and Spacing.

.. versionadded:: 0.9.0
"""

from __future__ import annotations
from typing import Any, Tuple, cast, Type, overload, TypeVar
from numbers import Real

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.direct.write.para.indent_space.indent import Indent
from ooodev.format.inner.direct.write.para.indent_space.line_spacing import LineSpacing
from ooodev.utils.kind.line_spacing_mode_kind import ModeKind
from ooodev.format.inner.direct.write.para.indent_space.spacing import Spacing
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.units.unit_obj import UnitT

_TIndentSpacing = TypeVar("_TIndentSpacing", bound="IndentSpacing")


class IndentSpacing(StyleMulti):
    """
    Paragraph Indents and Spacing

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        *,
        id_before: float | UnitT | None = None,
        id_after: float | UnitT | None = None,
        id_first: float | UnitT | None = None,
        id_auto: bool | None = None,
        sp_above: float | UnitT | None = None,
        sp_below: float | UnitT | None = None,
        sp_style_no_space: bool | None = None,
        ln_mode: ModeKind | None = None,
        ln_value: Real | None = None,
        ln_active_spacing: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            id_before (float, optional): Determines the left margin of the paragraph (in ``mm`` units) or :ref:`proto_unit_obj`.
            id_after (float, optional): Determines the right margin of the paragraph (in ``mm`` units) or :ref:`proto_unit_obj`.
            id_first (float, optional): specifies the indent for the first line (in ``mm`` units) or :ref:`proto_unit_obj`.
            id_auto (bool, optional): Determines if the first line should be indented automatically.
            sp_above (float, optional): Determines the top margin of the paragraph (in ``mm`` units) or :ref:`proto_unit_obj`.
            sp_below (float, optional): Determines the bottom margin of the paragraph (in ``mm`` units) or :ref:`proto_unit_obj`.
            sp_style_no_space (bool, optional): Do not add space between paragraphs of the same style.
            ln_mode (ModeKind, optional): mode (ModeKind, optional): Determines the mode that is use to apply units.
            ln_value (Real, optional): Value of line spacing. Only applies when ``ModeKind`` is ``PROPORTIONAL``, ``AT_LEAST``, ``LEADING``, or ``FIXED``.
            ln_active_spacing (bool, optional): Determines active page line-spacing.
        Returns:
            None:

        Note:
            Arguments that start with ``id_`` set Indent.

            Arguments that start with ``sp_`` set Spacing.

            Arguments that start with ``ln_`` set Line Spacing.

            When ``mode`` is ``ModeKind.AT_LEAST``, ``ModeKind.LEADING``, or ``ModeKind.FIXED``
            then the units are mm units (as float).

            When ``mode`` is ``ModeKind.PROPORTIONAL`` then the unit is percentage (as int).
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html

        ls = LineSpacing(mode=ln_mode, value=ln_value, active_ln_spacing=ln_active_spacing)  # type: ignore

        spc = Spacing(above=sp_above, below=sp_below, style_no_space=sp_style_no_space)

        indent = Indent(before=id_before, after=id_after, first=id_first, auto=id_auto)

        super().__init__()
        if ls.prop_has_attribs:
            self._set_style("line_spacing", ls, *ls.get_attrs())  # type: ignore
        if spc.prop_has_attribs:
            self._set_style("spacing", spc, *spc.get_attrs())  # type: ignore
        if indent.prop_has_attribs:
            self._set_style("indent", indent, *indent.get_attrs())  # type: ignore

    # endregion init

    # region methods

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.ParagraphProperties",
                "com.sun.star.style.ParagraphStyle",
            )
        return self._supported_services_values

    def _on_modifying(self, source: Any, event_args: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event_args)

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TIndentSpacing], obj: Any) -> _TIndentSpacing: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TIndentSpacing], obj: Any, **kwargs) -> _TIndentSpacing: ...

    @classmethod
    def from_obj(cls: Type[_TIndentSpacing], obj: Any, **kwargs) -> _TIndentSpacing:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            IndentSpacing: ``IndentSpacing`` instance that represents ``obj`` Indents and spacing.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError("Object is not supported for conversion to IndentSpacing")
        ls = LineSpacing.from_obj(obj)
        if ls.prop_has_attribs:
            inst._set_style("line_spacing", ls, *ls.get_attrs())
        spc = Spacing.from_obj(obj)
        if spc.prop_has_attribs:
            inst._set_style("spacing", spc, *spc.get_attrs())
        indent = Indent.from_obj(obj)
        if indent.prop_has_attribs:
            inst._set_style("indent", indent, *indent.get_attrs())
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()

    # endregion methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA
        return self._format_kind_prop

    @property
    def prop_inner_line_spacing(self) -> LineSpacing | None:
        """Gets Line Spacing instance"""
        try:
            return self._direct_inner_ls
        except AttributeError:
            self._direct_inner_ls = cast(LineSpacing, self._get_style_inst("line_spacing"))
        return self._direct_inner_ls

    @property
    def prop_inner_spacing(self) -> Spacing | None:
        """Gets Spacing instance"""
        try:
            return self._direct_inner_spacing
        except AttributeError:
            self._direct_inner_spacing = cast(Spacing, self._get_style_inst("spacing"))
        return self._direct_inner_spacing

    @property
    def prop_inner_indent(self) -> Indent | None:
        """Gets Indent instance"""
        try:
            return self._direct_inner_indent
        except AttributeError:
            self._direct_inner_indent = cast(Indent, self._get_style_inst("indent"))
        return self._direct_inner_indent

    @property
    def default(self: _TIndentSpacing) -> _TIndentSpacing:
        """Gets ``IndentSpacing`` default."""
        try:
            return self._default_inst
        except AttributeError:
            ls = LineSpacing().default
            indent = Indent().default
            spc = Spacing().default
            # pylint: disable=unexpected-keyword-arg
            self._default_inst = self.__class__(
                ln_mode=ls.prop_mode,
                ln_value=ls.prop_value,
                ln_active_spacing=ls.prop_active_ln_spacing,
                sp_above=spc.prop_above,
                sp_below=spc.prop_below,
                sp_style_no_space=spc.prop_style_no_space,
                id_before=indent.prop_before,
                id_after=indent.prop_after,
                id_first=indent.prop_first,
                id_auto=indent.prop_auto,
                _cattribs=self._get_internal_cattribs(),  # type: ignore
            )
            # pylint: disable=protected-access
            self._default_inst._is_default_inst = True
        return self._default_inst

    # endregion properties
