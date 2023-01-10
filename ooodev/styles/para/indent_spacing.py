"""
Modele for managing paragraph Indents and Spacing.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, overload
from numbers import Real

from ...meta.static_prop import static_prop
from ..style_base import StyleMulti
from ..kind.style_kind import StyleKind
from .indent import Indent
from .spacing import Spacing
from .line_spacing import LineSpacing, ModeKind as ModeKind


class IndentSpacing(StyleMulti):
    """
    Paragraph Indents and Spacing

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    # region init

    def __init__(
        self,
        *,
        before: float | None = None,
        after: float | None = None,
        first: float | None = None,
        auto: bool | None = None,
        above: float | None = None,
        below: float | None = None,
        style_no_space: bool | None = None,
        mode: ModeKind | None = None,
        value: Real | None = None,
        active_ln_spacing: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            before (float, optional): Determines the left margin of the paragraph (in mm units).
            after (float, optional): Determines the right margin of the paragraph (in mm units).
            first (float, optional): specifies the indent for the first line (in mm units).
            auto (bool, optional): Determines if the first line should be indented automatically.
            above (float, optional): Determines the top margin of the paragraph (in mm units).
            below (float, optional): Determines the bottom margin of the paragraph (in mm units).
            style_no_space (bool, optional): Do not add space between paragraphs of the same style.
            mode (ModeKind, optional): Determines the left margin of the paragraph (in mm units).
            value (Real, optional): Value of line spacing. Only applies when ``ModeKind`` is ``PORPORTINAL``, ``AT_LEAST``, ``LEADING``, or ``FIXED``.
            active_ln_spacing (bool, optional): Determines active page line-spacing.
        Returns:
            None:

        Note:
            Arguments ``before``, ``after``, ``first`` and ``auto`` set Indent.

            Arguments ``above``, ``below``, and ``style_no_space`` set Spacing.

            Arguments ``mode``, ``value``, and ``active_ln_spacing`` set Line Spacing.

            When ``mode`` is ``ModeKind.AT_LEAST``, ``ModeKind.LEADING``, or ``ModeKind.FIXED``
            then the units are mm units (as float).

            When ``mode`` is ``ModeKind.PORPORTINAL`` then the unit is percentage (as int).
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        init_vals = {}

        ls = LineSpacing(mode=mode, value=value, active_ln_spacing=active_ln_spacing)

        spc = Spacing(above=above, below=below, style_no_space=style_no_space)

        indent = Indent(before=before, after=after, first=first, auto=auto)

        super().__init__(**init_vals)
        if ls.prop_has_attribs:
            self._set_style("line_spacing", ls, *ls.get_attrs())
        if spc.prop_has_attribs:
            self._set_style("spacing", spc, *spc.get_attrs())
        if indent.prop_has_attribs:
            self._set_style("indent", indent, *indent.get_attrs())

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.ParagraphProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.ParagraphProperties",)

    @staticmethod
    def from_obj(obj: object) -> IndentSpacing:
        """
        Gets instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support  ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            IndentSpacing: ``IndentSpacing`` instance that represents ``obj`` Indents and spacing.
        """
        ids = IndentSpacing()
        ls = LineSpacing.from_obj(obj)
        if ls.prop_has_attribs:
            ids._set_style("line_spacing", ls, *ls.get_attrs())
        spc = Spacing.from_obj(obj)
        if spc.prop_has_attribs:
            ids._set_style("spacing", spc, *spc.get_attrs())
        indent = Indent.from_obj(obj)
        if indent.prop_has_attribs:
            ids._set_style("indent", indent, *indent.get_attrs())
        return ids

    # endregion methods

    # region properties
    @property
    def prop_style_kind(self) -> StyleKind:
        """Gets the kind of style"""
        return StyleKind.PARA

    @static_prop
    def default(cls) -> IndentSpacing:  # type: ignore[misc]
        """Gets ``IndentSpacing`` default. Static Property."""
        if cls._DEFAULT is None:
            ls = LineSpacing.default
            indent = Indent.default
            spc = Spacing.default

            cls._DEFAULT = IndentSpacing(
                mode=ls.prop_mode,
                value=ls.prop_value,
                active_ln_spacing=ls.prop_active_ln_spacing,
                above=spc.prop_above,
                below=spc.prop_below,
                style_no_space=spc.prop_style_no_space,
                before=indent.prop_before,
                after=indent.prop_after,
                first=indent.prop_first,
                auto=indent.prop_auto,
            )
        return cls._DEFAULT

    # endregion properties
