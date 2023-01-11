"""
Modele for managing paragraph Indents and Spacing.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple
from numbers import Real

from ....meta.static_prop import static_prop
from ...style_base import StyleMulti
from ...kind.style_kind import StyleKind
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
        id_before: float | None = None,
        id_after: float | None = None,
        id_first: float | None = None,
        id_auto: bool | None = None,
        sp_above: float | None = None,
        sp_below: float | None = None,
        sp_style_no_space: bool | None = None,
        ln_mode: ModeKind | None = None,
        ln_value: Real | None = None,
        ln_active_spacing: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            id_before (float, optional): Determines the left margin of the paragraph (in mm units).
            id_after (float, optional): Determines the right margin of the paragraph (in mm units).
            id_first (float, optional): specifies the indent for the first line (in mm units).
            id_auto (bool, optional): Determines if the first line should be indented automatically.
            sp_above (float, optional): Determines the top margin of the paragraph (in mm units).
            sp_below (float, optional): Determines the bottom margin of the paragraph (in mm units).
            sp_style_no_space (bool, optional): Do not add space between paragraphs of the same style.
            ln_mode (ModeKind, optional): Determines the left margin of the paragraph (in mm units).
            ln_value (Real, optional): Value of line spacing. Only applies when ``ModeKind`` is ``PORPORTINAL``, ``AT_LEAST``, ``LEADING``, or ``FIXED``.
            ln_active_spacing (bool, optional): Determines active page line-spacing.
        Returns:
            None:

        Note:
            Arguments that start with ``id_`` set Indent.

            Arguments that start with ``sp_`` set Spacing.

            Arguments that start with ``ln_`` set Line Spacing.

            When ``mode`` is ``ModeKind.AT_LEAST``, ``ModeKind.LEADING``, or ``ModeKind.FIXED``
            then the units are mm units (as float).

            When ``mode`` is ``ModeKind.PORPORTINAL`` then the unit is percentage (as int).
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        init_vals = {}

        ls = LineSpacing(mode=ln_mode, value=ln_value, active_ln_spacing=ln_active_spacing)

        spc = Spacing(above=sp_above, below=sp_below, style_no_space=sp_style_no_space)

        indent = Indent(before=id_before, after=id_after, first=id_first, auto=id_auto)

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
    def default() -> IndentSpacing:  # type: ignore[misc]
        """Gets ``IndentSpacing`` default. Static Property."""
        if IndentSpacing._DEFAULT is None:
            ls = LineSpacing.default
            indent = Indent.default
            spc = Spacing.default

            IndentSpacing._DEFAULT = IndentSpacing(
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
            )
        return IndentSpacing._DEFAULT

    # endregion properties
