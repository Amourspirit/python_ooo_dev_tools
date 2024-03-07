"""
Draw Style Line Spacing.

.. versionadded:: 0.17.12
"""

from __future__ import annotations
from typing import cast, Any, TYPE_CHECKING
import uno

from ooodev.format.draw.direct.para.indent_spacing.line_spacing import LineSpacing as InnerLineSpacing
from ooodev.format.draw.style.kind import DrawStyleFamilyKind
from ooodev.format.draw.style.lookup import FamilyGraphics
from ooodev.utils.kind.line_spacing_mode_kind import ModeKind
from ooodev.format.inner.modify.draw.para_style_base_multi import ParaStyleBaseMulti

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT


class LineSpacing(ParaStyleBaseMulti):
    """
    Line Spacing Style values.

    .. seealso::

        - :ref:`help_draw_format_modify_indent_space_line_spacing`

    .. versionadded:: 0.17.12
    """

    def __init__(
        self,
        *,
        mode: ModeKind | None = None,
        value: int | float | UnitT = 0,
        style_name: str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> None:
        """
        Constructor

        Args:
            mode (ModeKind, optional): Determines the mode that is used to apply units.
            value (Real, UnitT, optional): Value of line spacing. Only applies when ``ModeKind`` is ``PROPORTIONAL``,
                ``AT_LEAST``, ``LEADING``, or ``FIXED``.
            style_name (FamilyGraphics, str, optional): Specifies the Style that instance applies to.
                Default is Default ``standard`` Style.
            style_family (str, DrawStyleFamilyKind, optional): Family Style. Defaults to ``graphics``.

        Returns:
            None:

        Note:
            If ``ModeKind`` is ``SINGLE``, ``LINE_1_15``, ``LINE_1_5``, or ``DOUBLE`` then ``value`` is ignored.

            If ``ModeKind`` is ``AT_LEAST``, ``LEADING``, or ``FIXED`` then ``value`` is a float (``in mm units``)
            or :ref:`proto_unit_obj`

            If ``ModeKind`` is ``PROPORTIONAL`` then value is an int representing percentage.
            For example ``95`` equals ``95%``, ``130`` equals ``130%``

        See Also:
            - :ref:`help_draw_format_modify_indent_space_line_spacing`
        """

        direct = InnerLineSpacing(mode=mode, value=value)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = str(style_family)
        self._set_style("direct", direct)

    @classmethod
    def from_style(
        cls,
        doc: Any,
        style_name: str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> LineSpacing:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (FamilyGraphics, str, optional): Specifies the Style that instance applies to.
                Default is ``FamilyGraphics.DEFAULT_DRAWING_STYLE``.
            style_family (DrawStyleFamilyKind, str, optional): Style family. Default ``DrawStyleFamilyKind.GRAPHICS``.

        Returns:
            LineSpacing: ``LineSpacing`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerLineSpacing.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct)
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | FamilyGraphics):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerLineSpacing:
        """Gets/Sets Inner Font instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerLineSpacing, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerLineSpacing) -> None:
        if not isinstance(value, InnerLineSpacing):
            raise TypeError(f'Expected type of InnerLineSpacing, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
