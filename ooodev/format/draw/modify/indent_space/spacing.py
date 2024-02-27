"""
Draw Style Spacing.

.. versionadded:: 0.17.12
"""

from __future__ import annotations
from typing import cast, Any, TYPE_CHECKING
import uno

from ooodev.format.draw.style.kind import DrawStyleFamilyKind
from ooodev.format.draw.style.lookup import FamilyGraphics
from ooodev.format.inner.modify.draw.para_style_base_multi import ParaStyleBaseMulti
from ooodev.format.inner.direct.write.para.indent_space.spacing import Spacing as InnerSpacing

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT


class Spacing(ParaStyleBaseMulti):
    """
    Spacing Style values.

    .. seealso::

        - :ref:`help_draw_format_modify_indent_space_spacing`

    .. versionadded:: 0.17.12
    """

    def __init__(
        self,
        *,
        above: float | UnitT | None = None,
        below: float | UnitT | None = None,
        style_no_space: bool | None = None,
        style_name: str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> None:
        """
        Constructor

        Args:
            above (float, UnitT, optional): Determines the top margin of the paragraph (in ``mm`` units)
                or :ref:`proto_unit_obj`.
            below (float, UnitT, optional): Determines the bottom margin of the paragraph (in ``mm`` units)
                or :ref:`proto_unit_obj`.
            style_no_space (bool, optional): Do not add space between paragraphs of the same style. See note below.
            style_name (FamilyGraphics, str, optional): Specifies the Style that instance applies to.
                Default is Default ``standard`` Style.
            style_family (str, DrawStyleFamilyKind, optional): Family Style. Defaults to ``graphics``.

        Returns:
            None:

        Note:
            The ``style_no_space`` argument is suppose to set the styles ``ParaContextMargin`` property.
            The ``ParaContextMargin`` is suppose to be part of the ``com.sun.star.style.ParagraphProperties`` service;
            However, for some reason it is missing for Draw styles. Setting this ``style_no_space`` argument will result
            in a print warning message in verbose mode. It is better to not set this argument.
            It is left in just in case it starts working in the future.

            There is a option in the Indent and Spacing dialog ``Do not add space between paragraphs of the same style``.
            It seems to work, but it is not clear how it is implemented. It is not clear if it is a style property.

        See Also:
            - :ref:`help_draw_format_modify_indent_space_spacing`
        """

        direct = InnerSpacing(above=above, below=below, style_no_space=style_no_space)
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
    ) -> Spacing:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (FamilyGraphics, str, optional): Specifies the Style that instance applies to.
                Default is ``FamilyGraphics.DEFAULT_DRAWING_STYLE``.
            style_family (DrawStyleFamilyKind, str, optional): Style family. Default ``DrawStyleFamilyKind.GRAPHICS``.

        Returns:
            Spacing: ``Spacing`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerSpacing.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerSpacing:
        """Gets/Sets Inner Font instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerSpacing, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerSpacing) -> None:
        if not isinstance(value, InnerSpacing):
            raise TypeError(f'Expected type of InnerSpacing, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
