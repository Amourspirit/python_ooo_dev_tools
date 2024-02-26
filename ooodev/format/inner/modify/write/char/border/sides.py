# region Imports
from __future__ import annotations
from typing import cast

import uno

from ooodev.format.writer.style.char.kind.style_char_kind import StyleCharKind
from ooodev.format.inner.modify.write.char.char_style_base_multi import CharStyleBaseMulti
from ooodev.format.inner.direct.write.char.border.sides import Sides as InnerSides
from ooodev.format.inner.direct.structs.side import Side

# endregion Imports


class Sides(CharStyleBaseMulti):
    """
    Character Style Border Sides (lines).

    .. seealso::

        - :ref:`help_writer_format_modify_char_borders`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        left: Side | None = None,
        right: Side | None = None,
        top: Side | None = None,
        bottom: Side | None = None,
        border_side: Side | None = None,
        style_name: StyleCharKind | str = StyleCharKind.STANDARD,
        style_family: str = "CharacterStyles",
    ) -> None:
        """
        Constructor

        Args:
            left (Side | None, optional): Determines the line style at the left edge.
            right (Side | None, optional): Determines the line style at the right edge.
            top (Side | None, optional): Determines the line style at the top edge.
            bottom (Side | None, optional): Determines the line style at the bottom edge.
            border_side (Side | None, optional): Determines the line style at the top, bottom, left, right edges. If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
            shadowed (bool, optional): Specifies if the characters are formatted and displayed with a shadow effect.
            style_name (StyleParaKind, str, optional): Specifies the Character Style that instance applies to. Default is Default Character Style.
            style_family (str, optional): Style family. Default ``CharacterStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_char_borders`
        """

        direct = InnerSides(left=left, right=right, top=top, bottom=bottom, all=border_side)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleCharKind | str = StyleCharKind.STANDARD,
        style_family: str = "CharacterStyles",
    ) -> Sides:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleCharKind, str, optional): Specifies the Character Style that instance applies to. Default is Default Character Style.
            style_family (str, optional): Style family. Default ``CharacterStyles``.

        Returns:
            Sides: ``Sides`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerSides.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleCharKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerSides:
        """Gets/Sets Inner Sides instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerSides, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerSides) -> None:
        if not isinstance(value, InnerSides):
            raise TypeError(f'Expected type of InnerSides, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
