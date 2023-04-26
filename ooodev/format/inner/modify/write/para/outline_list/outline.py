# region Import
from __future__ import annotations
from typing import cast
import uno

from ooodev.format.writer.style.para.kind import StyleParaKind as StyleParaKind
from ooodev.format.writer.style.char.kind import StyleCharKind as StyleCharKind
from ooodev.format.inner.direct.write.para.outline_list.outline import Outline as InnerOutline
from ooodev.format.inner.direct.write.para.outline_list.outline import LevelKind as LevelKind
from ..para_style_base_multi import ParaStyleBaseMulti

# endregion Import


class Outline(ParaStyleBaseMulti):
    """
    Paragraph Style Outline

    .. seealso::

        - :ref:`help_writer_format_modify_para_outline_and_list`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        level: LevelKind = LevelKind.TEXT_BODY,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            level (LevelKind): Outline level.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_para_outline_and_list`
        """

        direct = InnerOutline(level=level)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> Outline:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            Outline: ``Outline`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerOutline.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleParaKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerOutline:
        """Gets/Sets Inner Outline instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerOutline, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerOutline) -> None:
        if not isinstance(value, InnerOutline):
            raise TypeError(f'Expected type of InnerOutline, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
