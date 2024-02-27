# region Import
from __future__ import annotations
from typing import Any, cast
import uno
from ooodev.format.writer.style.page.kind.writer_style_page_kind import WriterStylePageKind
from ooodev.utils.data_type.intensity import Intensity
from ooodev.format.inner.direct.write.fill.transparent.transparency import Transparency as InnerTransparency
from ooodev.format.inner.modify.write.page.page_style_base_multi import PageStyleBaseMulti

# endregion Import


class Transparency(PageStyleBaseMulti):
    """
    Page Style Transparency.

    .. seealso::

        - :ref:`help_writer_format_modify_page_transparency`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        value: Intensity | int = 0,
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            value (Intensity, int, optional): Specifies the transparency value from ``0`` to ``100``.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_page_transparency`
        """

        direct = InnerTransparency(value=value)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: Any,
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> Transparency:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (StyleParaKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            Transparency: ``Transparency`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerTransparency.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | WriterStylePageKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerTransparency:
        """Gets/Sets Inner Transparency instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerTransparency, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerTransparency) -> None:
        if not isinstance(value, InnerTransparency):
            raise TypeError(f'Expected type of InnerTransparency, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
