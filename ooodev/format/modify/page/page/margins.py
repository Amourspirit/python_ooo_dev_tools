from __future__ import annotations
from typing import cast
import uno
from .....utils.data_type.intensity import Intensity as Intensity
from ....writer.style.page.kind import StylePageKind as StylePageKind
from ..page_style_base_multi import PageStyleBaseMulti
from ....direct.page.page.margins import Margins as DirectMargins


class Margins(PageStyleBaseMulti):
    """
    Page Style Margins

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        left: float | None = None,
        right: float | None = None,
        top: float | None = None,
        bottom: float | None = None,
        gutter: float | None = None,
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            left (float, optional): Left Margin Value in ``mm`` units.
            right (float, optional): Right Margin Value in ``mm`` units.
            top (float, optional): Top Margin Value in ``mm`` units.
            bottom (float, optional): Bottom Margin Value in ``mm`` units.
            gutter (float, optional): Gutter Margin Value in ``mm`` units.
            style_name (StyleParaKind, str, optional): Specifies the Page Style that instance applies to. Deftult is Default Page Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            None:
        """

        direct = DirectMargins(left=left, right=right, top=top, bottom=bottom, gutter=gutter)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> Margins:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            Margins: ``Margins`` instance from document properties.
        """
        inst = super(Margins, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = DirectMargins.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StylePageKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> DirectMargins:
        """Gets Inner Margins instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DirectMargins, self._get_style_inst("direct"))
        return self._direct_inner
