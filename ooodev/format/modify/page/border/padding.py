from __future__ import annotations
from typing import Tuple, cast
import uno
from ....writer.style.page.kind.style_page_kind import StylePageKind as StylePageKind
from ..page_style_base_multi import PageStyleBaseMulti

from ....direct.para.border.padding import Padding as DirectPadding


class Padding(PageStyleBaseMulti):
    """
    Page Style Border Padding.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        left: float | None = None,
        right: float | None = None,
        top: float | None = None,
        bottom: float | None = None,
        padding_all: float | None = None,
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            left (float, optional): Paragraph left padding (in mm units).
            right (float, optional): Paragraph right padding (in mm units).
            top (float, optional): Paragraph top padding (in mm units).
            bottom (float, optional): Paragraph bottom padding (in mm units).
            padding_all (float, optional): Paragraph left, right, top, bottom padding (in mm units). If argument is present then ``left``, ``right``, ``top``, and ``bottom`` arguments are ignored.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            None:
        """

        direct = DirectPadding(left=left, right=right, top=top, bottom=bottom, padding_all=padding_all)
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
    ) -> Padding:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            Padding: ``Padding`` instance from document properties.
        """
        inst = super(Padding, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = DirectPadding.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> DirectPadding:
        """Gets Inner Padding instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DirectPadding, self._get_style_inst("direct"))
        return self._direct_inner
