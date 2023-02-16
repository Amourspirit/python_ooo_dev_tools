from __future__ import annotations
from typing import cast
import uno
from ....direct.common.format_types.size_mm import SizeMM as SizeMM
from .....utils.data_type.intensity import Intensity as Intensity
from ....writer.style.page.kind import StylePageKind as StylePageKind
from ..page_style_base_multi import PageStyleBaseMulti
from ....preset.preset_paper_format import PaperFormatKind as PaperFormatKind
from ....direct.page.page.paper_format import PaperFormat as DirectPaperFormat


class PaperFormat(PageStyleBaseMulti):
    """
    Page Style Paper Format

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        size: SizeMM = SizeMM(215.9, 279.4),
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            size (SizeMM, optional): Width and height in ``mm`` units. Defaults to Letter size in Portrait mode.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            None:
        """

        direct = DirectPaperFormat(size=size)
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
    ) -> PaperFormat:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            PaperFormat: ``PaperFormat`` instance from document properties.
        """
        inst = super(PaperFormat, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = DirectPaperFormat.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @classmethod
    def from_preset(
        cls,
        preset: PaperFormatKind,
        landscape: bool = False,
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> PaperFormat:
        """
        Gets instance from preset

        Args:
            preset (PaperFormatKind): Preset kind
            landscape (bool, optional): Specifies if the preset is in landscape mode. Defaults to ``False``.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            PaperFormat: Format from preset
        """
        inst = super(PaperFormat, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = DirectPaperFormat.from_preset(preset=preset, landscape=landscape)
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
    def prop_inner(self) -> DirectPaperFormat:
        """Gets Inner Paper Format instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DirectPaperFormat, self._get_style_inst("direct"))
        return self._direct_inner
