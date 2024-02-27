# region Import
from __future__ import annotations
from typing import cast

from ooodev.format.calc.style.page.kind.calc_style_page_kind import CalcStylePageKind
from ooodev.format.inner.direct.write.page.page.paper_format import PaperFormat as InnerPaperFormat
from ooodev.format.inner.modify.calc.cell_style_base_multi import CellStyleBaseMulti
from ooodev.format.inner.preset.preset_paper_format import PaperFormatKind
from ooodev.utils.data_type.size_mm import SizeMM

# endregion Import


class PaperFormat(CellStyleBaseMulti):
    """
    Page Style Paper Format

    .. seealso::

        - :ref:`help_calc_format_modify_page_page`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        size: SizeMM = SizeMM(215.9, 279.4),
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            size (SizeMM, optional): Width and height in ``mm`` units. Defaults to Letter size in Portrait mode.
            style_name (CalcStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_page_page`
        """

        direct = InnerPaperFormat(size=size)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> PaperFormat:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (CalcStylePageKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            PaperFormat: ``PaperFormat`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerPaperFormat.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @classmethod
    def from_preset(
        cls,
        preset: PaperFormatKind,
        landscape: bool = False,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> PaperFormat:
        """
        Gets instance from preset

        Args:
            preset (PaperFormatKind): Preset kind
            landscape (bool, optional): Specifies if the preset is in landscape mode. Defaults to ``False``.
            style_name (CalcStylePageKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            PaperFormat: Format from preset
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerPaperFormat.from_preset(preset=preset, landscape=landscape)
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | CalcStylePageKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerPaperFormat:
        """Gets/Sets Inner Paper Format instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerPaperFormat, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerPaperFormat) -> None:
        if not isinstance(value, InnerPaperFormat):
            raise TypeError(f'Expected type of InnerPaperFormat, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
