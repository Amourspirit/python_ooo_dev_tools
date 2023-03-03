from __future__ import annotations
from typing import cast
import uno
from .....proto.unit_obj import UnitObj
from .....utils.data_type.intensity import Intensity as Intensity
from ....writer.style.page.kind import StylePageKind as StylePageKind
from ..page_style_base_multi import PageStyleBaseMulti
from ....direct.page.page.margins import Margins as InnerMargins


class Margins(PageStyleBaseMulti):
    """
    Page Style Margins

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        left: float | UnitObj | None = None,
        right: float | UnitObj | None = None,
        top: float | UnitObj | None = None,
        bottom: float | UnitObj | None = None,
        gutter: float | UnitObj | None = None,
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            left (float, optional): Left Margin Value in ``mm`` units  or :ref:`proto_unit_obj`.
            right (float, optional): Right Margin Value in ``mm`` units  or :ref:`proto_unit_obj`.
            top (float, optional): Top Margin Value in ``mm`` units  or :ref:`proto_unit_obj`.
            bottom (float, optional): Bottom Margin Value in ``mm`` units  or :ref:`proto_unit_obj`.
            gutter (float, optional): Gutter Margin Value in ``mm`` units  or :ref:`proto_unit_obj`.
            style_name (StyleParaKind, str, optional): Specifies the Page Style that instance applies to. Deftult is Default Page Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            None:
        """

        direct = InnerMargins(left=left, right=right, top=top, bottom=bottom, gutter=gutter)
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
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerMargins.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerMargins:
        """Gets/Sets Inner Margins instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerMargins, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerMargins) -> None:
        if not isinstance(value, InnerMargins):
            raise TypeError(f'Expected type of InnerMargins, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
