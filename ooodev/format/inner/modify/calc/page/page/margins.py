# region Import
from __future__ import annotations
from typing import cast
from ooodev.units import UnitObj
from ooodev.format.calc.style.page.kind import CalcStylePageKind as CalcStylePageKind
from ooodev.format.inner.direct.calc.page.page.margins import Margins as InnerMargins
from ...cell_style_base_multi import CellStyleBaseMulti

# endregion Import


class Margins(CellStyleBaseMulti):
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
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            left (float, optional): Left Margin Value in ``mm`` units  or :ref:`proto_unit_obj`.
            right (float, optional): Right Margin Value in ``mm`` units  or :ref:`proto_unit_obj`.
            top (float, optional): Top Margin Value in ``mm`` units  or :ref:`proto_unit_obj`.
            bottom (float, optional): Bottom Margin Value in ``mm`` units  or :ref:`proto_unit_obj`.
            style_name (CalcStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:
        """

        direct = InnerMargins(left=left, right=right, top=top, bottom=bottom)
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
    ) -> Margins:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (CalcStylePageKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

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
    def prop_style_name(self, value: str | CalcStylePageKind):
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
