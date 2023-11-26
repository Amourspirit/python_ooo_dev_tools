# region Import
from __future__ import annotations
from typing import Any, cast
from ooodev.units import UnitT
from ooodev.format.writer.style.page.kind import WriterStylePageKind as WriterStylePageKind
from ooodev.format.inner.direct.write.page.page.margins import Margins as InnerMargins
from ..page_style_base_multi import PageStyleBaseMulti

# endregion Import


class Margins(PageStyleBaseMulti):
    """
    Page Style Margins

    .. seealso::

        - :ref:`help_writer_format_modify_page_page`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        left: float | UnitT | None = None,
        right: float | UnitT | None = None,
        top: float | UnitT | None = None,
        bottom: float | UnitT | None = None,
        gutter: float | UnitT | None = None,
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
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
            style_name (WriterStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_page_page`
        """

        direct = InnerMargins(left=left, right=right, top=top, bottom=bottom, gutter=gutter)
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
    ) -> Margins:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (WriterStylePageKind, str, optional): Specifies the Paragraph Style that instance applies to.
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
    def prop_style_name(self, value: str | WriterStylePageKind):
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
