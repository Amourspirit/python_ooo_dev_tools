# region Imports
from __future__ import annotations
from typing import NamedTuple
import uno

from ooodev.exceptions import ex as mEx
from ooodev.format.calc.style.page.kind.calc_style_page_kind import CalcStylePageKind
from ooodev.format.inner.modify.calc.cell_style_base import CellStyleBase
from ooodev.utils import props as mProps

# endregion Imports


class OrderProps(NamedTuple):
    top_btm: str
    first_pg: str


class Order(CellStyleBase):
    """
    Page Style Order.

    .. seealso::

        - :ref:`help_calc_format_modify_page_sheet`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        top_btm: bool | None = None,
        first_pg: int | None = None,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            top_btm (bool, optional): Specifies page order. ``True`` for Top to Bottom, then right, ``False`` for Left to right then down.
            first_pg (int, optional): Specifies first page number. Set to ``0`` for no page number.
            style_name (CalcStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_page_sheet`
        """

        super().__init__(style_name=style_name, style_family=style_family)
        if top_btm is not None:
            self.prop_top_btm = top_btm
        if first_pg is not None:
            self.prop_first_pg = first_pg

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> Order:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (CalcStylePageKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Order: ``Order`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        style_props = inst.get_style_props(doc)
        if not inst._is_valid_obj(style_props):
            raise mEx.NotSupportedError(f"Object is not support to convert to {cls.__name__}")

        inst.prop_top_btm = bool(mProps.Props.get(style_props, inst._props.top_btm))
        inst.prop_first_pg = int(mProps.Props.get(style_props, inst._props.first_pg))
        return inst

    # region Properties
    @property
    def prop_top_btm(self) -> bool | None:
        """
        Gets/Sets page order. ``True`` for Top to Bottom, then right, ``False`` for Left to right then down.
        """
        return self._get(self._props.top_btm)

    @prop_top_btm.setter
    def prop_top_btm(self, value: bool | None):
        if value is None:
            self._remove(self._props.top_btm)
            return
        self._set(self._props.top_btm, value)

    @property
    def prop_first_pg(self) -> int | None:
        """
        Gets/Sets first page number. Set to ``0`` for no page number.
        """
        return self._get(self._props.first_pg)

    @prop_first_pg.setter
    def prop_first_pg(self, value: int | None):
        if value is None:
            self._remove(self._props.first_pg)
            return
        # zero for no page number
        value = max(value, 0)
        self._set(self._props.first_pg, value)

    @property
    def _props(self) -> OrderProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = OrderProps(top_btm="PrintDownFirst", first_pg="FirstPageNumber")
        return self._props_internal_attributes

    # endregion Properties
