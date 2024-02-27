# region Imports
from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno

from ooodev.format.inner.modify.calc.cell_style_base_multi import CellStyleBaseMulti
from ooodev.format.calc.style.cell.kind.style_cell_kind import StyleCellKind
from ooodev.format.inner.direct.calc.alignment.text_align import VertAlignKind
from ooodev.format.inner.direct.calc.alignment.text_align import HoriAlignKind
from ooodev.format.inner.direct.calc.alignment.text_align import TextAlign as InnerTextAlign

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

# endregion Imports


class TextAlign(CellStyleBaseMulti):
    """
    Cell Style Text Align.

    .. seealso::

        - :ref:`help_calc_format_modify_cell_alignment`

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(
        self,
        *,
        hori_align: HoriAlignKind | None = None,
        indent: float | UnitT | None = None,
        vert_align: VertAlignKind | None = None,
        style_name: StyleCellKind | str = StyleCellKind.DEFAULT,
        style_family: str = "CellStyles",
    ) -> None:
        """
        Constructor

        Args:
            hori_align (HoriAlignKind, optional): Specifies Horizontal Alignment.
            indent: (float, UnitT, optional): Specifies indent in ``pt`` (point) units
                or :ref:`proto_unit_obj`. Only used when ``hori_align`` is set to ``HoriAlignKind.LEFT``
            vert_align (VertAdjustKind, optional): Specifies Vertical Alignment.
            style_name (StyleCellKind, str, optional): Specifies the Cell Style that instance applies to.
                Default is Default Cell Style.
            style_family (str, optional): Style family. Default ``CellStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_cell_alignment`
        """

        direct = InnerTextAlign(hori_align=hori_align, indent=indent, vert_align=vert_align)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct)

    # endregion Init

    # region Static Methods
    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleCellKind | str = StyleCellKind.DEFAULT,
        style_family: str = "CellStyles",
    ) -> TextAlign:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleCellKind, str, optional): Specifies the Cell Style that instance applies to.
                Default is Default Cell Style.
            style_family (str, optional): Style family. Default ``CellStyles``.

        Returns:
            TextAlign: ``TextAlign`` instance from style properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerTextAlign.from_obj(obj=inst.get_style_props(doc))
        inst._set_style("direct", direct)
        return inst

    # endregion Static Methods

    # region Properties
    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleCellKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerTextAlign:
        """Gets/Sets Inner Text Align instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerTextAlign, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerTextAlign) -> None:
        if not isinstance(value, InnerTextAlign):
            raise TypeError(f'Expected type of InnerTextAlign, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value)

    # endregion Properties
