# region Imports
from __future__ import annotations
from typing import cast
import uno

from ooodev.units.angle import Angle
from ooodev.format.inner.modify.calc.cell_style_base_multi import CellStyleBaseMulti
from ooodev.format.calc.style.cell.kind.style_cell_kind import StyleCellKind
from ooodev.format.inner.direct.calc.alignment.text_orientation import EdgeKind
from ooodev.format.inner.direct.calc.alignment.text_orientation import TextOrientation as InnerTextOrientation

# endregion Imports


class TextOrientation(CellStyleBaseMulti):
    """
    Cell Style Text Orientation.

    .. seealso::

        - :ref:`help_calc_format_modify_cell_alignment`

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(
        self,
        *,
        vert_stack: bool | None = None,
        rotation: int | Angle | None = None,
        edge: EdgeKind | None = None,
        style_name: StyleCellKind | str = StyleCellKind.DEFAULT,
        style_family: str = "CellStyles",
    ) -> None:
        """
        Constructor

        Args:
            vert_stack (bool, optional): Specifies if vertical stack is to be used.
            rotation (int, Angle, optional): Specifies if the rotation.
            edge (EdgeKind, optional): Specifies the Reference Edge.
            style_name (StyleCellKind, str, optional): Specifies the Cell Style that instance applies to.
                Default is Default Cell Style.
            style_family (str, optional): Style family. Default ``CellStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_cell_alignment`
        """

        direct = InnerTextOrientation(vert_stack=vert_stack, rotation=rotation, edge=edge)
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
    ) -> TextOrientation:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleCellKind, str, optional): Specifies the Cell Style that instance applies to. Default is Default Cell Style.
            style_family (str, optional): Style family. Default ``CellStyles``.

        Returns:
            TextOrientation: ``TextOrientation`` instance from style properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerTextOrientation.from_obj(obj=inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerTextOrientation:
        """Gets/Sets Inner Inner Text Orientation instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerTextOrientation, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerTextOrientation) -> None:
        if not isinstance(value, InnerTextOrientation):
            raise TypeError(f'Expected type of InnerTextOrientation, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value)

    # endregion Properties
