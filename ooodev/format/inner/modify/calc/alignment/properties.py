# region Imports
from __future__ import annotations
from typing import cast
import uno

from ooodev.format.inner.modify.calc.cell_style_base_multi import CellStyleBaseMulti
from ooodev.format.calc.style.cell.kind.style_cell_kind import StyleCellKind
from ooodev.format.inner.direct.calc.alignment.properties import TextDirectionKind
from ooodev.format.inner.direct.calc.alignment.properties import Properties as InnerProperties

# endregion Imports


class Properties(CellStyleBaseMulti):
    """
    Cell Style Properties.

    .. seealso::

        - :ref:`help_calc_format_modify_cell_alignment`

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(
        self,
        *,
        wrap_auto: bool | None = None,
        hyphen_active: bool | None = None,
        shrink_to_fit: bool | None = None,
        direction: TextDirectionKind | None = None,
        style_name: StyleCellKind | str = StyleCellKind.DEFAULT,
        style_family: str = "CellStyles",
    ) -> None:
        """
        Constructor

        Args:
            wrap_auto (bool, optional): Specifies wrap text automatically.
            hyphen_active (bool, optional): Specifies hyphenation active.
            shrink_to_fit (bool, optional): Specifies if text will shrink to cell.
            direction (TextDirectionKind, optional): Specifies Text Direction.
            style_name (StyleCellKind, str, optional): Specifies the Cell Style that instance applies to.
                Default is Default Cell Style.
            style_family (str, optional): Style family. Default ``CellStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_cell_alignment`
        """

        direct = InnerProperties(
            wrap_auto=wrap_auto, hyphen_active=hyphen_active, shrink_to_fit=shrink_to_fit, direction=direction
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct)  # type: ignore

    # endregion Init

    # region Static Methods
    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleCellKind | str = StyleCellKind.DEFAULT,
        style_family: str = "CellStyles",
    ) -> Properties:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleCellKind, str, optional): Specifies the Cell Style that instance applies to.
                Default is Default Cell Style.
            style_family (str, optional): Style family. Default ``CellStyles``.

        Returns:
            Properties: ``Properties`` instance from style properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerProperties.from_obj(obj=inst.get_style_props(doc))
        inst._set_style("direct", direct)  # type: ignore
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
    def prop_inner(self) -> InnerProperties:
        """Gets/Sets Inner Properties instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerProperties, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerProperties) -> None:
        if not isinstance(value, InnerProperties):
            raise TypeError(f'Expected type of InnerProperties, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value)

    # endregion Properties
