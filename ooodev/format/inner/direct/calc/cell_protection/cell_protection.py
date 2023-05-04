from __future__ import annotations
from ...structs.cell_protection_struct import CellProtectionStruct


class CellProtection(CellProtectionStruct):
    """
    Cell Protection.

    Warning:
        Cell protection is only effective after the sheet is has been applied to is also been protected.

    .. seealso::

        - :ref:`help_calc_format_direct_cell_cell_protection`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self, hide_all: bool = False, protected: bool = False, hide_formula: bool = False, hide_print: bool = False
    ) -> None:
        """
        Constructor

        Args:
            hide_all (bool, optional): Specifies if all is hidden. Defaults to ``False``.
            protected (bool, optional): Specifies protected value. Defaults to ``False``.
            hide_formula (bool, optional): Specifies if the formula is hidden. Defaults to ``False``.
            hide_print (bool, optional): Specifies if the cell are to be omitted during print. Defaults to ``False``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_direct_cell_cell_protection`
        """
        super().__init__(hide_all=hide_all, protected=protected, hide_formula=hide_formula, hide_print=hide_print)
