from __future__ import annotations
import uno

from ooodev.format.inner.direct.calc.border.padding import Padding as DirectPadding
from ooodev.units.unit_obj import UnitObj


class Padding(DirectPadding):
    """
    Border Padding

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. seealso::

        - :ref:`help_calc_format_modify_cell_borders`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        left: float | UnitObj | None = None,
        right: float | UnitObj | None = None,
        top: float | UnitObj | None = None,
        bottom: float | UnitObj | None = None,
        all: float | UnitObj | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (float, UnitObj, optional): Left (in ``mm`` units) or :ref:`proto_unit_obj`.
            right (float, UnitObj, optional): Right (in ``mm`` units)  or :ref:`proto_unit_obj`.
            top (float, UnitObj, optional): Top (in ``mm`` units)  or :ref:`proto_unit_obj`.
            bottom (float, UnitObj,  optional): Bottom (in ``mm`` units)  or :ref:`proto_unit_obj`.
            all (float, UnitObj, optional): Left, right, top, bottom (in ``mm`` units)  or :ref:`proto_unit_obj`.
                If argument is present then ``left``, ``right``, ``top``, and ``bottom`` arguments are ignored.

        Raises:
            ValueError: If any argument value is less than zero.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_cell_borders`
        """
        super().__init__(left=left, right=right, top=top, bottom=bottom, all=all)
