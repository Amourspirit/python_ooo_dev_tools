from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.format.inner.direct.write.para.indent_space.indent import Indent as WriteIndent


if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT


class Indent(WriteIndent):
    """
    Shape Indent

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. seealso::

        - :ref:`help_writer_format_direct_para_indent_spacing`

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        *,
        before: float | UnitT | None = None,
        after: float | UnitT | None = None,
        first: float | UnitT | None = None,
        auto: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            before (float, UnitT, optional): Determines the left margin of the paragraph (in ``mm`` units)
                or :ref:`proto_unit_obj`.
            after (float, UnitT, optional): Determines the right margin of the paragraph (in ``mm`` units)
                or :ref:`proto_unit_obj`.
            first (float, UnitT, optional): specifies the indent for the first line (in ``mm`` units)
                or :ref:`proto_unit_obj`.
            auto (bool, optional): Determines if the first line should be indented automatically.

        Returns:
            None:

        See Also:

            - :ref:`help_writer_format_direct_para_indent_spacing`
        """
        super().__init__(before=before, after=after, first=first, auto=auto)

    # endregion init
