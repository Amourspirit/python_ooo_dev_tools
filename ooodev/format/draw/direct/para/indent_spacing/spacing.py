"""
Module for managing shape paragraph spacing.

.. versionadded:: 0.17.8
"""

# region Import
from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.format.inner.direct.write.para.indent_space.spacing import Spacing as WriteSpacing


if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

# endregion Import


class Spacing(WriteSpacing):
    """
    Paragraph Spacing

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. seealso::

        - :ref:`help_writer_format_direct_para_indent_spacing`

    .. versionadded:: 0.17.8
    """

    # region init

    def __init__(
        self,
        *,
        above: float | UnitT | None = None,
        below: float | UnitT | None = None,
        style_no_space: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            above (float, UnitT, optional): Determines the top margin of the paragraph (in ``mm`` units)
                or :ref:`proto_unit_obj`.
            below (float, UnitT, optional): Determines the bottom margin of the paragraph (in ``mm`` units)
                or :ref:`proto_unit_obj`.
            style_no_space (bool, optional): Do not add space between paragraphs of the same style.

        Returns:
            None:

        See Also:

            - :ref:`help_writer_format_direct_para_indent_spacing`
        """
        super().__init__(above=above, below=below, style_no_space=style_no_space)

    # endregion init
