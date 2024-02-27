"""
Module for managing paragraph line numbers.

.. seealso::

    - :ref:`help_writer_format_direct_para_outline_and_list`

.. versionadded:: 0.9.0
"""

from __future__ import annotations

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.common.abstract.abstract_line_number import AbstractLineNumber
from ooodev.format.inner.common.abstract.abstract_line_number import LineNumberProps


class LineNum(AbstractLineNumber):
    """
    Paragraph Line Numbers

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. seealso::

        - :ref:`help_writer_format_direct_para_outline_and_list`

    .. versionadded:: 0.9.0
    """

    def __init__(self, num_start: int = 0) -> None:
        """
        Constructor

        Args:
            num_start (int, optional): Restart paragraph with number.
                If ``0`` then this paragraph is include in line numbering.
                If ``-1`` then this paragraph is excluded in line numbering.
                If greater than zero this paragraph is included in line numbering and the numbering is restarted with
                value of ``num_start``.

        Returns:
            None:

        See Also:

            - :ref:`help_writer_format_direct_para_outline_and_list`
        """
        super().__init__(num_start=num_start)

    @property
    def _props(self) -> LineNumberProps:
        try:
            return self._props_line_num
        except AttributeError:
            self._props_line_num = LineNumberProps(value="ParaLineNumberStartValue", count="ParaLineNumberCount")
        return self._props_line_num

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA
        return self._format_kind_prop
