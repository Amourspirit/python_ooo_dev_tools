"""
Module for managing paragraph line numbers.

.. versionadded:: 0.9.0
"""
from __future__ import annotations

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.common.abstract.abstract_line_number import AbstractLineNumber, LineNumberProps


class LineNum(AbstractLineNumber):
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
