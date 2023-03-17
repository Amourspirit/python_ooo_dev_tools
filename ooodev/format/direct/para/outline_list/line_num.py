"""
Modele for managing paragraph line numbrs.

.. versionadded:: 0.9.0
"""
from __future__ import annotations

from ....kind.format_kind import FormatKind
from ...common.abstract.abstract_line_number import AbstractLineNumber, LineNumeProps

# from ...events.args.key_val_cancel_args import KeyValCancelArgs


class LineNum(AbstractLineNumber):
    @property
    def _props(self) -> LineNumeProps:
        try:
            return self._props_line_num
        except AttributeError:
            self._props_line_num = LineNumeProps(value="ParaLineNumberStartValue", count="ParaLineNumberCount")
        return self._props_line_num

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA
        return self._format_kind_prop
