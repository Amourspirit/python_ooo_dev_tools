"""
Modele for managing paragraph line numbrs.

.. versionadded:: 0.9.0
"""
from __future__ import annotations

from .....meta.static_prop import static_prop
from ....kind.format_kind import FormatKind
from ...common.abstract_line_number import AbstractLineNumber, LineNumeProps

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
        return FormatKind.PARA

    @static_prop
    def default() -> LineNum:  # type: ignore[misc]
        """Gets ``LineNum`` default. Static Property."""
        try:
            return LineNum._DEFAULT_INST
        except AttributeError:
            LineNum._DEFAULT_INST = LineNum(0)
            LineNum._DEFAULT_INST._is_default_inst = True
        return LineNum._DEFAULT_INST
