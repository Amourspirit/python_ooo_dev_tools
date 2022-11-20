from __future__ import annotations
from typing import Any
from dataclasses import dataclass
from .formatter_list_item import FormatterListItem
from ..kind.formatter_table_item_kind import FormatterTableItemKind as FormatterTableItemKind


@dataclass
class FormatterTableItem(FormatterListItem):
    item_kind: FormatterTableItemKind = FormatterTableItemKind.NONE

    def _get_stripped(self, val: str, idx_col: int, idx_col_last: int) -> str:
        l_stripped = False
        r_stripped = False
        result = val

        if FormatterTableItemKind.COL_LEFT_STRIP in self.item_kind:
            result = result.lstrip()
            l_stripped = True

        if FormatterTableItemKind.COL_RIGHT_STRIP in self.item_kind:
            result = result.rstrip()
            r_stripped = True

        if l_stripped and r_stripped:
            return result

        if idx_col == 0:
            if FormatterTableItemKind.START_COL_LEFT_STRIP in self.item_kind:
                if not l_stripped:
                    result = result.lstrip()
                    l_stripped = True
            if FormatterTableItemKind.START_COL_RIGHT_STRIP in self.item_kind:
                if not r_stripped:
                    result = result.rstrip()
                    r_stripped = True

        if l_stripped and r_stripped:
            return result

        if idx_col_last == idx_col:
            if FormatterTableItemKind.END_COL_LEFT_STRIP in self.item_kind:
                if not l_stripped:
                    result = result.lstrip()
                    l_stripped = True
            if FormatterTableItemKind.END_COL_RIGHT_STRIP in self.item_kind:
                if not r_stripped:
                    result = result.rstrip()
                    r_stripped = True

        return result

    def _apply_single_format(self, value, fmt: str) -> str:
        try:
            return format(value, fmt)
        except Exception:
            return str(value)

    def _apply_all_formats(self, value) -> str:
        if isinstance(self.format, tuple):
            v = value
            for fmt in self.format:
                v = self._apply_single_format(v, fmt)
            return str(v)
        else:
            return self._apply_single_format(value, self.format)

    def get_formatted(self, val: Any, idx_col: int, idx_col_last: int) -> str:
        result = self._apply_all_formats(val)
        result = self._get_stripped(result, idx_col=idx_col, idx_col_last=idx_col_last)
        return result
