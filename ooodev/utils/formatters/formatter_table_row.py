from __future__ import annotations
from dataclasses import dataclass
from .formatter_list_item import FormatterListItem
from ..kind.formatter_table_row_kind import FormatterTableRowKind as FormatterTableRowKind


@dataclass
class FormatterTableRow(FormatterListItem):
    row_kind: FormatterTableRowKind = FormatterTableRowKind.NONE

    def _get_stripped(self, val: str) -> str:
        result = val

        if FormatterTableRowKind.LEFT_STRIP in self.row_kind:
            result = result.lstrip()

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

    def get_formatted(self, val: str) -> str:
        result = self._get_stripped(val)
        result = self._apply_all_formats(result)
        return result
