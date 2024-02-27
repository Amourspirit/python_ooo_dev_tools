from __future__ import annotations
from dataclasses import dataclass
from ooodev.formatters.format_list_item import FormatListItem
from ooodev.formatters.string_kind import StringKind as StringKind


@dataclass
class FormatString(FormatListItem):
    row_kind: StringKind = StringKind.NONE

    def _get_stripped(self, val: str) -> str:
        result = val

        if StringKind.LEFT_STRIP in self.row_kind:
            result = result.lstrip()

        return result

    def _apply_single_format(self, value, fmt: str) -> str:
        try:
            return format(value, fmt)
        except Exception:
            return str(value)

    def _apply_all_formats(self, value) -> str:
        if not isinstance(self.format, tuple):
            return self._apply_single_format(value, self.format)
        v = value
        for fmt in self.format:
            v = self._apply_single_format(v, fmt)
        return str(v)

    def get_formatted(self, val: str) -> str:
        result = self._get_stripped(val)
        result = self._apply_all_formats(result)
        return result


__all__ = ["FormatString"]
