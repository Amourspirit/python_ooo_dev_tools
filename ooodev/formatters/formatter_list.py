from __future__ import annotations
from typing import List, Any, Tuple
from dataclasses import dataclass, field
from ooodev.utils.validation import check
from ooodev.formatters.only_ignore_kind import OnlyIgnoreKind as OnlyIgnoreKind
from ooodev.formatters.format_list_item import FormatListItem as FormatListItem


@dataclass
class FormatterList:
    format: str | Tuple[str, ...]
    """
    Format option such as ``.2f``
    
    Multiple formats can be added such as ``(".2f", "<10")``.
    Formats are applied in the order they are added.
    In this case first float is formatted as string with two decimal places, and
    then value is padded to the right with spaces.
    """
    idx_rule: OnlyIgnoreKind = field(default=OnlyIgnoreKind.IGNORE)
    """
    Determines what indexes are affected.
    """
    idxs: Tuple[int, ...] = field(default_factory=tuple)
    """
    Indexes to apply formatting to or ignore, depending on ``index_rule``.
    """
    custom_formats: List[FormatListItem] = field(default_factory=list)

    def __post_init__(self) -> None:
        s_str = f"{self}"
        msg = "indexes can only be of type int"
        for index in self.idxs:
            check(isinstance(index, int), s_str, msg)

    def _custom_format(self, current_index: int) -> FormatListItem | None:
        if current_index < 0:
            return None
        return next((cf for cf in self.custom_formats if cf.has_index(current_index)), None)

    def apply_format(self, val: Any, current_index: int = -1) -> str:
        """
        Applies formatting to val if ``format`` has been set

        Args:
            val (Any): Any value
            current_index (int, optional): When value is equal or greater then ``0`` then
                formatting is applied following rules for indexes.

        Returns:
            str: Formatted string
        """

        def _apply_single_format(value, fmt: str) -> str:
            try:
                return format(value, fmt)
            except Exception:
                return str(value)

        def _apply_all_formats(value, formats: str | Tuple[str, ...]) -> str:
            if isinstance(formats, tuple):
                v = value
                for fmt in formats:
                    v = _apply_single_format(v, fmt)
                return str(v)
            else:
                return _apply_single_format(value, formats)

        cf = self._custom_format(current_index)
        if cf:
            try:
                return _apply_all_formats(val, cf.format)
            except Exception:
                return str(val)
        try:
            if not self.format or self.idx_rule == OnlyIgnoreKind.NONE:
                return str(val)
            if current_index > -1:
                if self.idx_rule == OnlyIgnoreKind.IGNORE and current_index in self.idxs:
                    return str(val)
                if self.idx_rule == OnlyIgnoreKind.ONLY and current_index not in self.idxs:
                    return str(val)
            return _apply_all_formats(val, self.format)
        except Exception:
            return str(val)


__all__ = ["FormatterList"]
