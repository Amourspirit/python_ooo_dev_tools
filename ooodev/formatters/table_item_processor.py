from __future__ import annotations
from typing import Any, Tuple

from ooodev.formatters.format_table_item import FormatTableItem
from ooodev.formatters.table_item_kind import TableItemKind


class TableItemProcessor:
    """
    Process ``FormatterTableItem`` Instances

    Warning:
        This class is intended as an internal class.
    """

    @staticmethod
    def _get_stripped(itm: FormatTableItem, val: str, idx_col: int, idx_col_last: int) -> str:
        l_stripped = False
        r_stripped = False
        result = val

        if TableItemKind.COL_LEFT_STRIP in itm.item_kind:
            result = result.lstrip()
            l_stripped = True

        if TableItemKind.COL_RIGHT_STRIP in itm.item_kind:
            result = result.rstrip()
            r_stripped = True

        if l_stripped and r_stripped:
            return result

        if idx_col == 0:
            if TableItemKind.START_COL_LEFT_STRIP in itm.item_kind and not l_stripped:
                result = result.lstrip()
                l_stripped = True
            if TableItemKind.START_COL_RIGHT_STRIP in itm.item_kind and not r_stripped:
                result = result.rstrip()
                r_stripped = True

        if l_stripped and r_stripped:
            return result

        if idx_col_last == idx_col:
            if TableItemKind.END_COL_LEFT_STRIP in itm.item_kind and not l_stripped:
                result = result.lstrip()
                l_stripped = True
            if TableItemKind.END_COL_RIGHT_STRIP in itm.item_kind and not r_stripped:
                result = result.rstrip()
                r_stripped = True

        return result

    @staticmethod
    def _is_format(itm: FormatTableItem, idx_row: int, idx: int) -> bool:
        has_rows = bool(itm.row_idxs_exc)
        has_idx = bool(itm.idxs_inc)

        if has_rows and itm.is_row_exc_index(idx_row):
            # excludes take priority over includes
            # explicitly excluded
            return False

        return bool(has_idx and itm.is_index(idx))

    @staticmethod
    def _apply_single_format(value, fmt: str) -> str:
        try:
            return format(value, fmt)
        except Exception:
            return str(value)

    @classmethod
    def _apply_all_formats(cls, itm: FormatTableItem, value: Any) -> str:
        if not isinstance(itm.format, tuple):
            return cls._apply_single_format(value, itm.format)
        v = value
        for fmt in itm.format:
            v = cls._apply_single_format(v, fmt)
        return str(v)

    @classmethod
    def process_row(
        cls, itm: FormatTableItem, idx_row: int, idx_col: int, idx_col_last: int, value: Any
    ) -> Tuple[bool, Any]:
        """
        Process row

        Args:
            itm (FormatterTableItem): Current Instance
            idx_row (int): Current row index
            idx_col (int): Current col index
            idx_col_last (int): Index of last column
            value (Any): Value to format

        Returns:
            Tuple[bool, Any]: ``(True, str)`` if conditions are met; Otherwise, ``(False, Any)``
        """
        if cls._is_format(itm, idx_row=idx_row, idx=idx_row):
            result = cls._apply_all_formats(itm, value)
            result = cls._get_stripped(itm=itm, val=result, idx_col=idx_col, idx_col_last=idx_col_last)
            return (True, result)
        return (False, value)

    @classmethod
    def process_col(
        cls, itm: FormatTableItem, idx_row: int, idx_col: int, idx_col_last: int, value: Any
    ) -> Tuple[bool, Any]:
        """
        Process Col

        Args:
            itm (FormatterTableItem): Current Instance
            idx_row (int): Current row index
            idx_col (int): Current col index
            idx_col_last (int): Index of last column
            value (Any): Value to format

        Returns:
            Tuple[bool, Any]: ``(True, str)`` if conditions are met; Otherwise, ``(False, Any)``
        """
        if cls._is_format(itm, idx_row=idx_row, idx=idx_col):
            result = cls._apply_all_formats(itm, value)
            result = cls._get_stripped(itm=itm, val=result, idx_col=idx_col, idx_col_last=idx_col_last)
            return (True, result)
        return (False, value)
