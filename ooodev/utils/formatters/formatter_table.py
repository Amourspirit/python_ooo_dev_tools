from __future__ import annotations
from typing import List, Set, Any, Tuple, TYPE_CHECKING
from ..kind.formatter_table_rule_kind import FormatterTableRuleKind as FormatterTableRuleKind
from ..kind.formatter_table_rule_kind import FormatterTableRuleKind as FormatterTableRuleKind
from .formatter_table_item import FormatterTableItem as FormatterTableItem
from .formatter_table_row import FormatterTableRow as FormatterTableRow

if TYPE_CHECKING:
    from ..type_var import Row


class FormatterTable:
    def __init__(
        self,
        format: str | Tuple[str, ...],
        idx_rule: FormatterTableRuleKind = FormatterTableRuleKind.IGNORE,
        idxs: Tuple[int, ...] | None = None,
    ) -> None:
        self._format = format
        self._idx_rule = idx_rule
        if idxs is None:
            self._idxs = ()
        else:
            self._idxs = idxs
        self._custom_formats_row_items: List[FormatterTableItem] = []
        self._custom_formats_row: List[FormatterTableRow] = []
        self._custom_formats_col: List[FormatterTableItem] = []
        self._col_formats: List[FormatterTableItem] = []
        self._cols = None
        self._idx_col_last = -1

    def _get_cols(self) -> Tuple[int, ...]:
        if self._cols is None:
            cols: Set[int] = set()
            for cf in self.col_formats:
                cols.update(cf.idxs)
            self._cols = tuple(cols)
        return self._cols

    def _has_col(self, idx_col: int) -> bool:
        cols = self._get_cols()
        return idx_col in cols

    def _custom_row_item_format(self, idx_row: int) -> FormatterTableItem | None:
        if idx_row < 0:
            return None
        for cf in self._custom_formats_row_items:
            if cf.has_index(idx_row):
                return cf
        return None

    def _custom_row_format(self, idx_row: int) -> FormatterTableRow | None:
        if idx_row < 0:
            return None
        for cf in self._custom_formats_row:
            if cf.has_index(idx_row):
                return cf
        return None

    def _custom_col_format(self, idx_col: int) -> FormatterTableItem | None:
        if idx_col < 0:
            return None
        for cf in self._custom_formats_col:
            if cf.has_index(idx_col):
                return cf
        return None

    def _get_col_formatter(self, idx_col: int) -> FormatterTableItem | None:
        if idx_col < 0:
            return None
        for cf in self._col_formats:
            if cf.has_index(idx_col):
                return cf
        return None

    def _format_col(self, idx_row: int, idx_col: int, val: Any) -> Tuple[bool, Any]:
        # returns original value or formatting string.
        ccf = self._custom_col_format(idx_col=idx_col)

        do_col_format = True
        if not self._is_format_row(idx_row=idx_row):
            # this row may be ignored, check for col override:
            do_col_format = FormatterTableRuleKind.COL_OVER_ROW in self.idx_rule

        if not do_col_format and ccf:
            # there is a custom column format
            # apply it if col overrides row for custom format
            if FormatterTableRuleKind.CUSTOM_COL_OVER_ROW in self._idx_rule:
                return (True, ccf.get_formatted(val=val, idx_col=idx_col, idx_col_last=self._idx_col_last))

        if ccf:
            # there is a custom column format and column expects formatting.
            # ignore column formatting and apply only custom formatting.
            return (True, ccf.get_formatted(val=val, idx_col=idx_col, idx_col_last=self._idx_col_last))

        if not do_col_format:
            # no custom formatting, no column formatting
            return (False, val)

        cf = self._get_col_formatter(idx_col=idx_col)
        if not cf:
            # no column formatting
            return (False, val)
        # apply column formatting
        return (True, self._apply_all_formats(value=val, formats=cf.format))

    def _is_format_row(self, idx_row: int) -> bool:
        if FormatterTableRuleKind.IGNORE in self._idx_rule:
            return idx_row not in self._idxs
        if FormatterTableRuleKind.ONLY in self._idx_rule:
            return idx_row in self._idxs
        return False

    def _format_row(self, idx_row, row_data: List[str], join_str: str) -> str:
        if len(self._custom_formats_row) == 0:
            return join_str.join(row_data).rstrip()

        custom_fmt_row = self._custom_row_format(idx_row=idx_row)
        if custom_fmt_row:
            row_str = join_str.join(row_data)
            return custom_fmt_row.get_formatted(row_str).rstrip()

        return join_str.join(row_data).rstrip()

    def _format_row_items(self, idx_row: int, row_data: Row, join_str: str) -> str:
        is_row_format = self._is_format_row(idx_row=idx_row)
        fmt_cols_idxs = set()
        # do not apply column formatting to rows that are ignored.
        if is_row_format:
            fmt_cols = []
            for i, val in enumerate(row_data):
                col_formated, col_val = self._format_col(idx_row=idx_row, idx_col=i, val=val)
                fmt_cols.append(col_val)
                if col_formated:
                    fmt_cols_idxs.add(i)
        else:
            fmt_cols = row_data

        custom_row_fmt_itm = self._custom_row_item_format(idx_row=idx_row)
        if custom_row_fmt_itm:
            # there is a custom format for this row.
            # ignore other format for row and apply custom.
            #
            # perhaps later other flag such as CUSTOM_COL_OVER_CUSTOM_ROW should be added.
            # if col over rows then rows should ignore cols that have been formatted
            # for now keep it a simple as possible.
            s_row = []
            for i, val in enumerate(fmt_cols):
                if self._has_col(idx_col=i):
                    s_row.append(val)  # will be string because it is formatted
                else:
                    s_row.append(custom_row_fmt_itm.get_formatted(val=val, idx_col=i, idx_col_last=self._idx_col_last))
            # if FormatterTableRuleKind.COL_OVER_ROW in self._idx_rule:
            #     for i, val in enumerate(fmt_cols):
            #         if self._has_col(idx_col=i):
            #             s_row.append(val)  # will be string because it is formatted
            #         else:
            #             s_row.append(crf.get_formatted(val=val, idx_col=i, idx_col_last=self._idx_col_last))
            # else:
            #     for i, val in enumerate(fmt_cols):
            #         if i in fmt_cols_idxs:
            #             s_row.append(val)  # will be string because it is formatted
            #         else:
            #             s_row.append(self._apply_all_formats(val, self._format))
            return self._format_row(idx_row=idx_row, row_data=s_row, join_str=join_str)

        if is_row_format:
            # if col over rows then rows should ignore cols that have been formatted
            s_row = []
            if FormatterTableRuleKind.COL_OVER_ROW in self._idx_rule:
                for i, val in enumerate(fmt_cols):
                    if i in fmt_cols_idxs:
                        # if self._has_col(idx_col=i):
                        s_row.append(val)  # will be string because it is formatted
                    else:
                        s_row.append(self._apply_all_formats(val, self._format))
            else:
                for i, val in enumerate(fmt_cols):
                    if i in fmt_cols_idxs:
                        s_row.append(val)  # will be string because it is formatted
                    else:
                        s_row.append(self._apply_all_formats(val, self._format))
        else:
            s_row = [str(v) for v in fmt_cols]

        return self._format_row(idx_row=idx_row, row_data=s_row, join_str=join_str)

    def _apply_single_format(self, value, fmt: str) -> str:
        try:
            return format(value, fmt)
        except Exception:
            return str(value)

    def _apply_all_formats(self, value, formats: str | Tuple[str, ...]) -> str:
        if isinstance(formats, tuple):
            v = value
            for fmt in formats:
                v = self._apply_single_format(v, fmt)
            return str(v)
        else:
            return self._apply_single_format(value, formats)

    def get_formatted(self, idx_row: int, row_data: Row, join_str: str = " ") -> str:
        """
        Applies formatting to rows and columns.

        Args:
            val (Any): Any value
            idx_row (int): current row index being processed
            row_data: (Row): List of data to format

        Returns:
            List[str]: List of formatted values
        """
        self._idx_col_last = len(row_data) - 1  # maybe -1
        return self._format_row_items(idx_row=idx_row, row_data=row_data, join_str=join_str)

    # region Properties
    @property
    def format(self) -> str | Tuple[str, ...]:
        """
        Gets format

        Format option such as ``.2f``

        Multiple formats can be added such as ``(".2f", "<10")``.
        Formats are applied in the order they are added.
        In this case first float is formated as string with two decimal places, and
        then value is padded to the right with spaces.
        """
        return self._format

    @property
    def col_formats(self) -> List[FormatterTableItem]:
        """Gets col_formats value"""
        return self._col_formats

    @property
    def custom_formats_row(self) -> List[FormatterTableRow]:
        """Gets custom_formats_row value"""
        return self._custom_formats_row

    @property
    def custom_formats_row_items(self) -> List[FormatterTableItem]:
        """Gets custom_formats_row_items value"""
        return self._custom_formats_row_items

    @property
    def custom_formats_col(self) -> List[FormatterTableItem]:
        """Gets custom_formats_col value"""
        return self._custom_formats_col

    # endregion Properties
