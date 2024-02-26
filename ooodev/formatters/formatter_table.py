from __future__ import annotations
from typing import List, Set, Any, Tuple, TYPE_CHECKING
from ooodev.formatters.table_rule_kind import TableRuleKind as TableRuleKind
from ooodev.formatters.format_table_item import FormatTableItem as FormatTableItem
from ooodev.formatters.format_string import FormatString as FormatString
from ooodev.formatters.table_item_processor import TableItemProcessor

if TYPE_CHECKING:
    from ..utils.type_var import Row


class FormatterTable:
    """
    Formatter for formatting 2d sequences.

    See Also:
        :py:meth:`.Calc.print_array`
    """

    def __init__(
        self,
        format: str | Tuple[str, ...],
        idx_rule: TableRuleKind = TableRuleKind.IGNORE,
        idxs: Tuple[int, ...] | None = None,
        **kwargs,
    ) -> None:
        """
        Constructor

        Args:
            format (str | Tuple[str, ...]): Formatting that is applied to  specified data.
            idx_rule (FormatterTableRuleKind, optional): Flag Options that determine behaviors. Defaults to FormatterTableRuleKind.IGNORE.
            idxs (Tuple[int, ...] | None, optional): Row Indexes that specify which rows are affect. Defaults to None.
                If ``idx_rule`` contains ``FormatterTableRuleKind.IGNORE`` then ``idxs`` are the rows that are excluded from formatting.
                If ``idx_rule`` contains ``FormatterTableRuleKind.ONLY`` then ``idxs`` are the rows that are formatted.
            **Kwargs: Expandable list or Key, value args that are auto assigned as class Attributes. Useful for child classes.
        """
        self._format = format
        self._idx_rule = idx_rule
        self._idxs = () if idxs is None else idxs
        self._row_formats: List[FormatTableItem] = []
        self._custom_formats_str: List[FormatString] = []
        # self._custom_formats_col: List[FormatterTableItem] = []
        self._col_formats: List[FormatTableItem] = []
        self._cols = None
        self._idx_col_last = -1
        # for classes that inherit FormatterTable and need to expand constructor
        for key, value in kwargs.items():
            setattr(self, key, value)

    def _get_cols(self) -> Tuple[int, ...]:
        if self._cols is None:
            cols: Set[int] = set()
            for cf in self.col_formats:
                cols.update(cf.idxs_inc)
            self._cols = tuple(cols)
        return self._cols

    def _has_col(self, idx_col: int) -> bool:
        cols = self._get_cols()
        return idx_col in cols

    def _custom_row_item_format(self, idx_row: int) -> FormatTableItem | None:
        if idx_row < 0:
            return None
        return next((cf for cf in self._row_formats if cf.is_index(idx_row)), None)

    def _custom_row_format(self, idx_row: int) -> FormatString | None:
        if idx_row < 0:
            return None
        return next((cf for cf in self._custom_formats_str if cf.has_index(idx_row)), None)

    def _get_col_formatter(self, idx_col: int) -> FormatTableItem | None:
        if idx_col < 0:
            return None
        return next((cf for cf in self._col_formats if cf.is_index(idx_col)), None)

    def _format_col(self, idx_row: int, idx_col: int, val: Any) -> Tuple[bool, Any]:
        # returns original value or formatting string.
        cf = self._get_col_formatter(idx_col=idx_col)

        do_col_format = True
        if not self._is_format_row(idx_row=idx_row):
            # this row may be ignored, check for col override:
            do_col_format = TableRuleKind.COL_OVER_ROW in self._idx_rule

        if not do_col_format and cf and TableRuleKind.COL_OVER_ROW in self._idx_rule:
            return TableItemProcessor.process_col(
                itm=cf, idx_row=idx_row, idx_col=idx_col, idx_col_last=self._idx_col_last, value=val
            )

        if cf:
            # there is a custom column format and column expects formatting.
            # ignore column formatting and apply only custom formatting.
            return TableItemProcessor.process_col(
                itm=cf, idx_row=idx_row, idx_col=idx_col, idx_col_last=self._idx_col_last, value=val
            )

        # no column formatting
        return (False, val)

    def _is_format_row(self, idx_row: int) -> bool:
        if TableRuleKind.IGNORE in self._idx_rule:
            return idx_row not in self._idxs
        return idx_row in self._idxs if TableRuleKind.ONLY in self._idx_rule else False

    def _format_row(self, idx_row, row_data: List[str], join_str: str) -> str:
        if len(self._custom_formats_str) == 0:
            return join_str.join(row_data).rstrip()

        if custom_fmt_row := self._custom_row_format(idx_row=idx_row):
            row_str = join_str.join(row_data)
            return custom_fmt_row.get_formatted(row_str).rstrip()

        return join_str.join(row_data).rstrip()

    def _format_row_items(self, idx_row: int, row_data: Row, join_str: str) -> str:
        is_row_format = self._is_format_row(idx_row=idx_row)
        fmt_rows_idxs = set()
        # do not apply column formatting to rows that are ignored.
        if is_row_format:
            fmt_rows = []
            for i, val in enumerate(row_data):
                col_formatted, col_val = self._format_col(idx_row=idx_row, idx_col=i, val=val)
                fmt_rows.append(col_val)
                if col_formatted:
                    fmt_rows_idxs.add(i)
        else:
            fmt_rows = row_data

        if custom_row_fmt_itm := self._custom_row_item_format(idx_row=idx_row):
            # there is a custom format for this row.
            # ignore other format for row and apply custom.
            #
            # perhaps later other flag such as CUSTOM_COL_OVER_CUSTOM_ROW should be added.
            # if col over rows then rows should ignore cols that have been formatted
            # for now keep it a simple as possible.
            s_row = []
            for i, val in enumerate(fmt_rows):
                # if self._has_col(idx_col=i):
                #     s_row.append(str(val))
                # else:
                p_row_state, p_row_val = TableItemProcessor.process_row(
                    itm=custom_row_fmt_itm, idx_row=idx_row, idx_col=i, idx_col_last=self._idx_col_last, value=val
                )
                if p_row_state:
                    s_row.append(p_row_val)
                else:
                    s_row.append(str(p_row_val))
            return self._format_row(idx_row=idx_row, row_data=s_row, join_str=join_str)

        if is_row_format:
            # if col over rows then rows should ignore cols that have been formatted
            s_row = []
            for i, val in enumerate(fmt_rows):
                if i in fmt_rows_idxs:
                    # if self._has_col(idx_col=i):
                    s_row.append(val)  # will be string because it is formatted
                else:
                    s_row.append(self._apply_all_formats(val, self._format))
        else:
            s_row = [str(v) for v in fmt_rows]

        return self._format_row(idx_row=idx_row, row_data=s_row, join_str=join_str)

    def _apply_single_format(self, value, fmt: str) -> str:
        try:
            return format(value, fmt)
        except Exception:
            return str(value)

    def _apply_all_formats(self, value, formats: str | Tuple[str, ...]) -> str:
        if not isinstance(formats, tuple):
            return self._apply_single_format(value, formats)
        v = value
        for fmt in formats:
            v = self._apply_single_format(v, fmt)
        return str(v)

    def get_formatted(self, idx_row: int, row_data: Row, join_str: str = " ") -> str:
        """
        Applies formatting to rows and columns.

        Args:
            val (Any): Any value
            idx_row (int): current row index being processed
            row_data: (Row): List of data to format

        Returns:
            str: formatted row
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
        In this case first float is formatted as string with two decimal places, and
        then value is padded to the right with spaces.
        """
        return self._format

    @property
    def col_formats(self) -> List[FormatTableItem]:
        """
        Gets Column formats

        New formats can be added.

        Applies to any column that matches.

        Example:
            .. code-block:: python

                fl.col_formats.append(FormatTableItem(format=(".0%", ">9"), idxs=(4,)))
        """
        return self._col_formats

    @property
    def custom_formats_str(self) -> List[FormatString]:
        """
        Gets Custom formats string value.

        If you need to apply extra formatting to a row after is has been formatted as string
        then ``FormatterString`` instances can be added.
        """
        return self._custom_formats_str

    @property
    def row_formats(self) -> List[FormatTableItem]:
        """
        Gets Row formats

        New formats can be added.

        Applies to any column that matches.

        Example:
            .. code-block:: python

                fl.row_formats.append(FormatTableItem(format=">9", idxs_inc=(0, 9)))
        """
        return self._row_formats

    # endregion Properties


__all__ = ["FormatterTable"]
