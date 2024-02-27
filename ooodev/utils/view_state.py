# coding: utf-8
# Python conversion of ViewState.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https:#fivedots.coe.psu.ac.th/~ad/jlop/
from __future__ import annotations
from enum import IntEnum
from ooodev.office import calc as mCalc
from ooodev.loader import lo as mLo


class ViewState:
    """
    For moving the pane focus

    See Also:
        :ref:`ch23_view_states_top_pane`
    """

    class PaneEnum(IntEnum):
        MOVE_UP = 0
        MOVE_DOWN = 1
        MOVE_LEFT = 2
        MOVE_RIGHT = 3

    def __init__(self, state: str) -> None:
        """
        Constructor

        Args:
            state (str): State in format of '0/4998/0/1/0/218/2/0/0/4988/4998'

        Raises:
            ValueError: if state does not contains 10 '/' (11 parts)
        """
        # The state string has the format:
        #       0/4998/0/1/0/218/2/0/0/4988/4998

        self._cursor_column: int = 0
        self._cursor_row: int = 0
        self._col_split_mode: int = 0
        self._col_split_mode: int = 0
        self._row_split_mode: int = 0
        self._row_split_mode: int = 0
        self._vertical_split: int = 0
        self._horizontal_split: int = 0
        self._pane_focus_num: int = 0
        self._column_left_pane: int = 0
        self._column_right_pane: int = 0
        self._row_upper_pane: int = 0
        self._row_lower_pane: int = 0
        states = state.split("/")
        if len(states) != 11:
            raise ValueError(f"Incorrect number of states, Expected 11 got {len(states)}")

        self._cursor_column = ViewState.parse_int(states[0])  # 0: cursor position column
        self._cursor_row = ViewState.parse_int(states[1])  # 1: cursor position row

        self._col_split_mode = ViewState.parse_int(states[2])  # 2: column split mode
        self._row_split_mode = ViewState.parse_int(states[3])  # 3: row split mode

        self._vertical_split = ViewState.parse_int(states[4])  # 4: vertical split position
        self._horizontal_split = ViewState.parse_int(states[5])  # 5: horizontal split position

        self._pane_focus_num = ViewState.parse_int(states[6])  # 6: focused pane number

        self._column_left_pane = ViewState.parse_int(states[7])  # 7: left column index of left pane
        self._column_right_pane = ViewState.parse_int(states[8])  # 8: left column index of right pane

        self._row_upper_pane = ViewState.parse_int(states[9])  # 9: top row index of upper pane
        self._row_lower_pane = ViewState.parse_int(states[10])  # 10: top row index of lower pane

    @staticmethod
    def parse_int(s: str) -> int:
        """
        Parses int

        Args:
            s (str): string value that contains int

        Returns:
            int: string value converted to int on success; Otherwise, 0
        """
        if not s:
            return 0
        try:
            return int(s)
        except ValueError:
            mLo.Lo.print(f"'{s}' could not be parsed as an int; using 0")
        return 0

    @property
    def cursor_column(self) -> int:
        """Gets/Sets cursor position column"""
        return self._cursor_column

    @cursor_column.setter
    def cursor_column(self, value: int) -> None:
        if value < 0:
            raise ValueError("Column position must be positive")
        self._cursor_column = value

    @property
    def cursor_row(self) -> int:
        """Gets/Sets cursor position row"""
        return self._cursor_row

    @cursor_row.setter
    def cursor_row(self, value: int) -> None:
        if value < 0:
            raise ValueError("Row position must be positive")
        self._cursor_row = value

    @property
    def column_split_mode(self) -> int:
        """Gets/Sets column split mode"""
        return self._col_split_mode

    @column_split_mode.setter
    def column_split_mode(self, value: bool | int) -> None:
        self._col_split_mode = int(value)
        if self._col_split_mode == 0:  # no column splitting
            self._vertical_split = 0
            if self._pane_focus_num in (1, 3):
                self._pane_focus_num -= 1  # move focus to left

    @property
    def row_split_mode(self) -> int:
        """Gets/Sets row split mode"""
        return self._row_split_mode

    @row_split_mode.setter
    def row_split_mode(self, value: bool | int) -> None:
        self._row_split_mode = int(value)
        if self._row_split_mode == 0:  # no row splitting
            self._horizontal_split = 0
            if self._pane_focus_num in (2, 3):
                self._pane_focus_num -= 2  # move focus up

    @property
    def vertical_split(self) -> int:
        """Gets/Sets vertical split position"""
        return self._vertical_split

    @vertical_split.setter
    def vertical_split(self, value: int) -> None:
        if value < 0:
            raise ValueError("Position must be positive")
        self._vertical_split = value

    @property
    def horizontal_split(self) -> int:
        """Gets/Sets horizontal split position"""
        return self._horizontal_split

    @horizontal_split.setter
    def horizontal_split(self, value: int) -> None:
        if value < 0:
            raise ValueError("Position must be positive")
        self._horizontal_split = value

    @property
    def pane_focus_num(self) -> int:
        """Gets/Sets focused pane number"""
        return self._pane_focus_num

    @pane_focus_num.setter
    def pane_focus_num(self, value: int) -> None:
        if value < 0 or value > 3:
            raise ValueError("Focus number is out of range 0-3")

        if self._horizontal_split == 0 and value in {1, 3}:
            raise ValueError("No horizontal split, so focus number must be 0 or 2")

        if self._vertical_split == 0 and value in {2, 3}:
            raise ValueError("No vertical split, so focus number must be 0 or 1")
        self._pane_focus_num = value

    def move_pane_focus(self, dir: int | ViewState.PaneEnum) -> bool:
        """
        Moves pane focus

        Args:
            dir (int | PaneEnum): Direction to move

        Raises:
            ValueError: If dir is unknown

        Returns:
            bool: True if move is successful; Otherwise False

        Note:
            The 4 possible view panes are numbered like so
            ::

                0  |  1
                -------
                2  |  3

            If there's no horizontal split then the panes are numbered 0 and 2.
            If there's no vertical split then the panes are numbered 0 and 1.
        """
        try:
            d = ViewState.PaneEnum(dir)
        except Exception as e:
            raise ValueError("Unknown move direction") from e

        if d == ViewState.PaneEnum.MOVE_UP:
            if self._pane_focus_num == 3:
                self._pane_focus_num = 1
            elif self._pane_focus_num == 2:
                self._pane_focus_num = 0
            else:
                mLo.Lo.print("cannot move up")
                return False
        elif d == ViewState.PaneEnum.MOVE_DOWN:
            if self._pane_focus_num == 1:
                self._pane_focus_num = 3
            elif self._pane_focus_num == 0:
                self._pane_focus_num = 2
            else:
                mLo.Lo.print("cannot move down")
                return False
        elif d == ViewState.PaneEnum.MOVE_LEFT:
            if self._pane_focus_num == 1:
                self._pane_focus_num = 0
            elif self._pane_focus_num == 3:
                self._pane_focus_num = 2
            else:
                mLo.Lo.print("cannot move left")
                return False
        elif d == ViewState.PaneEnum.MOVE_RIGHT:
            if self._pane_focus_num == 0:
                self._pane_focus_num = 1
            elif self._pane_focus_num == 2:
                self._pane_focus_num = 3
            else:
                mLo.Lo.print("cannot move right")
                return False
        return True

    @property
    def column_left_pane(self) -> int:
        """Gets/Sets left column index of left pane"""
        return self._column_left_pane

    @column_left_pane.setter
    def column_left_pane(self, value: int) -> None:
        if value < 0:
            raise IndexError("value must be positive")
        self._column_left_pane = value

    @property
    def column_right_pane(self) -> int:
        """Gets/Sets left column index of right pane"""
        return self._column_right_pane

    @column_right_pane.setter
    def column_right_pane(self, value: int) -> None:
        if value < 0:
            raise IndexError("value must be positive")
        self._column_right_pane = value

    @property
    def row_upper_pane(self) -> int:
        """Gets/Sets top row index of upper pane"""
        return self._row_upper_pane

    @row_upper_pane.setter
    def row_upper_pane(self, value: int) -> None:
        if value < 0:
            raise IndexError("value must be positive")
        self._row_upper_pane = value

    @property
    def row_lower_pane(self) -> int:
        """Gets/Sets top row index of lower pane"""
        return self._row_lower_pane

    @row_lower_pane.setter
    def row_lower_pane(self, value: int) -> None:
        if value < 0:
            raise IndexError("value must be positive")
        self._row_lower_pane = value

    def report(self) -> None:
        """
        Prints a report to console
        """
        print("Sheet View State")
        print(
            f"  Cursor pos (column, row): ({self.cursor_column}, {self.cursor_row}) or '{mCalc.Calc.get_cell_str(col=self.cursor_column, row=self.cursor_row)}'"
        )
        if self.column_split_mode == 1 and self.row_split_mode == 1:
            print(f"  Sheet is split vertically and horizontally at {self.vertical_split} / {self.horizontal_split}")
        elif self.column_split_mode == 1:
            print(f"  Sheet is split vertically at {self.vertical_split}")
        elif self.row_split_mode == 1:
            print(f"  Sheet is split horizontally at {self.horizontal_split}")
        else:
            print("  Sheet is not split")

        print(f"  Number of focused pane: {self.pane_focus_num}")
        print(f"  Left column indices of left/right panes: {self.column_left_pane} / {self.column_right_pane}")
        print(f"  Top row indices of upper/lower panes: {self.row_upper_pane} / {self.row_lower_pane}")
        print()

    def to_string(self) -> str:
        """
        Gets string Representation of object.

        String representation can also be used to create a new instance of this class.

        same as ``str(instance)``
        """
        lst = [
            self.cursor_column,
            self.cursor_row,
            self.column_split_mode,
            self.row_split_mode,
            self.vertical_split,
            self.horizontal_split,
            self.pane_focus_num,
            self.column_left_pane,
            self.column_right_pane,
            self.row_upper_pane,
            self.row_lower_pane,
        ]
        return "/".join([str(val) for val in lst])

    def __str__(self) -> str:
        return self.to_string()
