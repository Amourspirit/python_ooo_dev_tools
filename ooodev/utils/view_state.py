# coding: utf-8
# Python conversion of ViewState.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https:#fivedots.coe.psu.ac.th/~ad/jlop/
from __future__ import annotations
from enum import IntEnum
from ..office import calc as mCalc

class ViewState:
    # for moving he pane focus
    class PaneEnum(IntEnum):
        MOVE_UP = 0
        MOVE_DOWN = 1
        MOVE_LEFT = 2
        MOVE_RIGHT = 3
    
    
    
    def __init__(self, state: str) -> None:
        # The state string has the format:
        #       0/4998/0/1/0/218/2/0/0/4988/4998
        
        self._cursor_column:int = 0
        self._cursor_row: int = 0
        self._col_split_mode: int = 0
        self._col_split_mode: int = 0
        self._row_split_mode: int = 0
        self._row_split_mode: int = 0
        self._vertical_split: int = 0
        self._horizontal_split: int = 0
        self._focus_num: int = 0
        self._column_left_pane: int = 0
        self._column_right_pane: int = 0
        self._row_upper_pane: int = 0
        self._row_lower_pane: int = 0
        states = state.split('/')
        if len(states) != 11:
            print("Incorrect number of states")
            return
        self._cursor_column = ViewState.parse_int(states[0])        # 0: cursor position column
        self._cursor_row = ViewState.parse_int(states[1])           # 1: cursor position row

        self._col_split_mode = ViewState.parse_int(states[2])       # 2: column split mode
        self._row_split_mode = ViewState.parse_int(states[3])       # 3: row split mode

        self._vertical_split = ViewState.parse_int(states[4])       # 4: vertical split position
        self._horizontal_split = ViewState.parse_int(states[5])     # 5: horizontal split position

        self._focus_num = ViewState.parse_int(states[6])            # 6: focused pane number

        self._column_left_pane = ViewState.parse_int(states[7])     # 7: left column index of left pane
        self._column_right_pane = ViewState.parse_int(states[8])    # 8: left column index of right pane

        self._row_upper_pane = ViewState.parse_int(states[9])       # 9: top row index of upper pane
        self._row_lower_pane = ViewState.parse_int(states[10])      # 10: top row index of lower pane
            
    @staticmethod
    def parse_int(s: str) -> int:
        if not s:
            return 0
        try:
            return int(s)
        except ValueError:
            print(f"'{s}' could not be parsed as an int; using 0")
        return 0
    
    def get_cursor_column(self) -> int:
        return self._cursor_column
    
    def set_cursor_column(self, col_pos: int) -> None:
        if col_pos < 0:
            print("Column position must be positive")
            return
        self._cursor_column = col_pos
    
    def get_cursor_row(self) -> int:
        return self._cursor_row
    
    def set_cursor_row(self, row_pos: int) -> None:
        if row_pos < 0:
            print("Row position must be positive")
            return
        self._cursor_row = row_pos
        
    def get_column_split_mode(self) -> int:
        return self._col_split_mode
    
    def set_column_split_mode(self, is_split: bool) -> None:
        self._col_split_mode = 1 if is_split else 0
        if self._col_split_mode == 0: # no column splitting
            self._vertical_split = 0
            if self._focus_num ==1 or self._focus_num == 3:
                self._focus_num -= 1 # move focus to left
    
    def get_row_split_mode(self) -> int:
        return self._row_split_mode
    
    def set_row_split_mode(self, is_split: bool) -> None:
        self._row_split_mode = 1 if is_split else 0
        if self._row_split_mode == 0: # no row splitting
            self._horizontal_split = 0
            if self._focus_num == 2 or self._focus_num == 3:
                self._focus_num -= 2 # move focus up
    
    def get_vertical_split(self) -> int:
        return self._vertical_split
    
    def set_vertical_split(self, split_pos: int) -> None:
        if split_pos < 0:
            print("Position must be positive")
            return
        self._vertical_split = split_pos
    
    def get_horizontal_split(self) -> int:
        return self._horizontal_split
    
    def set_horizontal_split(self, split_pos: int) -> None:
        if split_pos < 0:
            print("Position must be positive")
            return
        self._horizontal_split = split_pos
    
    def get_pane_focus_num(self) -> int:
        return self._focus_num
    
    def set_pane_focus_num(self, n: int) -> None:
        if n < 0 or n > 3:
            print("Focus number is out of range 0-3")
            return
        
        if self._horizontal_split == 0 and (n == 1 or n == 3):
            print("No horizontal split, so focus number must be 0 or 2")
            return
        
        if self._vertical_split == 0 and (n == 2 or n ==3):
            print("No horizontal split, so focus number must be 0 or 1")
            return
        self._focus_num = n
    
    def move_pane_focus(self, dir:int | PaneEnum) -> None:
        """
        The 4 posible view panes are numbered like so
        ::
        
            0  |  1
            -------
            2  |  3
            
        If there's no horizontal split then the panes are numbered 0 and 2.
        If there's no vertical split then the panes are numbered 0 and 1.
        """
        try:
            d = ViewState.PaneEnum(dir)
        except Exception:
            print("Unknown move direction")
            return
        
        if d == ViewState.PaneEnum.MOVE_UP:
            if self._focus_num == 3:
                self._focus_num = 1
            elif self._focus_num == 2:
                self._focus_num = 0
            else:
                print("cannot move up")
        elif d == ViewState.PaneEnum.MOVE_DOWN:
            if self._focus_num == 1:
                self._focus_num = 3
            elif self._focus_num == 0:
                self._focus_num = 2
            else:
                print("cannot move down")
        elif d == ViewState.PaneEnum.MOVE_LEFT:
            if self._focus_num == 1:
                self._focus_num = 0
            elif self._focus_num == 3:
                self._focus_num = 2
            else:
                print("cannot move left")
        elif d == ViewState.PaneEnum.MOVE_RIGHT:
            if self._focus_num == 0:
                self._focus_num = 1
            elif self._focus_num == 2:
                self._focus_num = 3
            else:
                print("cannot move right")
    
    def get_column_left_pane(self) -> int:
        return self._column_left_pane
    
    def set_column_left_pane(self, idx: int) -> None:
        if idx < 0:
            print("Index must be positive")
            return
        self._column_left_pane = idx
    
    def get_column_right_pane(self, idx:int) -> None:
        if idx < 0:
            print("Index must be positive")
            return
        self._column_right_pane = idx
    
    def get_row_upper_pane(self) -> int:
        return self._column_right_pane
    
    def set_row_upper_pane(self, idx: int) -> None:
        if idx < 0:
            print("Index must be positive")
            return
        self._row_upper_pane = idx
    
    def get_row_lower_pane(self) -> int:
        return self._row_lower_pane
    
    def set_row_lower_pane(self, idx:int) -> None:
        if idx < 0:
            print("Index must be positive")
            return
        self._row_lower_pane = idx
    
    def report(self) -> None:
        print("Sheet View State")
        print(f"  Cursor pos (column, row): ({self._cursor_column}, {self._cursor_row}) or '{mCalc.Calc.get_cell_str(col=self._cursor_column, row=self._cursor_row)}'")
        if self._col_split_mode == 1 and self._row_split_mode == 1:
            print(f"  Sheet is split vertically and horizontally at {self._vertical_split} / {self.get_horizontal_split}")
        elif self._col_split_mode == 1:
            print(f"  Sheet is split vertically at {self._vertical_split}")
        elif self._row_split_mode ==1:
            print(f"  Sheet is split horizontally at {self._horizontal_split}")
        else:
            print("  Sheet is not split")
        
        print(f"  Number of focused pane: {self._focus_num}")
        print(f"  Left column indicies of left/right panes: {self._column_left_pane} / {self._column_right_pane}")
        print(f"  Top row indicies of upper/lower panes: {self._row_upper_pane} / {self._row_lower_pane}")
        print()
    
    def to_string(self) -> str:
        lst = [
            self._cursor_column,
            self._cursor_row,
            self._vertical_split,
            self._horizontal_split,
            self._focus_num,
            self._column_left_pane,
            self._column_right_pane,
            self._row_upper_pane,
            self._row_lower_pane
        ]
        return '/'.join([str(val) for val in lst])

    def __str__(self) -> str:
        return self.to_string()
    
    