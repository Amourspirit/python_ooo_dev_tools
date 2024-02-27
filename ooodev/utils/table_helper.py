# coding: utf-8
"""General Utilities"""

from __future__ import annotations
import contextlib
import re
import sys
import string
from typing import Callable, Iterable, Sequence, List, Any, Tuple, overload, TypeVar, NamedTuple, TYPE_CHECKING

from ooodev.utils import gen_util as gUtil

if TYPE_CHECKING:
    from ooodev.utils.type_var import DictTable
    from ooodev.utils.type_var import Table
else:
    DictTable = Any
    Table = Any

T = TypeVar("T")


class CellParts(NamedTuple):
    """Cell Named parts"""

    sheet: str
    col: str
    row: int

    def __str__(self) -> str:
        if self.sheet:
            return f"{self.sheet}.{self.col}{self.row}"
        else:
            return f"{self.col}{self.row}"


class RangeParts(NamedTuple):
    """Range Named parts"""

    sheet: str
    col_start: str
    row_start: int
    col_end: str
    row_end: int

    def __str__(self) -> str:
        if self.sheet:
            return f"{self.sheet}.{self.col_start}{self.row_start}:{self.col_end}{self.row_end}"
        else:
            return f"{self.col_start}{self.row_start}:{self.col_end}{self.row_end}"


class TableHelper:
    @classmethod
    def get_cell_parts(cls, cell_name: str) -> CellParts:
        """
        Gets cell parts from a cell name.

        Args:
            cell_name (str): Cell name such as ``A23`` or ``Sheet1.A23``

        Returns:
            CellParts: Cell Parts

        Note:
            If a range name such as ``A23:G45`` or ``Sheet1.A23:G45`` then only the first cell is used.

            Column name is upper case.

        .. versionadded:: 0.8.3
        """
        doc_idx = cell_name.find(".")

        if doc_idx >= 0:
            sheet_name = cell_name[:doc_idx]
            cell_name = cell_name[doc_idx + 1 :]
        else:
            sheet_name = ""
        # split will cover if a range is passed in, return first cell
        cells = cell_name.split(":")

        col = cells[0].rstrip(string.digits).upper()
        row = cls.row_name_to_int(cells[0])

        return CellParts(sheet=sheet_name, col=col, row=row)

    @classmethod
    def get_range_parts(cls, range_name: str) -> RangeParts:
        """
        Gets range parts from a range name.

        Args:
            range_name (str): Range name such as ``A23:G45`` or ``Sheet1.A23:G45``

        Returns:
            RangeParts: Range Parts

        Notes:
            Column names are upper case.

        .. versionadded:: 0.8.2
        """
        doc_idx = range_name.find(".")

        if doc_idx >= 0:
            sheet_name = range_name[:doc_idx]
            range_name = range_name[doc_idx + 1 :]
        else:
            sheet_name = ""

        cells = range_name.split(":")
        col_start = cells[0].rstrip(string.digits).upper()
        col_end = cells[1].rstrip(string.digits).upper()
        row_start = cls.row_name_to_int(cells[0])
        row_end = cls.row_name_to_int(cells[1])
        start_col_num = cls.col_name_to_int(col_start)
        end_col_num = cls.col_name_to_int(col_end)
        # check and see if the range name need to be reversed
        # e.g. D12:B3 => B3:D12
        if start_col_num > end_col_num:
            # swap
            col_start, col_end = col_end, col_start
        # B12:D3 => B3:D12
        if row_start > row_end:
            row_start, row_end = row_end, row_start

        return RangeParts(sheet=sheet_name, col_start=col_start, row_start=row_start, col_end=col_end, row_end=row_end)

    @staticmethod
    def col_name_to_int(name: str, zero_index: bool = False) -> int:
        """
        Converts a Column Name into an int.

        Args:
            name (str):Case insensitive column name such as 'a' or 'AB'
            zero_index (bool, optional): determines if return is zero based or one based. Default ``False``.

        Raises:
            ValueError: If number of Column letters (``name``) is ``0`` or greater than ``5``.
            ValueError: If ``name`` is not valid characters, ``[a-zA-Z0-9]``.

        Returns:
            int: One based int representing column name

        Example:
            .. code-block:: python

                >>> TableHelper.col_name_to_int('a')
                1
                >>> TableHelper.col_name_to_int('a', True)
                0

        .. versionchanged:: 0.8.2
            Added ``zero_index`` parameter.
        """
        chars = name.rstrip(string.digits).strip()
        c_len = len(chars)
        if c_len == 0:
            raise ValueError("Empty name value or Invalid or characters detected.")
        if c_len > 5:
            raise ValueError(
                f"Maximum number of letters that can be processed is 5. currently name contains {len(chars)}"
            )
        if re.fullmatch(r"[a-zA-Z]+", chars) is None:
            raise ValueError(f'name of "{name}" contains invalid characters. Must be in format of "A" or "AB" or "A3"')

        power = 1
        col_num = 0
        for letter in chars[::-1]:  # reverse chars
            col_num += (int(letter, 36) - 9) * power
            power *= 26
        return col_num - 1 if zero_index else col_num

    @staticmethod
    def row_name_to_int(name: str, zero_index: bool = False) -> int:
        """
        Converts a row name into an int.
        Leading Alpha chars are ignore. String ``4`` converts to integer ``4``.
        String ``C5`` converts to integer ``5``.

        If ``zero_index`` is ``True`` string ``4`` converts to integer ``3``.
        String ``C5`` converts to integer ``4``.

        Args:
            name (str): row name to convert
            zero_index (bool, optional): determines if return is zero based or one based. Default ``False``.

        Returns:
            int: converted name as int.

        Example:
            .. code-block:: python

                >>> TableHelper.row_name_to_int('C5')
                5
                >>> TableHelper.row_name_to_int('C5', True)
                4

        .. versionchanged:: 0.8.2
            Added ``zero_index`` parameter.
        """
        # sourcery skip: assign-if-exp, reintroduce-else, use-named-expression
        chars = name.rstrip(f"{string.digits}-")
        if chars:
            s = name[len(chars) :]  # drop leading chars that are not numbers.
        else:
            s = name
        result = int(s)
        if result < 0:
            raise ValueError(f"Cannot parse negative values: {name}")
        if zero_index:
            return result - 1
        return result

    @classmethod
    def make_cell_name(cls, row: int, col: int, zero_index: bool = False) -> str:
        """
        Convert given row and column number to ``A1`` style cell name.

        Args:
            row (int): Row number.
            col (int): Column Number.
            zero_index (bool, optional): determines if return is zero based or one based. Default ``False``.

        Raises:
            ValueError: If col Value is ``<1`` for one based or ``<0`` for zero based.

        Returns:
            str: row and col as cell name such as A1, AB3

        Example:
            .. code-block:: python

                >>> TableHelper.make_cell_name(1, 1)
                A1
                >>> TableHelper.make_cell_name(0, 0, True)
                A1

        .. versionchanged:: 0.8.2
            Added ``zero_index`` parameter.
        """
        idx_min = 0 if zero_index else 1
        if row < idx_min:
            raise ValueError(f"Row value cannot be less then {idx_min}: {row}")
        row_index = row + 1 if zero_index else row
        return f"{cls.make_column_name(col, zero_index)}{row_index}"

    @staticmethod
    def make_column_name(col: int, zero_index: bool = False) -> str:  # col is 1 based
        """
        Makes a cell style name. eg: A, B, C, ... AA, AB, AC

        Args:
            col (int): Column number.
            zero_index (bool, optional): determines if return is zero based or one based. Default ``False``.

        Raises:
            ValueError: If col Value is ``<1`` for one based or ``<0`` for zero based.

        Returns:
            str: column name. eg: A, B, C, ... AA, AB, AC

        .. versionchanged:: 0.8.2
            Added ``zero_index`` parameter.
        """
        idx_min = 0 if zero_index else 1
        if col < idx_min:
            raise ValueError(f"Value cannot be less then {idx_min}: {col}")
        str_col = str()
        div = col + 1 if zero_index else col
        while div:
            (div, mod) = divmod(div - 1, 26)  # will return (x, 0 .. 25)
            str_col = chr(mod + 65) + str_col
        return str_col

    @overload
    @staticmethod
    def make_2d_array(num_rows: int, num_cols: int) -> List[List[Any]]:
        """
        Make a 2-Dimensional List of values with each element having a value of ``1``

        Args:
            num_rows (int): Number of rows
            num_cols (int): Number of Columns in each row.

        Returns:
            List[List[Any]]: 2-Dimensional List of values
        """
        ...

    @overload
    @staticmethod
    def make_2d_array(num_rows: int, num_cols: int, val: Any) -> List[List[Any]]:
        """
        Make a 2-Dimensional List of values

        Args:
            num_rows (int): Number of rows
            num_cols (int): Number of Columns in each row.
            val (Any): Value of each element in the list.

        Returns:
            List[List[Any]]: 2-Dimensional List of values
        """
        ...

    @overload
    @staticmethod
    def make_2d_array(num_rows: int, num_cols: int, val: Callable[[int, int, Any], Any]) -> List[List[Any]]:
        """
        Make a 2-Dimensional List of values

        Args:
            num_rows (int): Number of rows
            num_cols (int): Number of Columns in each row.
            val (Callable[[int, int, Any], Any]): Callable that provide each value.
                Callback e.g. cb(row: int, col: int, prev_value: None | int) -> int:...

        Returns:
            List[List[Any]]: 2-Dimensional List of values
        """
        ...

    @staticmethod
    def make_2d_array(num_rows: int, num_cols: int, val=None) -> List[List[Any]]:
        """
        Make a 2-Dimensional List of values

        Args:
            num_rows (int): Number of rows
            num_cols (int): Number of Columns in each row.
            val (Callable[[int, int, Any], Any]): Callable that provide each value.
                Callback e.g. ``cb(row: int, col: int, prev_value: None | int) -> int:...``

        Returns:
            List[List[Any]]: 2-Dimensional List of values

        .. collapse:: Example

            Example of array filled with 1's

            .. code-block:: python

                arr = TableHelper.make_2d_array(num_rows=3, num_cols=4, val=1)
                # arr
                # [
                #   [1, 1, 1, 1],
                #   [1, 1, 1, 1],
                #   [1, 1, 1, 1]
                # ]


            Example of using call back method.

            The following example creates an array that loops through each animals and adds
            to array. When end of animals is reached the start with the beginning of animals and
            continues in this fashion until array is built.

            .. code-block:: python

                animals = ("ass", "cat", "cow", "cub", "doe", "dog", "elk",
                            "ewe", "fox", "gnu", "hog", "kid", "kit", "man",
                            "orc", "pig", "pup", "ram", "rat", "roe", "sow", "yak")
                total_rows = 15
                total_cols = 6

                def cb(row:int, col:int, prev) -> str:
                    # return animals repeating until all cells are filled
                    v = (row * total_cols) + col

                    a_len = len(animals)
                    if v > a_len - 1:
                        i = (v % a_len)
                    else:
                        i = v
                    return animals[i]

                arr = TableHelper.make_2d_array(num_rows=total_rows, num_cols=total_cols, val=cb)
                Calc.set_array(values=arr, sheet=sheet, name="A1")
        """
        if num_cols == 0 or num_rows == 0:
            return []
        if val is None:
            val = 1
        if callable(val):
            data = []
            new_val = None
            for row in range(num_rows):
                col_data = []
                for col in range(num_cols):
                    new_val = val(row, col, new_val)
                    col_data.append(new_val)
                data.append(col_data)
        else:
            data = [[val] * num_cols for _ in range(num_rows)]
        return data

    make_2d_list = make_2d_array

    @staticmethod
    def to_list(iter_obj: Sequence[Any] | Any) -> List[Any]:
        """
        Converts an iterable of objects into a list of objects

        If ``iter_obj`` is not iterable it will be return as a tuple containing ``iter_obj``

        Args:
            iter_obj (Iterable[Any] | object): iterable object or object.

        Returns:
            List[Any]: List containing same elements of ``iter_obj``
        """
        return list(iter_obj) if gUtil.Util.is_iterable(iter_obj) else [iter_obj]

    @staticmethod
    def to_tuple(iter_obj: Iterable[Any] | Any) -> Tuple[Any]:
        """
        Converts an iterable of objects or object into a tuple of objects

        If ``iter_obj`` is not iterable it will be return as a tuple containing ``iter_obj``

        Args:
            iter_obj (Iterable[Any] | object): iterable object or object.

        Returns:
            Tuple[Any]: Tuple containing same elements of ``iter_obj``
        """
        return tuple(iter_obj) if gUtil.Util.is_iterable(iter_obj) else (iter_obj,)  # type: ignore

    @classmethod
    def to_2d_list(cls, seq_obj: Sequence[Any]) -> List[List[Any]]:
        """
        Converts a sequence of sequence to a list.

        Converts 1-Dimensional or 2-Dimensional array such as a Tuple or a Tuple of Tuple's into a List of List.

        An array of tuples is immutable and can not add or remove elements whereas a list is mutable.

        Args:
            seq_obj (Sequence[Any]): 1-Dimensional or 2-Dimensional List

        Returns:
            List[List[Any]]: 2-Dimensional list
        """
        num_rows = len(seq_obj)
        if num_rows == 0:
            return []
        is_2d = False
        try:
            is_2d = gUtil.Util.is_iterable(seq_obj[0])
        except Exception:
            is_2d = False
        lst = []
        if is_2d:
            lst.extend(cls.to_list(row) for row in seq_obj)
        else:
            lst.append(cls.to_list(seq_obj))
        return lst

    @classmethod
    def to_2d_tuple(cls, seq_obj: Sequence[Any]) -> Tuple[Tuple[Any, ...], ...]:
        """
        Converts a sequence of sequence to a tuple of tuple.

        Converts 1-Dimensional or 2-Dimensional array such as a List or List of list's into a Tuple of Tuple.

        Args:
            seq_obj (Sequence[Any]): 1-Dimensional or 2-Dimensional Sequence
        Returns:
            Tuple[Tuple[Any, ...], ...]: 2-Dimensional tuple
        """
        num_rows = len(seq_obj)
        if num_rows == 0:
            return ()
        is_2d = False
        try:
            is_2d = gUtil.Util.is_iterable(seq_obj[0])
        except Exception:
            is_2d = False
        lst = []
        if is_2d:
            lst.extend(cls.to_tuple(row) for row in seq_obj)
        else:
            lst.append(cls.to_tuple(seq_obj))
        return tuple(lst)

    @staticmethod
    def table_2d_to_dict(tbl: Table) -> DictTable:
        """
        Converts a 2-D table into a Dictionary Table

        Args:
            tbl (Table): 2-D table

        Raises:
            ValueError: If tbl does not contain at least two rows

        Returns:
            DictTable: As List of Dict with each dict key representing column name

        See Also:
            :py:meth:`~.TableHelper.table_dict_to_table`
        """
        if len(tbl) < 2:
            raise ValueError("Cannot convert Table with less than two rows")
        # first row is column headers
        try:
            cols = list(tbl[0])
            return [dict(zip(cols, row)) for i, row in enumerate(tbl) if i != 0]
        except Exception as e:
            raise e

    @staticmethod
    def table_dict_to_table(tbl: DictTable) -> Table:
        """
        Converts a Dictionary Table to a 2-D table

        Args:
            tbl (DictTable): Dictionary Table

        Raises:
            ValueError: If tbl does not contain any rows

        Returns:
            Table: As 2-D table

        See Also:
            :py:meth:`~.TableHelper.table_2d_to_dict`
        """
        if len(tbl) == 0:
            raise ValueError("Cannot convert table with no rows")
        try:
            first = tbl[0]
            cols = list(first.keys())
            data = [cols]
            data.extend([v for _, v in row.items()] for row in tbl)
            return data
        except Exception as e:
            raise e

    # region Get Smallest or Largest int in a 1d or 2d sequence

    @staticmethod
    def _get_extreme_element_value_1d_int(seq_obj: Sequence[int], biggest: bool) -> int:
        # max_size = sys.maxsize
        # min_size = -sys.maxsize - 1
        min_max = -sys.maxsize - 1 if biggest else sys.maxsize
        result = min_max
        count = 0
        for i in seq_obj:
            with contextlib.suppress(Exception):
                if biggest and i > result or not biggest and i < result:
                    result = i
                # count valid numbers
                count += 1
        if count == 0:
            raise ValueError("sequence did not have any integers test.")
        return result

    @staticmethod
    def _get_extreme_element_value_2d_int(seq_obj: Sequence[Sequence[int]], biggest: bool) -> int:
        # max_size = sys.maxsize
        # min_size = -sys.maxsize - 1
        # for float max: float("inf"), min: float("-inf")
        min_max = -sys.maxsize - 1 if biggest else sys.maxsize
        result = min_max
        count = 0
        for row in seq_obj:
            for col in row:
                with contextlib.suppress(Exception):
                    if biggest and col > result or not biggest and col < result:
                        result = col
                    # count valid numbers
                    count += 1
        if count == 0:
            raise ValueError("sequence did not have any integers test.")
        return result

    @classmethod
    def _get_extreme_element_value_int(cls, seq_obj: Sequence[int] | Sequence[Sequence[int]], biggest: bool) -> int:
        num_rows = len(seq_obj)

        is_2d = False
        if num_rows > 0:
            try:
                is_2d = gUtil.Util.is_iterable(seq_obj[0])
            except Exception:
                is_2d = False

        if is_2d:
            return cls._get_extreme_element_value_2d_int(seq_obj, biggest)  # type: ignore

        return cls._get_extreme_element_value_1d_int(seq_obj, biggest)  # type: ignore

    @classmethod
    def get_largest_int(cls, seq_obj: Sequence[int] | Sequence[Sequence[int]]) -> int:
        """
        Gets the largest int in a ``1d`` or ``2d`` Sequence integers

        Args:
            seq_obj (Sequence[int] | Sequence[Sequence[int]]): Input sequence

        Raises:
            ValueError: If no integers in are found.

        Returns:
            int: Largest integer found in sequence

        .. versionadded:: 0.6.7
        """
        return cls._get_extreme_element_value_int(seq_obj=seq_obj, biggest=True)

    @classmethod
    def get_smallest_int(cls, seq_obj: Sequence[int] | Sequence[Sequence[int]]) -> int:
        """
        Gets the smallest int in a ``1d`` or ``2d`` Sequence integers

        Args:
            seq_obj (Sequence[int] | Sequence[Sequence[int]]): Input sequence

        Raises:
            ValueError: If no integers in are found.

        Returns:
            int: Smallest integer found in sequence

        .. versionadded:: 0.6.7
        """
        return cls._get_extreme_element_value_int(seq_obj=seq_obj, biggest=False)

    # endregion Get Smallest or Largest int in a 1d or 2d sequence

    # region Get Smallest or Largest string in a 1d or 2d sequence
    @staticmethod
    def _get_extreme_element_value_1d_str(seq_obj: Sequence[str], biggest: bool) -> int:
        if not seq_obj:
            return -1
        # max_size = sys.maxsize
        # min_size = -sys.maxsize - 1
        result_len = -1 if biggest else sys.maxsize
        count = 0
        for s in seq_obj:
            with contextlib.suppress(Exception):
                s_len = len(s)
                if biggest and s_len > result_len or not biggest and s_len < result_len:
                    result_len = s_len
                count += 1
        return -1 if count == 0 else result_len

    @staticmethod
    def _get_extreme_element_value_2d_str(seq_obj: Sequence[Sequence[str]], biggest: bool) -> int:
        if not seq_obj:
            return -1
        # max_size = sys.maxsize
        # min_size = -sys.maxsize - 1
        result_len = -1 if biggest else sys.maxsize
        count = 0
        for row in seq_obj:
            for s in row:
                with contextlib.suppress(Exception):
                    s_len = len(s)
                    if biggest and s_len > result_len or not biggest and s_len < result_len:
                        result_len = s_len
                    count += 1
        return -1 if count == 0 else result_len

    @classmethod
    def _get_extreme_element_value_str(cls, seq_obj: Sequence[str] | Sequence[Sequence[str]], biggest: bool) -> int:
        num_rows = len(seq_obj)

        is_2d = False
        if num_rows > 0:
            try:
                is_2d = gUtil.Util.is_iterable(seq_obj[0])
            except Exception:
                is_2d = False

        if is_2d:
            return cls._get_extreme_element_value_2d_str(seq_obj, biggest)

        return cls._get_extreme_element_value_1d_str(seq_obj, biggest)  # type: ignore

    @classmethod
    def get_largest_str(cls, seq_obj: Sequence[str] | Sequence[Sequence[str]]) -> int:
        """
        Gets the length of longest string in a ``1d`` or ``2d`` sequence of strings.

        Args:
            seq_obj (Sequence[str] | Sequence[Sequence[str]]): Input Sequence

        Returns:
            int: Length of longest string if found; Otherwise ``-1``

        .. versionadded:: 0.6.7
        """
        return cls._get_extreme_element_value_str(
            seq_obj=seq_obj,
            biggest=True,
        )

    @classmethod
    def get_smallest_str(cls, seq_obj: Sequence[str] | Sequence[Sequence[str]]) -> int:
        """
        Gets the length of shortest string in a ``1d`` or ``2d`` sequence of strings.

        Args:
            seq_obj (Sequence[str] | Sequence[Sequence[str]]): Input Sequence

        Returns:
            int: Length of shortest string if found; Otherwise ``-1``
        """
        return cls._get_extreme_element_value_str(
            seq_obj=seq_obj,
            biggest=False,
        )

    # endregion Get Smallest or Largest string in a 1d or 2d sequence

    # region convert_1d_to_2d()

    @overload
    @staticmethod
    def convert_1d_to_2d(seq_obj: Sequence[T], col_count: int) -> List[List[T]]: ...

    @overload
    @staticmethod
    def convert_1d_to_2d(seq_obj: Sequence[T], col_count: int, empty_cell_val: Any) -> List[List[T]]: ...

    @staticmethod
    def convert_1d_to_2d(seq_obj: Sequence[T], col_count: int, empty_cell_val: Any = gUtil.NULL_OBJ) -> List[List[T]]:
        """
        Converts a ``1d`` sequence into a ``2d`` list.

        Args:
            seq_obj (Sequence[T]): Input sequence
            col_count (int): the number of columns to create in the ``2d`` list.
            empty_cell_val (Any, optional): When included any columns missing in last row will be added containing this value.
                ``None`` is also an acceptable value.

        Raises:
            ValueError: If ``col_count`` is less then ``1``.

        Returns:
            List[List[T]]: ``2d`` list.

        .. versionadded:: 0.6.7
        """
        # if len(seq_obj) == 0:
        #     return []
        if col_count < 1:
            raise ValueError("Cols must not be less than 1")

        auto_fill = empty_cell_val is not gUtil.NULL_OBJ

        # see also: https://tinyurl.com/2elcg6fz
        def convert(lst, var_lst):
            idx = 0
            for var_len in var_lst:
                yield lst[idx : idx + var_len]
                idx += var_len

        if len(seq_obj) == 0:
            return [[empty_cell_val for _ in range(col_count)]] if auto_fill else [[]]
        lst = list(seq_obj)

        rows, overflow = divmod(len(seq_obj), col_count)
        if overflow > 0:
            if auto_fill:
                for _ in range(col_count - overflow):
                    lst.append(empty_cell_val)
            # add a new row
            rows += 1
        var_lst = [col_count for _ in range(rows)]

        new_lst = list(convert(lst, var_lst))
        return new_lst

    # endregion convert_1d_to_2d()
