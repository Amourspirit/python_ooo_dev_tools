# coding: utf-8
"""General Utilities"""
from __future__ import annotations
import string
from typing import Callable, Iterable, Iterator, Sequence, List, NamedTuple, Any, Tuple, cast, overload
from inspect import isclass
import string
from .type_var import DictTable, Table
from . import gen_util as gUtil


class TableHelper:
    @staticmethod
    def col_name_to_int(name: str) -> int:
        """
        Converts a Column Name into an int.
        Results are one based so ``a`` converts to ``1``

        Args:
            name (str):Case insensitive column name such as 'a' or 'AB'

        Returns:
            int: One based int representing column name
        """
        chars = name.rstrip(string.digits)
        pow = 1
        col_num = 0
        for letter in chars[::-1]:  # reverse chars
            col_num += (int(letter, 36) - 9) * pow
            pow *= 26
        return col_num

    @staticmethod
    def row_name_to_int(name: str) -> int:
        """
        Converts a row name into an int.
        Leading Alpha chars are ignore. ``'4'`` converts to ``4``. ``'C5'`` converts to ``5``

        Args:
            name (str): row name to convert

        Returns:
            int: converted name as int.
        """
        chars = name.rstrip(string.digits + "-")
        if chars:
            s = name[len(chars)]  # drop leading chars that are not numbers.
        else:
            s = name
        result = int(s)
        if result < 0:
            raise ValueError(f"Cannot parse negative values: {name}")
        return result

    @classmethod
    def make_cell_name(cls, row: int, col: int) -> str:
        """
        Convert given row and column number to ``A1`` style cell name.

        Args:
            row (int): Row number. This is a 1 based value.
            col (int): Column Number. This is 1 based value.

        Raises:
            ValueError: If row or col value < 1

        Returns:
            str: row and col as cell name such as A1, AB3
        """
        if row < 1:
            raise ValueError(f"row is one based. Value cannot be less then 1: {row}")
        return f"{cls.make_column_name(col)}{row}"

    @staticmethod
    def make_column_name(col: int) -> str:  # col is 1 based
        """
        Makes a cell style name. eg: A, B, C, ... AA, AB, AC

        Args:
            col (int): Column number. This is a one based value.

        Raises:
            ValueError: If col value < 1

        Returns:
            str: column name. eg: A, B, C, ... AA, AB, AC
        """
        if col < 1:
            raise ValueError(f"col is one based. Value cannot be less then 1: {col}")
        str_col = str()
        div = col
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
    def to_list(iter_obj: Iterable[Any] | object) -> List[Any]:
        """
        Converts an iterable of objects into a list of objects

        If ``iter_obj`` is not iterable it will be return as a tuple containing ``iter_obj``

        Args:
            iter_obj (Iterable[Any] | object): iterable object or object.

        Returns:
            List[Any]: List containing same elements of ``itter_obj``
        """
        if gUtil.Util.is_iterable(iter_obj):
            return list(iter_obj)
        return [iter_obj]

    @staticmethod
    def to_tuple(iter_obj: Iterable[Any] | object) -> Tuple[Any]:
        """
        Converts an iterable of objects or object into a tuple of objects

        If ``iter_obj`` is not iterable it will be return as a tuple containing ``iter_obj``

        Args:
            iter_obj (Iterable[Any] | object): iterable object or object.

        Returns:
            Tuple[Any]: Tuple containing same elements of ``itter_obj``
        """
        if gUtil.Util.is_iterable(iter_obj):
            return tuple(iter_obj)
        return (iter_obj,)

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
            for row in seq_obj:
                lst.append(cls.to_list(row))
        else:
            lst.append(cls.to_list(seq_obj))
        return lst

    @classmethod
    def to_2d_tuple(cls, seq_obj: Sequence[Any]) -> Tuple[Tuple[Any, ...], ...]:
        """
        Converts a sequence of sequence to a list.

        Converts 1-Dimensional or 2-Dimensional array such as a List or List of list's into a Tuple of Tuple.

        Args:
            seq_obj (Sequence[Any]): 1-Dimensional or 2-Dimensional Sequence
        Returns:
            Tuple[Tuple[Any, ...], ...]: 2-Dimensional tuple
        """
        num_rows = len(seq_obj)
        if num_rows == 0:
            return tuple()
        is_2d = False
        try:
            is_2d = gUtil.Util.is_iterable(seq_obj[0])
        except Exception:
            is_2d = False
        lst = []
        if is_2d:
            for row in seq_obj:
                lst.append(cls.to_tuple(row))
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
            cols = [value for value in tbl[0]]
            data = []
            for i, row in enumerate(tbl):
                if i == 0:  # col row
                    continue
                data.append(dict(zip(cols, row)))
            return data
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
            cols = [k for k in first.keys()]
            data = [cols]
            for row in tbl:
                data.append([v for _, v in row.items()])
            return data
        except Exception as e:
            raise e
