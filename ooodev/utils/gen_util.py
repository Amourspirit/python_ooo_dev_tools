# coding: utf-8
from typing import Iterable, Iterator, Sequence, List, NamedTuple, Any, Tuple, overload
from inspect import isclass


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
        pow = 1
        col_num = 0
        for letter in name[::-1]:
                col_num += (int(letter, 36) -9) * pow
                pow *= 26
        return col_num

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

    @staticmethod
    def make_2d_array(num_rows: int, num_cols: int, val=1) -> List[List[Any]]:
        if num_cols == 0 or num_rows == 0:
            return []
        data = [[val] * num_cols for _ in range(num_rows)]
        return data

    make_2d_list = make_2d_array

    @staticmethod
    def to_list(iter_obj: Iterable[Any]) -> List[Any]:
        """
        Converts an iterable of objects into a list of objects

        Args:
            iter_obj (Iterable[Any]): iterable object.

        Returns:
            List[Any]: List containing same elements of itter_obj
        """
        if Util.is_iterable(iter_obj):
            return list(iter_obj)
        return [iter_obj]

    @staticmethod
    def to_tuple(iter_obj: Iterable[Any]) -> Tuple[Any]:
        """
        Converts an iterable of objects into a tuple of objects

        Args:
            iter_obj (Iterable[Any]): iterable object.

        Returns:
            Tuple[Any]: Tuple containing same elements of itter_obj
        """
        if Util.is_iterable(iter_obj):
            return tuple(iter_obj)
        return (iter_obj,)

    @classmethod
    def to_2d_list(cls, seq_obj: Sequence[Sequence[Any]]) -> List[List[Any]]:
        """
        Converts a sequene of sequenc to a list.

        Converts 1-Dimensional or 2-Dimensional array such as a Tuple or a Tuple of Tuple's into a List of List.

        An array of tuples is immutable and can not add or remove elemetns whereas a list is mutable.

        Args:
            seq_obj (Sequence[Sequence[Any]]): Sequence of Sequence

        Returns:
            List[List[Any]]: 2-Dimensional list
        """
        num_rows = len(seq_obj)
        if num_rows == 0:
            return []
        is_2d = False
        try:
            is_2d = Util.is_iterable(seq_obj[0])
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
    def to_2d_tuple(cls, seq_obj: Sequence[Sequence[Any]]) -> Tuple[Tuple[Any, ...], ...]:
        """
        Converts a sequene of sequenc to a list.

        Converts 1-Dimensional or 2-Dimensional array such as a List or List of list's into a Tuple of Tuple.

        Args:
            seq_obj (Sequence[Sequence[Any]]): Sequence of Sequence

        Returns:
            Tuple[Tuple[Any, ...], ...]: 2-Dimensional tuple
        """
        num_rows = len(seq_obj)
        if num_rows == 0:
            return tuple()
        is_2d = False
        try:
            is_2d = Util.is_iterable(seq_obj[0])
        except Exception:
            is_2d = False
        lst = []
        if is_2d:
            for row in seq_obj:
                lst.append(cls.to_tuple(row))
        else:
            lst.append(cls.to_tuple(seq_obj))
        return tuple(lst)


class ArgsHelper:
    class NameValue(NamedTuple):
        name: str
        value: Any

class Util:
    @classmethod
    def is_iterable(cls, arg: object, excluded_types: Iterable[type]=(str,)) -> bool:
        """
        Gets if ``arg`` is iterable.

        Args:
            arg (object): object to test
            excluded_types (Iterable[type], optional): Iterable of type to exlcude.
                If ``arg`` matches any type in ``excluded_types`` then ``False`` will be returned.
                Default ``(str,)``

        Returns:
            bool: ``True`` if ``arg`` is an iterable object and not of a type in ``excluded_types``;
            Otherwise, ``False``.

        Note:
            if ``arg`` is of type str then return result is ``False``.

        Example:
            .. code-block:: python

                

                # non-string iterables    
                assert is_iterable(arg=("f", "f"))       # tuple
                assert is_iterable(arg=["f", "f"])       # list
                assert is_iterable(arg=iter("ff"))       # iterator
                assert is_iterable(arg=range(44))        # generator
                assert is_iterable(arg=b"ff")            # bytes (Python 2 calls this a string)

                # strings or non-iterables
                assert not is_iterable(arg=u"ff")        # string
                assert not is_iterable(arg=44)           # integer
                assert not is_iterable(arg=is_iterable)  # function
                
                # excluded_types, optionally exlcude types
                from enum import Enum, auto

                class Color(Enum):
                    RED = auto()
                    GREEN = auto()
                    BLUE = auto()
                
                assert is_iterable(arg=Color)             # Enum
                assert not is_iterable(arg=Color, excluded_types=(Enum, str)) # Enum
        """
        # if isinstance(arg, str):
        #     return False
        result = False
        try:
            result = isinstance(iter(arg), Iterator)
        except Exception:
            result = False
        if result is True:
            if cls._is_iterable_excluded(arg, excluded_types=excluded_types):
                result = False
        return result

    @staticmethod
    def _is_iterable_excluded(arg: object, excluded_types: Iterable) -> bool:
        try:
            isinstance(iter(excluded_types), Iterator)
        except Exception:
            return False

        if len(excluded_types) == 0:
            return False

        def _is_instance(obj: object) -> bool:
            # when obj is instance then isinstance(obj, obj) raises TypeError
            # when obj is not instance then isinstance(obj, obj) return False
            try:
                if not isinstance(obj, obj):
                    return False
            except TypeError:
                pass
            return True
        ex_types = excluded_types if isinstance(excluded_types, tuple) else tuple(excluded_types)
        arg_instance = _is_instance(arg)
        if arg_instance is True:
            if isinstance(arg, ex_types):
                return True
            return False
        if isclass(arg) and issubclass(arg, ex_types):
            return True
        return arg in ex_types