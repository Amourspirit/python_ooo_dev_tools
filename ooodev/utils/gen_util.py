# coding: utf-8
"""General Utilities"""

from __future__ import annotations
import contextlib
import re
from typing import Iterator, NamedTuple, Any, Sequence, TypeVar
from inspect import isclass
import secrets
import string

# match:
#   Any uppercase character that is not at the start of a line
#   Any Number that is preceded by a Upper or Lower case character
_REG_TO_SNAKE = re.compile(r"(?<!^)(?=[A-Z])|(?<=[A-zA-Z])(?=[0-9])")  # re.compile(r"(?<!^)(?=[A-Z])")
_REG_LETTER_AFTER_NUMBER = re.compile(r"(?<=\d)(?=[a-zA-Z])")

NULL_OBJ = object()
"""Null Object uses with None is not an option"""

TNullObj = TypeVar("TNullObj", bound=object)


class ArgsHelper:
    "Args Helper"

    class NameValue(NamedTuple):
        "Name Value pair"
        name: str
        """Name component"""
        value: Any
        """Value component"""


class Util:
    @staticmethod
    def get_obj_full_name(obj: Any) -> str:
        """
        Gets the full name of an object. The full name is the module and class name.

        Args:
            obj (Any): Object to get full name of.

        Returns:
            str: Full name of object on success; Otherwise, empty string.
        """
        if not obj:
            return ""
        with contextlib.suppress(Exception):
            return f"{obj.__class__.__module__}.{obj.__class__.__name__}"
        with contextlib.suppress(Exception):
            return f"{obj.__module__}.{obj.__name__}"
        return ""

    @staticmethod
    def generate_random_string(length: int = 8) -> str:
        """
        Generates a random string.

        Args:
            length (int, optional): Length of string to generate. Default ``8``

        Returns:
            str: Random string
        """
        return "".join(secrets.choice(string.ascii_letters) for _ in range(length))

    @staticmethod
    def generate_random_hex_string(length: int = 8) -> str:
        """
        Generates a random string.

        Args:
            length (int, optional): Length of string to generate. Default ``8``

        Returns:
            str: Random string
        """
        return "".join(secrets.choice(string.hexdigits) for _ in range(length))

    @staticmethod
    def generate_random_alpha_numeric(length: int = 8) -> str:
        """
        Generates a random alpha numeric string.

        Args:
            length (int, optional): Length of string to generate. Default ``8``

        Returns:
            str: Random string
        """
        s = string.ascii_letters + string.digits
        return "".join(secrets.choice(s) for _ in range(length))

    @classmethod
    def is_iterable(cls, arg: Any, excluded_types: Sequence[type] | None = None) -> bool:
        """
        Gets if ``arg`` is iterable.

        Args:
            arg (object): object to test
            excluded_types (Iterable[type], optional): Iterable of type to exclude.
                If ``arg`` matches any type in ``excluded_types`` then ``False`` will be returned.
                Default ``(str,)``

        Returns:
            bool: ``True`` if ``arg`` is an iterable object and not of a type in ``excluded_types``;
            Otherwise, ``False``.

        Note:
            if ``arg`` is of type str then return result is ``False``.

        .. collapse:: Example

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

                # excluded_types, optionally exclude types
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
        if excluded_types is None:
            excluded_types = (str,)
        result = False
        try:
            result = isinstance(iter(arg), Iterator)
        except Exception:
            result = False
        if result and cls._is_iterable_excluded(arg, excluded_types=excluded_types):
            result = False
        return result

    @staticmethod
    def _is_iterable_excluded(arg: object, excluded_types: Sequence) -> bool:
        try:
            isinstance(iter(excluded_types), Iterator)
        except Exception:
            return False

        if len(excluded_types) == 0:
            return False

        def _is_instance(obj: object) -> bool:
            # when obj is instance then isinstance(obj, obj) raises TypeError
            # when obj is not instance then isinstance(obj, obj) return False
            with contextlib.suppress(TypeError):
                if not isinstance(obj, obj):  # type: ignore
                    return False
            return True

        ex_types = excluded_types if isinstance(excluded_types, tuple) else tuple(excluded_types)
        arg_instance = _is_instance(arg)
        if arg_instance is True:
            return isinstance(arg, ex_types)
        return True if isclass(arg) and issubclass(arg, ex_types) else arg in ex_types

    @staticmethod
    def camel_to_snake(name: str) -> str:
        """Converts CamelCase to snake_case

        Args:
            name (str): CamelCase string

        Returns:
            str: snake_case string

        Note:
            This method is preferred over the `to_snake_case` method when converting CamelCase strings.
            It does a better job of handling leading caps. ``UICamelCase`` will be converted to ``ui_camel_case`` and not ``u_i_camel_case``.
        """
        # This function uses regular expressions to insert underscores between the lowercase and uppercase letters, then converts the entire string to lowercase.

        # The first `re.sub` call inserts an underscore before any uppercase letter that is preceded by a lowercase letter or a number.
        # The second `re.sub` call inserts an underscore before any uppercase letter that is followed by a lowercase letter.
        # The `lower` method then converts the entire string to lowercase.
        name = re.sub("(.)([A-Z][a-z]+)", r"\1_\2", name)
        return re.sub("([a-z0-9])([A-Z])", r"\1_\2", name).lower()

    @staticmethod
    def to_camel_case(s: str) -> str:
        """
        Converts string to ``CamelCase``

        Args:
            s (str): string to convert such as ``snake_case_word`` or ``pascalCaseWord``

        Returns:
            str: string converted to ``CamelCaseWord``
        """
        s = s.strip()
        if not s:
            return ""
        result = s
        if "_" in result:
            return "".join(word.title() for word in result.split("_"))
        return result[:1].upper() + result[1:]

    @classmethod
    def to_pascal_case(cls, s: str) -> str:
        """
        Converts string to ``pascalCase``

        Args:
            s (str): string to convert such as ``snake_case_word`` or  ``CamelCaseWord``

        Returns:
            str: string converted to ``pascalCaseWord``
        """
        result = cls.to_camel_case(s)
        if result:
            result = result[:1].lower() + result[1:]
        return result

    @staticmethod
    def to_snake_case(s: str) -> str:
        """
        Convert string to ``snake_case``

        Args:
            s (str): string to convert such as ``pascalCaseWord`` or  ``CamelCaseWord``

        Returns:
            str: string converted to ``snake_case_word``
        """
        s = s.strip()
        if not s:
            return ""
        result = _REG_TO_SNAKE.sub("_", s)
        result = _REG_LETTER_AFTER_NUMBER.sub("_", result)
        return result.lower()

    @classmethod
    def to_snake_case_upper(cls, s: str) -> str:
        """
        Convert string to ``SNAKE_CASE``

        Args:
            s (str): string to convert such as ``snake_case_word`` or ``pascalCaseWord`` or  ``CamelCaseWord``

        Returns:
            str: string converted to ``SNAKE_CASE_WORD``
        """
        result = cls.to_snake_case(s)
        return result.upper() if s else ""

    @staticmethod
    def to_single_space(s: str, strip=True) -> str:
        """
        Gets a string with multiple spaces converted to single spaces

        Args:
            s (str): String
            strip (bool, optional): If ``True`` then whitespace is stripped from start and end or string. Default ``True``.

        Returns:
            str: String with extra spaces removed.

        Example:
            .. code-block:: python

                >>> s = ' The     quick brown    fox'
                >>> print(Util.to_single_space(s))
                'The quick brown fox'
        """
        if not s:
            return ""
        result = re.sub(" +", " ", s)
        return result.strip() if strip else result

    @staticmethod
    def _atoi(text):
        return int(text) if text.isdigit() else text

    @staticmethod
    def natural_key_sorter(text: str) -> list:
        """
        Sort Key Sorts in human order.

        # Example:

        .. code-block:: python

            a_list.sort(key=Util.natural_key_sorter)
        """
        # a_list.sort(key=natural_keys) sorts in human order
        # http://nedbatchelder.com/blog/200712/human_sorting.html
        # (See Toothy's implementation in the comments)
        # https://stackoverflow.com/questions/5967500/how-to-correctly-sort-a-string-with-a-number-inside
        return [Util._atoi(c) for c in re.split(r"(\d+)", text)]

    @staticmethod
    def get_index(idx: int, count: int, allow_greater: bool = False) -> int:
        """
        Gets the index.

        Args:
            idx (int): Index of element. Can be a negative value to index from the end of the list.
            count (int): Number of elements.
            allow_greater (bool, optional): If True and index is greater then the count
                then the index becomes the next index if element is appended. Defaults to False.
                Only affect the ``-1`` index.

        Raises:
            ValueError: If ``count`` is less than ``0``.
            IndexError: If index is out or range.

        Returns:
            int: Index value.

        Note:
            ``-1`` is the last index in the sequence. Unless ``allow_greater`` is ``True`` then ``-1`` last index ``+ 1``.
            Only the ``-1`` is treated differently when ``allow_greater`` is ``True``.

            ``-2`` is the second to last index in the sequence. ``10`` items and ``idx=-2`` then index ``8`` is returned.
            ``-3`` is the third to last index in the sequence. ``10`` items and ``idx=-3`` then index ``7`` is returned.

        .. versionadded:: 0.20.2
        """
        if count < 0:
            raise ValueError("count cannot be less than 0")
        if idx == -1:
            index = 0
            if allow_greater:
                if count == 0:
                    return index
                else:
                    index = count
            else:
                index = count - 1
            if index < 0:
                raise IndexError("list index out of range")
            return index

        index = idx
        if index < 0:
            index = count + index
            if index < 0:
                raise IndexError("list index out of range")
        if index >= count:
            if allow_greater:
                index = count
            else:
                raise IndexError("list index out of range")
        return index
