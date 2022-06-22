# coding: utf-8
from __future__ import annotations
from typing import Sequence, TypeVar, Union, Any, Tuple, List
from os import PathLike

import uno
from com.sun.star.text import XText
from com.sun.star.text import XTextCursor
from com.sun.star.text import XTextDocument

PathOrStr = Union[str, PathLike]
"""Path like object or string"""

UnoInterface = object
"""Represents a uno interface class. Any uno Class that starts with X"""

T = TypeVar("T")

Row = Sequence[Any]
"""Represents a Row of a Table."""

Column = Sequence[Any]
"""Represents a Column of a Table."""

Table = Sequence[Row]
"""Represents a Table of Rows and Columns"""

TupleArray = Tuple[Tuple[Any, ...], ...]
"""Table like tuples with rows and columns"""

FloatList = List[float]
"""List of Floats"""

FloatTable = List[FloatList]
"""Table like array of floats with rows and columns"""

DocOrCursor = Union[XTextDocument, XTextCursor]
"""Type of Text Dcoument or Cursor"""

DocOrText = Union[XTextDocument, XText]
"""Type of Text Document of Text"""