# coding: utf-8
from __future__ import annotations
from typing import Callable, Protocol, Sequence, TypeVar, Union, Any, Tuple, List, Dict, TYPE_CHECKING
from os import PathLike

import uno


if TYPE_CHECKING:
    from com.sun.star.text import XText
    from com.sun.star.text import XTextCursor
    from com.sun.star.text import XTextDocument
    from ooodev.events.args.event_args import EventArgs
    from ooodev.events.args.cancel_event_args import CancelEventArgs
    from ooodev.events.args.listener_event_args import ListenerEventArgs
else:
    XText = object
    XTextCursor = object
    XTextDocument = object
    EventArgs = object
    ListenerEventArgs = object

PathOrStr = Union[str, PathLike]
"""Path like object or string"""

UnoInterface = object
"""Represents a uno interface class. Any uno Class that starts with X"""

T = TypeVar("T")

Row = Sequence[Any]
"""Represents a Row of a Table."""

DictRow = Dict[str, Any]
"""
Represents a Row of a Table as a dictionary of Key, value.
Where the key is the column name.
"""

DictTable = Sequence[DictRow]
"""
Represents a Dictionary table.
Each Element in the sequence as DictRow (dictionary of key, value)
with key as column name.
"""

Column = Sequence[Any]
"""Represents a Column of a Table."""

Table = Sequence[Row]
"""Represents a 2-D Table of Rows and Columns"""

TupleArray = Tuple[Tuple[Any, ...], ...]
"""Table like tuples with rows and columns"""

FloatList = List[float]
"""List of Floats"""

FloatTable = List[FloatList]
"""Table like array of floats with rows and columns"""

DocOrCursor = Union[XTextDocument, XTextCursor]
"""Type of Text Document or Cursor"""

DocOrText = Union[XTextDocument, XText]
"""Type of Text Document of Text"""


EventCallback = Callable[[Any, Any], None]
"""Event Callback"""

ListenerEventCallbackT = Callable[[Any, ListenerEventArgs], None]
"""Listener Event Callback"""


class EventArgsCallbackT(Protocol):
    def __call__(self, src: Any, event: EventArgs, *args: Any, **kwargs: Any) -> None:
        ...


class CancelEventArgsCallbackT(Protocol):
    def __call__(self, src: Any, event: CancelEventArgs, *args: Any, **kwargs: Any) -> None:
        ...


# EventArgsCallbackT = Callable[[Any, EventArgs], None]
