from __future__ import annotations
from typing import TYPE_CHECKING, Union

from ooodev.format.inner.kind.format_kind import FormatKind as FormatKind

if TYPE_CHECKING:
    from typing_extensions import Protocol
else:
    Protocol = object

# How to declare a Protocol with a field which supports both a simple type and property?
# https://stackoverflow.com/questions/68325221/how-to-declare-a-protocol-with-a-field-which-supports-both-a-simple-type-and-pro


class SizeStruct(Protocol):
    """Protocol Class for size"""

    Width: int
    """Size Width."""
    Height: int
    """Size Height"""


class SizeClass(Protocol):
    """Protocol Class for size"""

    @property
    def Width(self) -> int:
        """Size Width."""
        raise NotImplementedError()

    @property
    def Height(self) -> int:
        """Size Height"""
        raise NotImplementedError()


if TYPE_CHECKING:
    SizeObj = Union[SizeStruct, SizeClass]
else:
    SizeObj = object
