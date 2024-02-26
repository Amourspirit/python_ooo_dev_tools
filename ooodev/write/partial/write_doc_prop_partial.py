from __future__ import annotations

from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:
    from ooodev.write.write_doc import WriteDoc
else:
    WriteDoc = Any


class WriteDocPropPartial:
    """A partial class for Write Document."""

    def __init__(self, obj: WriteDoc) -> None:
        self.__write_doc = obj

    @property
    def write_doc(self) -> WriteDoc:
        """Write Document."""
        return self.__write_doc
