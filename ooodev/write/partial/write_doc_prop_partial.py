from __future__ import annotations

from typing import Any, TYPE_CHECKING
from ooodev.office.partial.office_document_prop_partial import OfficeDocumentPropPartial

if TYPE_CHECKING:
    from ooodev.write.write_doc import WriteDoc
else:
    WriteDoc = Any


class WriteDocPropPartial(OfficeDocumentPropPartial):
    """A partial class for Write Document."""

    def __init__(self, obj: WriteDoc) -> None:
        self.__write_doc = obj
        OfficeDocumentPropPartial.__init__(self, obj)

    @property
    def write_doc(self) -> WriteDoc:
        """Write Document."""
        return self.__write_doc
