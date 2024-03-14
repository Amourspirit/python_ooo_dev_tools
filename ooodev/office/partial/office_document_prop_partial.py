from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ooodev.proto.office_document_t import OfficeDocumentT
else:
    OfficeDocumentT = object


class OfficeDocumentPropPartial:
    """A partial class for Office Document."""

    def __init__(self, office_doc: OfficeDocumentT) -> None:
        self.__office_doc = office_doc

    @property
    def office_doc(self) -> OfficeDocumentT:
        """Office Document."""
        return self.__office_doc
