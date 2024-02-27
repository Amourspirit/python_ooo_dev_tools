from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.frame import XModule

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.proto.office_document_t import OfficeDocumentT

# pylint: disable=import-outside-toplevel


def is_known_doc(doc: Any, lo_inst: LoInst) -> bool:
    """
    Checks if the document is known.

    Args:
        doc (Any): Office document such as Writer, Calc, Draw, Impress.
        lo_inst (LoInst): Lo Instance. Can be ``Lo.current_lo``.

    Returns:
        bool: ``True`` if known, ``False`` otherwise.
    """
    try:
        module = lo_inst.qi(XModule, doc)
        if not module:
            return False
        identifier = module.getIdentifier()
        if not identifier:
            return False
        return identifier in (
            "com.sun.star.drawing.DrawingDocument",
            "com.sun.star.presentation.PresentationDocument",
            "com.sun.star.sheet.SpreadsheetDocument",
            "com.sun.star.text.TextDocument",
        )
    except Exception:
        return False


def doc_factory(doc: Any, lo_inst: LoInst) -> OfficeDocumentT:
    """
    Gets an instance of a document.

    Args:
        doc (Any): Office document such as Writer, Calc, Draw, Impress.
        lo_inst (LoInst): Lo Instance. Can be ``Lo.current_lo``.

    Raises:
        ValueError: If no identifier found.
        ValueError: If unknown identifier.

    Returns:
        OfficeDocumentT: A document instance. Such as ``ooodev.draw.DrawDoc`` or ``ooodev.calc.CalcDoc``.
    """
    module = lo_inst.qi(XModule, doc, True)
    identifier = module.getIdentifier()
    if not identifier:
        raise ValueError("No identifier found.")
    if identifier == "com.sun.star.drawing.DrawingDocument":
        from ooodev.draw.draw_doc import DrawDoc

        return DrawDoc(doc, lo_inst)
    if identifier == "com.sun.star.presentation.PresentationDocument":
        from ooodev.draw.impress_doc import ImpressDoc

        return ImpressDoc(doc, lo_inst)
    if identifier == "com.sun.star.sheet.SpreadsheetDocument":
        from ooodev.calc.calc_doc import CalcDoc

        return CalcDoc(doc, lo_inst)
    if identifier == "com.sun.star.text.TextDocument":
        from ooodev.write.write_doc import WriteDoc

        return WriteDoc(doc, lo_inst)
    raise ValueError(f"Unknown identifier: {identifier}")
