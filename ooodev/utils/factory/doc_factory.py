from __future__ import annotations
from typing import Any
import uno
from com.sun.star.frame import XModule
from ooodev.utils.inst.lo.lo_inst import LoInst
from ooodev.utils import lo as mLo


def doc_factory(doc: Any, lo_inst: LoInst | None) -> Any:
    """
    Gets an instance of a document.

    Args:
        doc (Any): Office document such as Writer, Calc, Draw, Impress.
        lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents.

    Raises:
        ValueError: If no identifier found.
        ValueError: If unknown identifier.

    Returns:
        Any: A document instance. Such as ``ooodev.draw.DrawDoc`` or ``ooodev.calc.CalcDoc``.
    """
    if lo_inst is None:
        lo_inst = mLo.Lo.current_lo
    module = lo_inst.qi(XModule, doc, True)
    identifier = module.getIdentifier()
    if not identifier:
        raise ValueError("No identifier found.")
    if identifier == "com.sun.star.drawing.DrawingDocument":
        from ooodev.draw import DrawDoc

        return DrawDoc(doc, lo_inst)
    if identifier == "com.sun.star.presentation.PresentationDocument":
        from ooodev.draw import ImpressDoc

        return ImpressDoc(doc, lo_inst)
    if identifier == "com.sun.star.sheet.SpreadsheetDocument":
        from ooodev.calc import CalcDoc

        return CalcDoc(doc, lo_inst)
    if identifier == "com.sun.star.text.TextDocument":
        from ooodev.write import WriteDoc

        return WriteDoc(doc, lo_inst)
    raise ValueError(f"Unknown identifier: {identifier}")
