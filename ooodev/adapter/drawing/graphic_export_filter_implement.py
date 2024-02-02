from __future__ import annotations
import uno
from typing import TYPE_CHECKING

from com.sun.star.drawing import XGraphicExportFilter

from ooodev.loader import lo as mLo
from .graphic_export_filter_comp import GraphicExportFilterComp

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst


class GraphicExportFilterImplement(GraphicExportFilterComp):
    """Creates an instance of ``com.sun.star.drawing.GraphicExportFilter`` UNO component."""

    def __init__(self, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            _lo = mLo.Lo.current_lo
        else:
            _lo = lo_inst
        comp = _lo.create_instance_mcf(
            XGraphicExportFilter, "com.sun.star.drawing.GraphicExportFilter", raise_err=True
        )
        GraphicExportFilterComp.__init__(self, comp)  # type: ignore
