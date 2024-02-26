from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from com.sun.star.drawing import XGraphicExportFilter

from ooodev.loader import lo as mLo
from ooodev.adapter.drawing.graphic_export_filter_comp import GraphicExportFilterComp

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst


class GraphicExportFilterImplement(GraphicExportFilterComp):
    """Creates an instance of ``com.sun.star.drawing.GraphicExportFilter`` UNO component."""

    def __init__(self, lo_inst: LoInst | None = None) -> None:
        _lo = mLo.Lo.current_lo if lo_inst is None else lo_inst
        comp = _lo.create_instance_mcf(
            XGraphicExportFilter, "com.sun.star.drawing.GraphicExportFilter", raise_err=True
        )
        GraphicExportFilterComp.__init__(self, comp)  # type: ignore
