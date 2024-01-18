from __future__ import annotations
import uno
from com.sun.star.drawing import XGraphicExportFilter

from ooodev.utils import lo as mLo
from .graphic_export_filter_comp import GraphicExportFilterComp


class GraphicExportFilterImplement(GraphicExportFilterComp):
    """Creates an instance of ``com.sun.star.drawing.GraphicExportFilter`` UNO component."""

    def __init__(self) -> None:
        comp = mLo.Lo.create_instance_mcf(
            XGraphicExportFilter, "com.sun.star.drawing.GraphicExportFilter", raise_err=True
        )
        GraphicExportFilterComp.__init__(self, comp)  # type: ignore
