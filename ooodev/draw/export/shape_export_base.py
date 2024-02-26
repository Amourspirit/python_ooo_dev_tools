from __future__ import annotations
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.draw.export.export_base import ExportBase


class ShapeExportBase(LoInstPropsPartial, ExportBase, EventsPartial):
    """Base Class used for exporting Draw Shapes in various ways."""

    def __init__(self, lo_inst: LoInst | None = None):
        """
        Constructor

        Args:
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        ExportBase.__init__(self)
        EventsPartial.__init__(self)
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
