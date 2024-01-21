from __future__ import annotations
from ooodev.events.partial.events_partial import EventsPartial
from .export_base import ExportBase


class ShapeExportBase(ExportBase, EventsPartial):
    """Base Class used for exporting Draw Shapes in various ways."""

    def __init__(self):
        ExportBase.__init__(self)
        EventsPartial.__init__(self)
