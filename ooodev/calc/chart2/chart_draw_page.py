from __future__ import annotations
from typing import TYPE_CHECKING
import uno


from ooodev.adapter.drawing.shapes2_partial import Shapes2Partial
from ooodev.adapter.drawing.shapes3_partial import Shapes3Partial
from ooodev.draw.partial.draw_page_partial import DrawPagePartial
from ooodev.draw.shapes.partial.shape_factory_partial import ShapeFactoryPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.adapter.drawing.shape_collection_comp import ShapeCollectionComp
from ooodev.adapter.lang.component_partial import ComponentPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial
from ooodev.calc.partial.calc_sheet_prop_partial import CalcSheetPropPartial

if TYPE_CHECKING:
    from com.sun.star.drawing import XDrawPage
    from ooodev.draw.shapes.shape_base import ShapeBase
    from ooodev.calc.chart2.table_chart import TableChart


class ChartDrawPage(
    LoInstPropsPartial,
    DrawPagePartial["ChartDrawPage"],
    ShapeCollectionComp,
    Shapes2Partial,
    Shapes3Partial,
    ComponentPartial,
    CalcDocPropPartial,
    CalcSheetPropPartial,
    QiPartial,
    ServicePartial,
    EventsPartial,
    PropPartial,
    StylePartial,
    ShapeFactoryPartial["ChartDrawPage"],
):
    """Represents writer Draw Page."""

    def __init__(self, owner: TableChart, component: XDrawPage, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XDrawPage): UNO object that supports ``com.sun.star.drawing.GenericDrawPage`` service.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        DrawPagePartial.__init__(self, owner=self, component=component, lo_inst=self.lo_inst)
        ShapeCollectionComp.__init__(self, component=component)  # type: ignore
        Shapes2Partial.__init__(self, component=component, interface=None)  # type: ignore
        Shapes3Partial.__init__(self, component=component, interface=None)  # type: ignore
        ComponentPartial.__init__(self, component=component, interface=None)  # type: ignore
        CalcDocPropPartial.__init__(self, owner.calc_doc)
        CalcSheetPropPartial.__init__(self, owner.calc_sheet)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        EventsPartial.__init__(self)
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        StylePartial.__init__(self, component=component)
        ShapeFactoryPartial.__init__(self, owner=self, lo_inst=self.lo_inst)
        self._forms = None

    def __len__(self) -> int:
        return self.get_count()

    def __getitem__(self, index: int) -> ShapeBase[ChartDrawPage]:
        shape = self.component.getByIndex(index)  # type: ignore
        return self.shape_factory(shape)

    def __next__(self) -> ShapeBase[ChartDrawPage]:
        shape = super().__next__()
        return self.shape_factory(shape)

    # region Overrides

    # endregion Overrides

    # region Properties

    @property
    def owner(self) -> TableChart:
        """Owner of this component."""
        return self._owner

    @property
    def name(self) -> str:
        """
        Gets/Sets the name of the draw page.

        Note:
            Naming for Impress pages seems a little different then Draw pages.
            Attempting to name a Draw page `Slide #` where `#` is a number will fail and Draw will auto name the page.
            It seems that `Slide` followed by a space and a number is reserved for Impress.
        """
        return self.component.Name  # type: ignore

    @name.setter
    def name(self, value: str) -> None:
        self.component.Name = value  # type: ignore

    # endregion Properties
