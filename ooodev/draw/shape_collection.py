from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.drawing import XShapes
from com.sun.star.drawing import XShape

from ooodev.adapter.drawing.shape_collection_comp import ShapeCollectionComp
from ooodev.exceptions import ex as mEx
from ooodev.units.unit_mm import UnitMM
from ooodev.loader import lo as mLo
from ooodev.utils.kind.shape_comb_kind import ShapeCombKind
from ooodev.utils.partial.gui_partial import GuiPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.proto.component_proto import ComponentT

if TYPE_CHECKING:
    from ooodev.draw import DrawPage
    from ooodev.loader.inst.lo_inst import LoInst

# https://wiki.openoffice.org/wiki/Documentation/DevGuide/Drawings/Grouping,_Combining_and_Binding


class ShapeCollection(LoInstPropsPartial, ShapeCollectionComp, QiPartial):
    """Represents a shape collection."""

    def __init__(self, owner: DrawPage[ComponentT], collection: Any = None, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (Any): Usually DrawDoc or ImpressDoc Instance.
            collection (Any, optional): The collection of shapes. If ``None``, a new empty collection will be created.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        if collection is None:
            collection = self.lo_inst.create_instance_mcf(
                XShapes, "com.sun.star.drawing.ShapeCollection", raise_err=True
            )
        # ShapeCollectionComp will validate the collection
        ShapeCollectionComp.__init__(self, collection)
        QiPartial.__init__(self, component=self.component, lo_inst=self.lo_inst)

        self._owner = owner

    def group(self) -> GroupShape:
        """
        Groups the shapes in the collection.

        Raises:
            ShapeError: If error occurs.

        Returns:
            GroupShape: Grouped shape.
        """
        if len(self) < 2:
            raise mEx.ShapeError(f"At least two shapes are required to group. Currently has {len(self)} shapes.")

        try:
            return self.owner.group(self.qi(XShapes, True))
        except mEx.ShapeError:
            raise
        except Exception as e:
            raise mEx.ShapeError("Unable to group shapes") from e

    def combine(self, combine_op: ShapeCombKind | int) -> DrawShape:
        """
        Combines the shapes in the collection.

        Args:
            combine_op (ShapeCombKind | int): Combine operation.
                ``0`` - Merge, ``1`` - Intersect, ``2`` - Subtract, ``3`` - Combine.

        Raises:
            ShapeError: If error occurs.

        Returns:
            DrawShape: Combined shape.

        Note:
            Combining is done using Dispatch Command.
        """
        if len(self) < 2:
            raise mEx.ShapeError(f"At least two shapes are required to combine. Currently has {len(self)} shapes.")
        doc = self.owner.owner  # expecting DrawDoc or ImpressDoc
        if not isinstance(doc, GuiPartial):
            raise mEx.ShapeError("Unable to combine shapes. DrawPage owner does not inherit from GuiPartial.")

        combine_op = ShapeCombKind(combine_op)

        try:
            sel_supp = (
                doc.get_selection_supplier()
            )  # self.lo_inst.qi(XSelectionSupplier, mGui.GUI.get_current_controller(doc), True)
            sel_supp.select(self.qi(XShapes, True))

            if combine_op == ShapeCombKind.INTERSECT:
                self.lo_inst.dispatch_cmd("Intersect")
            elif combine_op == ShapeCombKind.SUBTRACT:
                self.lo_inst.dispatch_cmd("Substract")  # misspelt!
            elif combine_op == ShapeCombKind.COMBINE:
                self.lo_inst.dispatch_cmd("Combine")
            else:
                self.lo_inst.dispatch_cmd("Merge")

            self.lo_inst.delay(500)  # give time for dispatches to arrive and be processed

            # extract the new single shape from the modified selection
            xs = self.lo_inst.qi(XShapes, sel_supp.getSelection(), True)
            shape = self.lo_inst.qi(XShape, xs.getByIndex(0), True)
            return DrawShape(owner=self.owner, component=shape, lo_inst=self.lo_inst)  # type: ignore
        except Exception as e:
            raise mEx.ShapeError("Unable to combine shapes") from e

    def get_height(self) -> UnitMM:
        """Returns the height of the collection."""
        height = 0
        for itm in self:
            shape = self.lo_inst.qi(XShape, itm)
            if shape is None:
                continue
            pos = shape.getPosition()
            size = shape.getSize()
            shape_bottom = size.Height + pos.Y
            if shape_bottom > height:
                height = shape_bottom
        if height == 0:
            return UnitMM(0)
        return UnitMM.from_mm100(height)

    def get_width(self) -> UnitMM:
        """Returns the width of the collection."""
        width = 0
        for itm in self:
            shape = self.lo_inst.qi(XShape, itm)
            if shape is None:
                continue
            pos = shape.getPosition()
            size = shape.getSize()
            shape_right = size.Width + pos.X
            if shape_right > width:
                width = shape_right
        if width == 0:
            return UnitMM(0)
        return UnitMM.from_mm100(width)

    @property
    def owner(self) -> DrawPage[ComponentT]:
        """Returns the owner of this instance."""
        return self._owner


from ooodev.draw.shapes.group_shape import GroupShape
from ooodev.draw.shapes.draw_shape import DrawShape
