from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno


from ooodev.adapter.drawing.generic_draw_page_comp import GenericDrawPageComp
from ooodev.adapter.drawing.shapes2_partial import Shapes2Partial
from ooodev.adapter.drawing.shapes3_partial import Shapes3Partial
from ooodev.draw.partial.draw_page_partial import DrawPagePartial
from ooodev.draw.shapes.partial.shape_factory_partial import ShapeFactoryPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.proto.component_proto import ComponentT
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write.write_forms import WriteForms

if TYPE_CHECKING:
    from com.sun.star.drawing import XDrawPage
    from ooodev.draw.shapes.shape_base import ShapeBase

_T = TypeVar("_T", bound="ComponentT")


class WriteDrawPage(
    LoInstPropsPartial,
    WriteDocPropPartial,
    DrawPagePartial[_T],
    Generic[_T],
    GenericDrawPageComp,
    Shapes2Partial,
    Shapes3Partial,
    QiPartial,
    PropPartial,
    StylePartial,
    ShapeFactoryPartial["WriteDrawPage[_T]"],
):
    """Represents writer Draw Page."""

    def __init__(self, owner: _T, component: XDrawPage, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XDrawPage): UNO object that supports ``com.sun.star.drawing.GenericDrawPage`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        if not isinstance(owner, WriteDocPropPartial):
            raise TypeError("WriteDocPropPartial is not inherited by owner.")
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        DrawPagePartial.__init__(self, owner=self, component=component, lo_inst=self.lo_inst)
        GenericDrawPageComp.__init__(self, component)  # type: ignore
        Shapes2Partial.__init__(self, component=component, interface=None)  # type: ignore
        Shapes3Partial.__init__(self, component=component, interface=None)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        StylePartial.__init__(self, component=component)
        ShapeFactoryPartial.__init__(self, owner=self, lo_inst=self.lo_inst)
        self._forms = None

    def __len__(self) -> int:
        """
        Gets the number of shapes in the draw page.

        Returns:
            int: Number of shapes in the draw page.
        """
        return self.get_count()

    def __getitem__(self, idx: int) -> ShapeBase[WriteDrawPage[_T]]:
        """
        Gets the shape at the specified index.

        Args:
            idx (int): The index of the shape.

        Returns:
            ShapeBase[WriteDrawPage[_T]]: _description_
        """
        shape = self.component.getByIndex(idx)  # type: ignore
        return self.shape_factory(shape)

    def __next__(self) -> ShapeBase[WriteDrawPage[_T]]:
        """
        Gets the next shape in the draw page.

        Returns:
            ShapeBase[WriteDrawPage[_T]]: The next shape in the draw page.
        """
        shape = super().__next__()
        return self.shape_factory(shape)

    # region Properties
    @property
    def owner(self) -> _T:
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

    @property
    def forms(self) -> WriteForms:
        """
        Gets the forms of the draw page.
        """
        if self._forms is None:
            self._forms = WriteForms(owner=self, forms=self.component.getForms(), lo_inst=self.lo_inst)  # type: ignore
        return self._forms

    # endregion Properties
