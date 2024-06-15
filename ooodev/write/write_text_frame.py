from __future__ import annotations
from typing import Any, TYPE_CHECKING, TypeVar, Generic
import uno
from com.sun.star.drawing import XShape

from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.text.text_frame_comp import TextFrameComp
from ooodev.draw.partial.draw_shape_partial import DrawShapePartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.proto.component_proto import ComponentT
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.draw.shapes.shape_base import ShapeBase

if TYPE_CHECKING:
    from com.sun.star.text import XTextFrame

T = TypeVar("T", bound="ComponentT")


class WriteTextFrame(
    ShapeBase,
    WriteDocPropPartial,
    Generic[T],
    TextFrameComp,
    PropertyChangeImplement,
    VetoableChangeImplement,
    DrawShapePartial,
    QiPartial,
    PropPartial,
    StylePartial,
):
    """Represents writer text content."""

    def __init__(self, owner: T, component: XTextFrame, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XTextFrame): UNO object that supports ``com.sun.star.text.TextFrame`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__owner = owner
        if not isinstance(owner, WriteDocPropPartial):
            raise TypeError("WriteDocPropPartial is not inherited by owner.")
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        # in this case QiPartial needs to be before ShapeBase.
        QiPartial.__init__(self, component=component, lo_inst=lo_inst)  # type: ignore
        ShapeBase.__init__(self, owner=self.__owner, component=component, lo_inst=lo_inst)  # type: ignore
        TextFrameComp.__init__(self, component)  # type: ignore
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)  # type: ignore
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)  # type: ignore
        DrawShapePartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        StylePartial.__init__(self, component=component)

    # region Overrides
    def _ComponentBase__get_is_supported(self, component: Any) -> bool:
        if component is None:
            return False
        shape = self.qi(XShape)
        return shape is not None

    # endregion Overrides

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self.__owner

    @property
    def name(self) -> str:
        """Gets/Sets the name of this text frame."""
        return self.component.Name  # type: ignore

    @name.setter
    def name(self, value: str) -> None:
        self.component.Name = value  # type: ignore

    # endregion Properties
