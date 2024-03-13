from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.draw.shapes.draw_shape import DrawShape
from ooodev.adapter.drawing.fill_properties_partial import FillPropertiesPartial
from ooodev.adapter.drawing.line_properties_partial import LinePropertiesPartial
from ooodev.adapter.drawing.shadow_properties_partial import ShadowPropertiesPartial
from ooodev.adapter.drawing.text_properties_partial import TextPropertiesPartial
from ooodev.adapter.document.link_target_properties_partial import LinkTargetPropertiesPartial
from ooodev.adapter.presentation.shape_properties_partial import ShapePropertiesPartial
from ooodev.office.partial.office_document_prop_partial import OfficeDocumentPropPartial
from ooodev.loader.inst.doc_type import DocType

from ooodev.adapter.drawing.rotation_descriptor_properties_partial import RotationDescriptorPropertiesPartial
from ooodev.utils import info as mInfo
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst


class ShapeClassFactory(OfficeDocumentPropPartial):
    def __init__(
        self, owner: Any, component: Any, class_name: str = "DrawShape", lo_inst: LoInst | None = None
    ) -> None:
        if not isinstance(owner, OfficeDocumentPropPartial):
            raise ValueError("owner must be an instance of OfficeDocumentPropPartial")
        OfficeDocumentPropPartial.__init__(self, owner.office_doc)
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        self._lo_inst = lo_inst
        self._name = class_name
        self._component = component
        self._bases: list = [DrawShape]
        self._supported = set(mInfo.Info.get_available_services(self._component))
        if "com.sun.star.drawing.LineProperties" in self._supported:
            self._bases.append(LinePropertiesPartial)
        if "com.sun.star.drawing.FillProperties" in self._supported:
            self._bases.append(FillPropertiesPartial)
        if "com.sun.star.drawing.ShadowProperties" in self._supported:
            self._bases.append(ShadowPropertiesPartial)
        if "com.sun.star.drawing.RotationDescriptor" in self._supported:
            self._bases.append(RotationDescriptorPropertiesPartial)
        if "com.sun.star.drawing.TextProperties" in self._supported:
            self._bases.append(TextPropertiesPartial)
        if "com.sun.star.document.LinkTarget" in self._supported:
            self._bases.append(LinkTargetPropertiesPartial)
        if self.office_doc.DOC_TYPE == DocType.IMPRESS:
            # Even though shapes in Draw and Impress are the same, the supported services are different.
            # A shape in Draw may say it supports a presentation service but it may not.
            # When a shape is drawn on a Drawpage in the Draw program. and the the draw.get_selected_shapes() is used to get
            # the selected shapes, the shape will not have the presentation service even thought is is reported in the getAvailableServices()
            if "com.sun.star.presentation.Shape" in self._supported:
                self._bases.append(ShapePropertiesPartial)

    def _generate_class(self, **kwargs) -> type:
        return type(self._name, tuple(self._bases), kwargs)

    def get_class(self, **kwargs) -> Any:
        clazz = self._generate_class(**kwargs)
        instance = clazz(self._owner, self._component, self._lo_inst)
        if len(self._bases) > 1:
            for i in range(1, len(self._bases)):
                self._bases[i].__init__(instance, self._component)  # type: ignore
        return instance
