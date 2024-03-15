from __future__ import annotations
from typing import Any, TYPE_CHECKING, Dict, Type
import uno
from com.sun.star.drawing import XGluePointsSupplier


from ooodev.draw.shapes.draw_shape import DrawShape
from ooodev.adapter.drawing.fill_properties_partial import FillPropertiesPartial
from ooodev.adapter.drawing.line_properties_partial import LinePropertiesPartial
from ooodev.adapter.drawing.shadow_properties_partial import ShadowPropertiesPartial
from ooodev.adapter.drawing.text_properties_partial import TextPropertiesPartial
from ooodev.adapter.document.link_target_properties_partial import LinkTargetPropertiesPartial
from ooodev.adapter.drawing.glue_points_supplier_partial import GluePointsSupplierPartial
from ooodev.adapter.presentation.shape_properties_partial import ShapePropertiesPartial
from ooodev.adapter.drawing.shape_group_partial import ShapeGroupPartial
from ooodev.adapter.drawing.shapes_partial import ShapesPartial
from ooodev.office.partial.office_document_prop_partial import OfficeDocumentPropPartial
from ooodev.loader.inst.doc_type import DocType

from ooodev.adapter.drawing.rotation_descriptor_properties_partial import RotationDescriptorPropertiesPartial
from ooodev.utils import info as mInfo
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst


class ShapeClassFactory(OfficeDocumentPropPartial):
    """ "
    Generates a class instance of a shape.

    If a shape supports a service, a partial class is created for that service.
    The partial class is a subclass of the DrawShape class.

    The Service to partial class mapping is as follows:

    - ``com.sun.star.drawing.LineProperties`` -> ``ooodev.adapter.drawing.line_properties_partial.LinePropertiesPartial``
    - ``com.sun.star.drawing.FillProperties`` -> ``ooodev.adapter.drawing.fill_properties_partial.FillPropertiesPartial``
    - ``com.sun.star.drawing.ShadowProperties`` -> ``ooodev.adapter.drawing.shadow_properties_partial.ShadowPropertiesPartial``
    - ``com.sun.star.drawing.RotationDescriptor`` -> ``ooodev.adapter.drawing.rotation_descriptor_properties_partial.RotationDescriptorPropertiesPartial``
    - ``com.sun.star.drawing.TextProperties`` -> ``ooodev.adapter.drawing.text_properties_partial.TextPropertiesPartial``
    - ``com.sun.star.document.LinkTarget``-> ``ooodev.adapter.document.link_target_properties_partial.LinkTargetPropertiesPartial``

    If the shape supports the ``com.sun.star.drawing.GroupShape`` service, the following partial classes are implemented:

    - ``ooodev.adapter.drawing.shape_group_partial.ShapeGroupPartial``
    - ``ooodev.adapter.drawing.shapes_partial.ShapesPartial``
    - ``ooodev.adapter.drawing.rotation_descriptor_properties_partial.RotationDescriptorPropertiesPartial``
    - ``ooodev.adapter.document.link_target_properties_partial.LinkTargetPropertiesPartial``

    If the shape supports the ``XGluePointsSupplier`` interface, the following partial class is also implemented:

    - ``ooodev.adapter.drawing.glue_points_supplier_partial.GluePointsSupplierPartial``

    If the shape document is of type ``DocType.IMPRESS`` (Impress document)  and the shape supports the ``com.sun.star.presentation.Shape`` service, the following partial class is also implemented:

    - ``ooodev.adapter.presentation.shape_properties_partial.ShapePropertiesPartial``

    """

    def __init__(
        self,
        *,
        owner: Any,
        component: Any,
        class_name: str = "",
        lo_inst: LoInst | None = None,
        base_class: Type[Any] = DrawShape,
    ) -> None:
        """
        Constructor

        Args:
            owner (Any): Owner of the shape
            component (Any): UNO component that represents the shape
            class_name (str, optional): Class Name. Default to the name of the ``base_class``.
            lo_inst (LoInst | None, optional): Load Office instance. Defaults to None.
            base_class (Type[Any], optional): Base Class Parent that the shape inherits from. Defaults to DrawShape.

        Raises:
            ValueError: _description_
        """

        # in Python 3.7 and later, dict keys are ordered. This is important because the order of the bases
        if not isinstance(owner, OfficeDocumentPropPartial):
            raise ValueError("owner must be an instance of OfficeDocumentPropPartial")
        OfficeDocumentPropPartial.__init__(self, owner.office_doc)
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        self._lo_inst = lo_inst
        self._name = class_name or base_class.__name__
        self._component = component
        self._base_class = base_class

        self._existing = set(type.mro(base_class))
        self._bases_partial: Dict[Type, Any] = {}
        self._bases_interfaces: Dict[Type, Any] = {}
        self._supported = set(mInfo.Info.get_available_services(self._component))
        if "com.sun.star.drawing.LineProperties" in self._supported and LinePropertiesPartial not in self._existing:
            self._bases_partial[LinePropertiesPartial] = None
        if "com.sun.star.drawing.FillProperties" in self._supported and FillPropertiesPartial not in self._existing:
            self._bases_partial[FillPropertiesPartial] = None
        if (
            "com.sun.star.drawing.ShadowProperties" in self._supported
            and ShadowPropertiesPartial not in self._existing
        ):
            self._bases_partial[ShadowPropertiesPartial] = None
        if (
            "com.sun.star.drawing.RotationDescriptor" in self._supported
            and RotationDescriptorPropertiesPartial not in self._existing
        ):
            self._bases_partial[RotationDescriptorPropertiesPartial] = None
        if "com.sun.star.drawing.TextProperties" in self._supported and TextPropertiesPartial not in self._existing:
            self._bases_partial[TextPropertiesPartial] = None
        if "com.sun.star.document.LinkTarget" in self._supported and LinkTargetPropertiesPartial not in self._existing:
            self._bases_partial[LinkTargetPropertiesPartial] = None

        if "com.sun.star.drawing.GroupShape" in self._supported:
            if ShapeGroupPartial not in self._existing:
                self._bases_interfaces[ShapeGroupPartial] = None
            if ShapesPartial not in self._existing:
                self._bases_interfaces[ShapesPartial] = None
            if RotationDescriptorPropertiesPartial not in self._existing:
                self._bases_partial[RotationDescriptorPropertiesPartial] = None
            if RotationDescriptorPropertiesPartial not in self._existing:
                self._bases_partial[RotationDescriptorPropertiesPartial] = None

        # sourcery skip: swap-nested-ifs
        # sourcery skip: merge-nested-ifs
        if self.office_doc.DOC_TYPE == DocType.IMPRESS:
            # Even though shapes in Draw and Impress are the same, the supported services are different.
            # A shape in Draw may say it supports a presentation service but it may not.
            # When a shape is drawn on a Drawpage in the Draw program. and the the draw.get_selected_shapes() is used to get
            # the selected shapes, the shape will not have the presentation service even thought is is reported in the getAvailableServices()
            if "com.sun.star.presentation.Shape" in self._supported:
                self._bases_partial[ShapePropertiesPartial] = None
        if mLo.Lo.qi(XGluePointsSupplier, component) is not None:
            if GluePointsSupplierPartial not in self._existing:
                self._bases_interfaces[GluePointsSupplierPartial] = None

    def _generate_class(self, **kwargs) -> type:
        """
        Combines all the partial classes and the base class to create a new class.

        If there are no new partial classes, the base class is returned
        """
        bases = [self._base_class] + list(self._bases_interfaces.keys()) + list(self._bases_partial.keys())
        if len(bases) == 1 and not kwargs and self._base_class.__name__ == self._name:
            return self._base_class
        else:
            return type(self._name, tuple(bases), kwargs)

    def get_class(self, **kwargs) -> Any:
        """
        Gets the class instance of the shape.

        The returned class instance is a subclass of DrawShape and the other partial classes that are supported by the shape.

        Returns:
            Any: Shape class instance
        """
        # pylint: disable=consider-iterating-dictionary
        # pylint: disable=consider-using-dict-items
        # pylint: disable=unnecessary-dunder-call
        clazz = self._generate_class(**kwargs)
        instance = clazz(self._owner, self._component, self._lo_inst)
        for t in self._bases_interfaces.keys():
            t.__init__(instance, self._component, None)  # type: ignore
        for t in self._bases_partial.keys():
            t.__init__(instance, self._component)  # type: ignore
        return instance
