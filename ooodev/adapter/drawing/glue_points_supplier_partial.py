from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.drawing import XGluePointsSupplier

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.container.index_container_comp import IndexContainerComp


if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.drawing import GluePoint2


class GluePointsSupplierPartial:
    """Partial class for XGluePointsSupplier  interface."""

    # Does no implement any methods.
    def __init__(self, component: XGluePointsSupplier, interface: UnoInterface | None = XGluePointsSupplier) -> None:
        """
        Constructor

        Args:
            component (XGluePointsSupplier): UNO Component that implements ``com.sun.star.drawing.XGluePointsSupplier`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XGluePointsSupplier``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XGluePointsSupplier
    def get_glue_points(self) -> IndexContainerComp["GluePoint2"]:
        """
        Gets a container of GluePoint2 structs.
        """
        return IndexContainerComp(self.__component.getGluePoints())

    # endregion XGluePointsSupplier
