from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.view import XSelectionSupplier

from ooodev.adapter.component_base import ComponentBase
from ooodev.utils import info as mInfo
from ooodev.utils import gui as mGui
from ooodev.loader import lo as mLo
from ooodev.utils.partial.gui_partial import GuiPartial
from ooodev.adapter.drawing.shapes_partial import ShapesPartial

if TYPE_CHECKING:
    from com.sun.star.drawing import ShapeCollection  # service
    from com.sun.star.frame import XController


class ShapeCollectionComp(
    ComponentBase,
    ShapesPartial,
):
    """
    Class for managing ShapeCollection Component.

    .. versionadded:: 0.20.5
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO component that implements ``com.sun.star.drawing.ShapeCollection`` service.
        """
        ComponentBase.__init__(self, component)
        ShapesPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.ShapeCollection",)

    # endregion Overrides

    # region Methods
    def select_all(self, obj: Any) -> None:
        """
        Selects all shapes in the collection.

        Args:
            obj (Any): The object to select shapes from.
                Can be ``XController`` or ``XModel``  or ``DrawDoc``, ``ImpressDoc`` ( or any object that implements ``GuiPartial``)

        Returns:
            None:
        """
        if mLo.Lo.is_uno_interfaces(obj, "com.sun.star.frame.XController"):
            controller = cast("XController", obj)
        elif mInfo.Info.is_instance(obj, GuiPartial):
            controller = obj.get_current_controller()
        else:
            controller = mGui.GUI.get_current_controller(obj)
        select_supp = mLo.Lo.qi(XSelectionSupplier, controller, True)
        select_supp.select(self.component)

    # endregion Methods

    # region Properties
    @property
    def component(self) -> ShapeCollection:
        """ShapeCollection Component"""
        # pylint: disable=no-member
        return cast("ShapeCollection", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
