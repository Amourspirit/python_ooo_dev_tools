from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.frame import XController

from ooodev.adapter.lang.component_partial import ComponentPartial

if TYPE_CHECKING:
    from com.sun.star.frame import XFrame
    from com.sun.star.frame import XModel
    from ooodev.utils.type_var import UnoInterface


class ControllerPartial(ComponentPartial):
    """
    Partial class for XController.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XController, interface: UnoInterface | None = XController) -> None:
        """
        Constructor

        Args:
            component (XController ): UNO Component that implements ``com.sun.star.frame.XController`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XController``.
        """

        ComponentPartial.__init__(self, component=component, interface=interface)
        self.__component = component

    # region XController
    def attach_frame(self, frame: XFrame) -> None:
        """
        Attaches the controller with its managing frame.

        Args:
            frame (XFrame): The frame to be attached.
        """
        self.__component.attachFrame(frame)

    def attach_model(self, model: XModel) -> None:
        """
        Attaches the controller to a new model.

        Args:
            model (XModel): The model to be attached.
        """
        self.__component.attachModel(model)

    def get_frame(self) -> XFrame:
        """
        Gets access to owner frame of this controller.

        Returns:
            XFrame: The owner frame of this controller.
        """
        return self.__component.getFrame()

    def get_model(self) -> XModel:
        """
        Gets access to currently attached model.

        Returns:
            XModel: The currently attached model.
        """
        return self.__component.getModel()

    def get_view_data(self) -> Any:
        """
        Gets the view data.

        Returns:
            Any: The view data.
        """
        return self.__component.getViewData()

    def restore_view_data(self, data: Any) -> None:
        """
        Restores the view data.

        Args:
            data (Any): The view data to be restored.
        """
        self.__component.restoreViewData(data)

    def suspend(self, suspend: bool) -> bool:
        """
        Is called to prepare the controller for closing the view.

        Args:
            suspend (bool): ``True`` Force the controller to suspend his work, ``False`` Try to reactivate the controller.

        Returns:
            bool: ``True`` If request was accepted and successfully finished; Otherwise, ``False``.
        """
        return self.__component.suspend(suspend)

    # endregion XController
