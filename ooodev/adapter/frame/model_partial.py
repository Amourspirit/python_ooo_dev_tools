from __future__ import annotations
from typing import TYPE_CHECKING, Tuple

import uno
from com.sun.star.frame import XModel

from ooodev.adapter.lang.component_partial import ComponentPartial

if TYPE_CHECKING:
    from com.sun.star.frame import XController
    from com.sun.star.beans import PropertyValue
    from com.sun.star.uno import XInterface
    from ooodev.utils.type_var import UnoInterface


class ModelPartial(ComponentPartial):
    """
    Partial class for XModel.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XModel, interface: UnoInterface | None = XModel) -> None:
        """
        Constructor

        Args:
            component (XModel ): UNO Component that implements ``com.sun.star.frame.XModel`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XModel``.
        """

        ComponentPartial.__init__(self, component=component, interface=interface)
        self.__component = component

    # region XModel
    def attach_resource(self, url: str, *args: PropertyValue) -> bool:
        """
        Informs a model about its resource description.
        """
        return self.__component.attachResource(url, args)

    def connect_controller(self, controller: XController) -> None:
        """
        Is called whenever a new controller is created for this model.

        The com.sun.star.lang.XComponent interface of the controller must be used to recognize when it is deleted.
        """
        self.__component.connectController(controller)

    def disconnect_controller(self, controller: XController) -> None:
        """
        is called whenever an existing controller should be deregistered at this model.

        The com.sun.star.lang.XComponent interface of the controller must be used to recognize when it is deleted.
        """
        self.__component.disconnectController(controller)

    def get_args(self) -> Tuple[PropertyValue, ...]:
        """
        Provides read access on currently representation of the com.sun.star.document.MediaDescriptor of this model which describes the model and its state
        """
        return self.__component.getArgs()

    def get_current_controller(self) -> XController:
        """
        Provides access to the controller which currently controls this model
        """
        return self.__component.getCurrentController()

    def get_current_selection(self) -> XInterface:
        """
        Provides read access on current selection on controller
        """
        return self.__component.getCurrentSelection()

    def get_url(self) -> str:
        """
        Provides information about the location of this model
        """
        return self.__component.getURL()

    def has_controllers_locked(self) -> bool:
        """
        determines if there is at least one lock remaining.

        While there is at least one lock remaining, some notifications for display updates are not broadcasted to the controllers.
        """
        return self.__component.hasControllersLocked()

    def lock_controllers(self) -> None:
        """
        suspends some notifications to the controllers which are used for display updates.

        The calls to XModel.lockControllers() and XModel.unlockControllers() may be nested and even overlapping, but they must be in pairs. While there is at least one lock remaining, some notifications for display updates are not broadcasted.
        """
        self.__component.lockControllers()

    def set_current_controller(self, controller: XController) -> None:
        """
        sets a registered controller as the current controller.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        self.__component.setCurrentController(controller)

    def unlock_controllers(self) -> None:
        """
        resumes the notifications which were suspended by XModel.lockControllers().

        The calls to XModel.lockControllers() and XModel.unlockControllers() may be nested and even overlapping, but they must be in pairs. While there is at least one lock remaining, some notifications for display updates are not broadcasted.
        """
        self.__component.unlockControllers()

    # endregion XModel
