from __future__ import annotations
from typing import Any, TYPE_CHECKING
from com.sun.star.form.runtime import XFormController

from ooodev.adapter.container.child_partial import ChildPartial
from ooodev.adapter.container.index_access_partial import IndexAccessPartial
from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial
from ooodev.adapter.lang.component_partial import ComponentPartial
from ooodev.adapter.util.modify_broadcaster_partial import ModifyBroadcasterPartial
from ooodev.adapter.form.confirm_delete_broadcaster_partial import ConfirmDeleteBroadcasterPartial
from ooodev.adapter.sdb.sql_error_broadcaster_partial import SQLErrorBroadcasterPartial
from ooodev.adapter.sdb.row_set_approve_broadcaster_partial import RowSetApproveBroadcasterPartial
from ooodev.adapter.form.database_parameter_broadcaster2_partial import DatabaseParameterBroadcaster2Partial
from ooodev.adapter.util.mode_selector_partial import ModeSelectorPartial
from ooodev.adapter.form.runtime.filter_controller_partial import FilterControllerPartial
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo


if TYPE_CHECKING:
    from com.sun.star.awt import XControl
    from com.sun.star.form import XFormControllerListener
    from com.sun.star.form.runtime import XFormControllerContext
    from com.sun.star.form.runtime import XFormOperations
    from com.sun.star.task import XInteractionHandler
    from ooodev.utils.type_var import UnoInterface


class FormControllerPartial(
    ChildPartial,
    IndexAccessPartial,
    EnumerationAccessPartial,
    ComponentPartial,
    ModifyBroadcasterPartial,
    ConfirmDeleteBroadcasterPartial,
    SQLErrorBroadcasterPartial,
    RowSetApproveBroadcasterPartial,
    DatabaseParameterBroadcaster2Partial,
    ModeSelectorPartial,
    FilterControllerPartial,
):
    """
    Partial Class for XFormController.

    This interface does not really provide an own functionality, it is only for easier runtime identification of form components.
    """

    def __init__(self, component: XFormController, interface: UnoInterface | None = XFormController) -> None:
        """
        Constructor

        Args:
            component (XFormController): UNO Component that implements ``com.sun.star.form.runtime.XFormController``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``None``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)

        ChildPartial.__init__(self, component, interface=None)
        IndexAccessPartial.__init__(self, component, interface=None)
        EnumerationAccessPartial.__init__(self, component, interface=None)
        ComponentPartial.__init__(self, component, interface=None)
        ModifyBroadcasterPartial.__init__(self, component, interface=None)
        ConfirmDeleteBroadcasterPartial.__init__(self, component, interface=None)
        SQLErrorBroadcasterPartial.__init__(self, component, interface=None)
        RowSetApproveBroadcasterPartial.__init__(self, component, interface=None)
        DatabaseParameterBroadcaster2Partial.__init__(self, component, interface=None)
        ModeSelectorPartial.__init__(self, component, interface=None)
        FilterControllerPartial.__init__(self, component, interface=None)
        self.__component = component

    # region XFormController
    def add_activate_listener(self, listener: XFormControllerListener) -> None:
        """
        Adds the specified listener to receive notifications whenever the activation state of the controller changes.
        """
        self.__component.addActivateListener(listener)

    def add_child_controller(self, child_controller: XFormController) -> None:
        """
        adds a controller to the list of child controllers

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        self.__component.addChildController(child_controller)

    def remove_activate_listener(self, listener: XFormControllerListener) -> None:
        """
        removes the specified listener from the list of components to receive notifications whenever the activation state of the controller changes.
        """
        self.__component.removeActivateListener(listener)

    @property
    def context(self) -> XFormControllerContext:
        """
        Gets/Sets - allows to delegate certain tasks to the context of the form controller
        """
        return self.__component.Context

    @context.setter
    def context(self, value: XFormControllerContext) -> None:
        self.__component.Context = value

    @property
    def current_control(self) -> XControl:
        """
        Gets access to the currently active control.
        """
        return self.__component.CurrentControl

    @property
    def form_operations(self) -> XFormOperations:
        """
        Gets the instance which is used to implement operations on the form which the controller works for.

        This instance can be used, for instance, to determine the current state of certain form features.
        """
        return self.__component.FormOperations

    @property
    def interaction_handler(self) -> XInteractionHandler:
        """
        Gets/Sets - used (if not ``None``) for user interactions triggered by the form controller.
        """
        return self.__component.InteractionHandler

    @interaction_handler.setter
    def interaction_handler(self, value: XInteractionHandler) -> None:
        self.__component.InteractionHandler = value

    # endregion XFormController
