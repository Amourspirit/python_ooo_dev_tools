from __future__ import annotations

# pylint: disable=invalid-name, unused-import
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.ui import XContextMenuInterceptor
from ooo.dyn.ui.context_menu_interceptor_action import ContextMenuInterceptorAction
from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase
from ooodev.events.args.event_args_generic import EventArgsGeneric
from ooodev.adapter.ui.context_menu_interceptor_event_data import ContextMenuInterceptorEventData
from ooodev.adapter.ui.context_menu_execute_event_comp import ContextMenuExecuteEventComp
from ooodev.utils import info as mInfo

# this is a listener class but does not have a listener suffix
# https://wiki.documentfoundation.org/Framework/Tutorial/Context_Menu_Interception


if TYPE_CHECKING:
    from com.sun.star.ui import ContextMenuExecuteEvent


class ContextMenuInterceptor(AdapterBase, XContextMenuInterceptor):
    """
    This interface enables the object to be registered as interceptor to change context menus or prevent them from being executed.

    See Also:
        `API XContextMenuInterceptor <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1ui_1_1XContextMenuInterceptor.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        super().__init__(trigger_args=trigger_args)

    # region XContextMenuInterceptor

    def notifyContextMenuExecute(self, event: ContextMenuExecuteEvent) -> ContextMenuInterceptorAction:
        """
        notifies the interceptor about the request to execute a ContextMenu.

        The interceptor has to decide whether the menu should be executed with or without being modified or may ignore the call.
        """
        event_data = ContextMenuInterceptorEventData(
            event=ContextMenuExecuteEventComp(event), action=ContextMenuInterceptorAction.IGNORED
        )
        event_args = EventArgsGeneric(source=self, event_data=event_data)
        self._trigger_direct_event("notifyContextMenuExecute", event_args)  # type: ignore
        return event_args.event_data.action

    def is_menu_entry(self, element: Any) -> bool:
        """
        Gets if whether the element is a menu element.
        """
        if element is None:
            return False
        return mInfo.Info.support_service(element, "com.sun.star.ui.ActionTrigger")

    def is_menu_separator(self, element: Any) -> bool:
        """
        Gets if whether the element is a menu separator.
        """
        if element is None:
            return False
        return mInfo.Info.support_service(element, "com.sun.star.ui.ActionTriggerSeparator")

    # endregion XContextMenuInterceptor
