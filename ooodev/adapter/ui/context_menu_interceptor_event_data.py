from __future__ import annotations
from dataclasses import dataclass
from ooo.dyn.ui.context_menu_interceptor_action import ContextMenuInterceptorAction
from ooodev.adapter.ui.context_menu_execute_event_comp import ContextMenuExecuteEventComp


@dataclass(frozen=False)
class ContextMenuInterceptorEventData:
    """
    Event Data for ContextMenuExecuteEvent

    See Also:
        - :py:class:`ooodev.adapter.ui.context_menu_interceptor.ContextMenuInterceptor`
        - :py:class:`ooodev.adapter.ui.context_menu_execute_event_comp.ContextMenuExecuteEventComp`
    """

    event: ContextMenuExecuteEventComp
    """Event Data"""
    action: ContextMenuInterceptorAction
    """
    Action to take

    Hint:
        - ``ContextMenuInterceptorAction`` is an enum that can be imported from ``ooo.dyn.ui.context_menu_interceptor_action``.
    """
