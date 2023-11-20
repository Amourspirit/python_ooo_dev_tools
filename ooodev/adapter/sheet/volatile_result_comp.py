from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.component_base import ComponentBase
from .result_events import ResultEvents


if TYPE_CHECKING:
    from com.sun.star.sheet import VolatileResult  # service


class VolatileResultComp(ComponentBase, ResultEvents):
    """
    Class for managing Volatile Result Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: VolatileResult) -> None:
        """
        Constructor

        Args:
            component (VolatileResult): UNO Volatile Result Component
        """
        ComponentBase.__init__(self, component)
        generic_args = self._get_generic_args()
        ResultEvents.__init__(self, trigger_args=generic_args, cb=self._on_result_events_add_remove)

    # region Lazy Listeners
    def _on_result_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addResultListener(self.events_listener_result)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.VolatileResult",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> VolatileResult:
        """Tree Data Model Component"""
        return cast("VolatileResult", self._get_component())

    # endregion Properties
