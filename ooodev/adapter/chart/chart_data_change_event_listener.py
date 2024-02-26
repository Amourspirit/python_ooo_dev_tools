from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.chart import XChartDataChangeEventListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.chart import ChartDataChangeEvent
    from com.sun.star.chart import XChartData


class ChartDataChangeEventListener(AdapterBase, XChartDataChangeEventListener):
    """
    Makes it possible to receive events when chart data changes.

    See Also:
        `API XChartDataChangeEventListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart_1_1XChartDataChangeEventListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: XChartData | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            subscriber (XChartData, optional): An UNO object that implements the ``XChartData`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            subscriber.addChartDataChangeEventListener(self)

    def chartDataChanged(self, event: ChartDataChangeEvent) -> None:
        """
        Event is invoked when chart data changes in value or structure.

        This interface must be implemented by components that wish to get notified of changes in chart data.
        They can be registered at an ``XChartData`` component.
        """
        self._trigger_event("chartDataChanged", event)

    def disposing(self, event: EventObject) -> None:
        """
        Gets called when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including ``XComponent.removeEventListener()`` ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at ``XComponent``.
        """
        self._trigger_event("disposing", event)
