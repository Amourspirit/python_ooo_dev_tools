from __future__ import annotations
from typing import Any
import uno
from ooodev.adapter.lang.type_provider_partial import TypeProviderPartial
from ooodev.adapter.lang.service_info_partial import ServiceInfoPartial
from ooodev.adapter.uno.weak_partial import WeakPartial
from ooodev.utils.partial.interface_partial import InterfacePartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.events.partial.events_partial import EventsPartial


class CompDefaultsPartial(
    WeakPartial, TypeProviderPartial, ServiceInfoPartial, InterfacePartial, QiPartial, EventsPartial
):
    """
    Class for managing default imports for the builder.

    Note:
        This class is used to provide default imports Types.
        This class is a internal class and is used for type hinting.

        The classes inherited by this class should match the classes imported by ``builder_helper.builder_add_comp_defaults`.
    """

    def __init__(self, component: Any, interface: Any | None = None) -> None:
        WeakPartial.__init__(self, component=component, interface=interface)
        TypeProviderPartial.__init__(self, component=component, interface=interface)
        ServiceInfoPartial.__init__(self, component=component, interface=interface)
        InterfacePartial.__init__(self)
        QiPartial.__init__(self, component=component)
        EventsPartial.__init__(self)
