from __future__ import annotations
from typing import Protocol
from ooodev.adapter.beans.property_set_info_t import PropertySetInfoT
from ooodev.adapter.beans.fast_property_set_t import FastPropertySetT
from ooodev.adapter.beans.multi_property_set_t import MultiPropertySetT
from ooodev.adapter.lang.service_info_t import ServiceInfoT
from ooodev.adapter.lang.type_provider_t import TypeProviderT


class ActionT(PropertySetInfoT, FastPropertySetT, MultiPropertySetT, ServiceInfoT, TypeProviderT, Protocol):
    """Type for context Menu item."""

    pass
