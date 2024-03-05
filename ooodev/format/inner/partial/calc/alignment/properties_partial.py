from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler
from ooodev.format.inner.style_factory import calc_align_properties_factory
from ooodev.events.partial.events_partial import EventsPartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.inner.direct.calc.alignment.properties import TextDirectionKind

    from ooodev.format.proto.calc.alignment.properties_t import PropertiesT
else:
    TextAlignT = Any
    LoInst = Any
    TextDirectionKind = Any

# ooodev.format.inner.direct.calc.alignment.properties.Properties


class PropertiesPartial:
    """
    Partial class for Calc Cell/Range Properties.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_align_properties",
            after_event="after_style_align_properties",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

    def style_align_properties(
        self,
        *,
        wrap_auto: bool | None = None,
        hyphen_active: bool | None = None,
        shrink_to_fit: bool | None = None,
        direction: TextDirectionKind | None = None,
    ) -> PropertiesT | None:
        """
        Style Alignment Properties.

        Args:
            hori_align (HoriAlignKind, optional): Specifies Horizontal Alignment.
            indent: (float, UnitT, optional): Specifies indent in ``pt`` (point) units or :ref:`proto_unit_obj`.
                Only used when ``hori_align`` is set to ``HoriAlignKind.LEFT``
            vert_align (VertAdjustKind, optional): Specifies Vertical Alignment.

        Raises:
            CancelEventError: If the event ``before_style_align_properties`` is cancelled and not handled.

        Returns:
            PropertiesT | None: Text Alignment instance or ``None`` if cancelled.

        Hint:
            - ``TextDirectionKind`` can be imported from ``ooodev.format.inner.direct.calc.alignment.properties``
        """
        factory = calc_align_properties_factory
        kwargs = {
            "wrap_auto": wrap_auto,
            "hyphen_active": hyphen_active,
            "shrink_to_fit": shrink_to_fit,
            "direction": direction,
        }
        return self.__styler.style(factory=factory, **kwargs)

    def style_align_properties_get(self) -> PropertiesT | None:
        """
        Gets the Alignment Properties Style.

        Raises:
            CancelEventError: If the event ``before_style_align_properties_get`` is cancelled and not handled.

        Returns:
            PropertiesT | None: Text Alignment style or ``None`` if cancelled.
        """
        return self.__styler.style_get(factory=calc_align_properties_factory)
