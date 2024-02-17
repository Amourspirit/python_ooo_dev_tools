from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from ooodev.events.partial.events_partial import EventsPartial
from ooodev.loader import lo as mLo
from ooodev.format.inner.partial.factory_styler import FactoryStyler
from ooodev.format.inner.style_factory import calc_align_text_factory

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.inner.direct.calc.alignment.text_align import VertAlignKind
    from ooodev.format.inner.direct.calc.alignment.text_align import HoriAlignKind

    from ooodev.format.proto.calc.alignment.text_align_t import TextAlignT
    from ooodev.units import UnitT
else:
    TextAlignT = Any
    LoInst = Any
    VertAlignKind = Any
    HoriAlignKind = Any
    UnitT = Any


class TextAlignPartial:
    """
    Partial class for Calc Text Align.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__styler = FactoryStyler(factory_name=factory_name, component=component, lo_inst=lo_inst)
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)
        self.__styler.after_event_name = "after_style_align_text"
        self.__styler.before_event_name = "before_style_align_text"

    def style_align_text(
        self,
        hori_align: HoriAlignKind | None = None,
        indent: float | UnitT | None = None,
        vert_align: VertAlignKind | None = None,
    ) -> TextAlignT | None:
        """
        Style Alignment Text.

        Args:
            hori_align (HoriAlignKind, optional): Specifies Horizontal Alignment.
            indent: (float, UnitT, optional): Specifies indent in ``pt`` (point) units or :ref:`proto_unit_obj`.
                Only used when ``hori_align`` is set to ``HoriAlignKind.LEFT``
            vert_align (VertAdjustKind, optional): Specifies Vertical Alignment.

        Raises:
            CancelEventError: If the event ``before_style_align_text`` is cancelled and not handled.

        Returns:
            TextAlignT | None: Text Alignment instance or ``None`` if cancelled.

        Hint:
            - ``HoriAlignKind`` can be imported from ``ooodev.format.inner.direct.calc.alignment.text_align``
            - ``VertAlignKind`` can be imported from ``ooodev.format.inner.direct.calc.alignment.text_align``
        """
        factory = calc_align_text_factory
        kwargs = {"hori_align": hori_align, "indent": indent, "vert_align": vert_align}
        return self.__styler.style(factory=factory, **kwargs)

    def style_align_text_get(self) -> TextAlignT | None:
        """
        Gets the Alignment Text Style.

        Raises:
            CancelEventError: If the event ``before_style_align_text_get`` is cancelled and not handled.

        Returns:
            TextAlignT | None: Text Alignment style or ``None`` if cancelled.
        """
        return self.__styler.style_get(factory=calc_align_text_factory)
