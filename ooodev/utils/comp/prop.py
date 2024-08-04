from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, TypeVar, Generic
import uno
from com.sun.star.beans import XPropertySet
from ooodev.adapter.beans.property_set_comp import PropertySetComp
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.beans import PropertySet  # service
    from ooodev.loader.inst.lo_inst import LoInst

_T = TypeVar("_T")


class Prop(Generic[_T], LoInstPropsPartial, PropPartial, PropertySetComp):
    """
    Class for managing PropertySet Component.
    """

    def __init__(self, owner: _T, component: Any, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            component (ChartType): UNO Chart2 ChartType Component.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        # validate that component is a PropertySet
        _ = lo_inst.qi(XPropertySet, component, True)
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        PropertySetComp.__init__(self, component)
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        self._owner = owner

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        # property set many not support service names.
        return ()

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> PropertySet:
        """PropertySet Component"""
        return cast("PropertySet", self._ComponentBase__get_component())  # type: ignore

    @property
    def owner(self) -> _T:
        """Owner of PropertySet"""
        return self._owner

    # endregion Properties
