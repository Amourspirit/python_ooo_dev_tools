from __future__ import annotations
from typing import Any, TYPE_CHECKING, TypeVar, Generic

from ooodev.adapter.chart2.title_comp import TitleComp
from ooodev.loader import lo as mLo
from ooodev.office import chart2 as mChart2
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.proto.component_proto import ComponentT

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.proto.style_obj import StyleT

_T = TypeVar("_T", bound="ComponentT")


class ChartTitle(Generic[_T], LoInstPropsPartial, TitleComp, PropPartial, QiPartial, ServicePartial, StylePartial):
    """
    Class for managing Chart2 Chart Title Component.
    """

    def __init__(self, owner: _T, component: Any, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Chart2 Title Component.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        TitleComp.__init__(self, component=component)
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        StylePartial.__init__(self, component=component)
        self._owner = owner

    # region StylePartial Overrides

    def apply_styles(self, *styles: StyleT, **kwargs) -> None:
        """
        Applies style to component.

        Args:
            styles expandable list of styles object such as ``Font`` to apply to ``obj``.
            kwargs (Any, optional): Expandable list of key value pairs.

        Returns:
            None:
        """
        mChart2.Chart2._style_title(self.__component, styles)

    # endregion

    @property
    def owner(self) -> _T:
        """Chart Document"""
        return self._owner
