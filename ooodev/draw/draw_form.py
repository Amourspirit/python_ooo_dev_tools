from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from ooodev.adapter.form.component.data_form_comp import DataFormComp
from ooodev.form.partial.form_partial import FormPartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial

if TYPE_CHECKING:
    from com.sun.star.form.component import Form
    from ooodev.draw.draw_forms import DrawForms


class DrawForm(LoInstPropsPartial, DataFormComp, QiPartial, FormPartial, ServicePartial):
    """Draw Form class"""

    def __init__(self, owner: DrawForms, component: Form, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        DataFormComp.__init__(self, component)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        draw_page = owner.owner.component
        FormPartial.__init__(self, owner=self, draw_page=draw_page, component=component, lo_inst=self.lo_inst)  # type: ignore
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)

    # region Properties
    @property
    def name(self) -> str:
        """
        Gets/Sets the name of the form.
        """
        return self.component.getName()

    @name.setter
    def name(self, value: str) -> None:
        self.component.setName(value)

    @property
    def owner(self) -> DrawForms:
        """Component Owner"""
        return self._owner

    # endregion Properties
