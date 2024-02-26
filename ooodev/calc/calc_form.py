from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from ooodev.adapter.form.component.data_form_comp import DataFormComp
from ooodev.form.partial.form_partial import FormPartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial

if TYPE_CHECKING:
    from com.sun.star.form.component import Form
    from ooodev.calc.calc_forms import CalcForms


class CalcForm(LoInstPropsPartial, DataFormComp, QiPartial, FormPartial, ServicePartial, CalcDocPropPartial):
    """
    Calc From. Represents a form in a Calc document.

    This class is Enumerable.

    ``len(calc_form)`` returns the number of controls in the form.
    """

    def __init__(self, owner: CalcForms, component: Form, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        DataFormComp.__init__(self, component)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        draw_page = owner.owner.component
        FormPartial.__init__(self, owner=self, draw_page=draw_page, component=component, lo_inst=self.lo_inst)  # type: ignore
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        CalcDocPropPartial.__init__(self, obj=owner.calc_doc)

    def __getitem__(self, index: str | int) -> Any:
        if isinstance(index, int):
            return self.get_by_index(index)
        return self.get_by_name(index)

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
    def owner(self) -> CalcForms:
        """Component Owner"""
        return self._owner

    # endregion Properties
