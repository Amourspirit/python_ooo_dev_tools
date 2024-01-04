from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno

from ooodev.adapter.form.component.data_form_comp import DataFormComp
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils import lo as mLo
from ooodev.form.partial.form_partial import FormPartial

if TYPE_CHECKING:
    from com.sun.star.form.component import Form
    from .calc_forms import CalcForms


class CalcForm(DataFormComp, QiPartial, FormPartial):
    def __init__(self, owner: CalcForms, component: Form) -> None:
        self.__owner = owner
        DataFormComp.__init__(self, component)
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)
        draw_page = owner.owner.component
        FormPartial.__init__(self, owner=self, draw_page=draw_page, component=component)

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
        return self.__owner

    # endregion Properties
