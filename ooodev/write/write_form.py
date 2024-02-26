from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from ooodev.adapter.form.component.data_form_comp import DataFormComp
from ooodev.form.partial.form_partial import FormPartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial

if TYPE_CHECKING:
    from com.sun.star.form.component import Form
    from ooodev.write.write_forms import WriteForms


class WriteForm(LoInstPropsPartial, DataFormComp, WriteDocPropPartial, QiPartial, FormPartial):
    """Writer Form"""

    def __init__(self, owner: WriteForms, component: Form, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (WriteForms): Owner of this component.
            component (Form): UNO object that supports ``com.sun.star.form.component.DataForm`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)
        self._owner = owner
        DataFormComp.__init__(self, component)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        draw_page = owner.owner.component
        FormPartial.__init__(self, owner=self, draw_page=draw_page, component=component, lo_inst=self.lo_inst)  # type: ignore

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
    def owner(self) -> WriteForms:
        """Component Owner"""
        return self._owner

    # endregion Properties
