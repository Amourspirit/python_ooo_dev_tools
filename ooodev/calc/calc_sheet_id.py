from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.form.controls.form_ctl_hidden import FormCtlHidden
from ooodev.utils import gen_util as gUtil

if TYPE_CHECKING:
    from ooodev.calc.calc_sheet import CalcSheet


class CalcSheetId:
    """
    Generates a unique id for the sheet if it does not exist and makes it available via the id property.
    """

    def __init__(self, sheet: CalcSheet) -> None:
        self._sheet = sheet
        self._hidden_name = "SheetUniqueID"
        self._id = ""

    def _get_sheet_id(self) -> str:
        # need to get a unique id for the sheet.

        if len(self._sheet.draw_page.forms) == 0:
            frm = self._sheet.draw_page.forms.add_form("Form1")
        else:
            frm = self._sheet.draw_page.forms[0]
        if frm.has_by_name(self._hidden_name):
            ctl = FormCtlHidden(frm.get_by_name(self._hidden_name), self._sheet.lo_inst)
            return ctl.hidden_value
            # sheet_id = cast(str, ctl.get_property("sheet_id"))
        str_id = gUtil.Util.generate_random_string(14).lower()
        ctl = frm.insert_control_hidden(name=self._hidden_name)
        ctl.hidden_value = str_id
        # extra properties can be added to the control if needed
        # ctl.add_property("custom_data", PropertyAttributeEnum.CONSTRAINED, "custom_data_value")
        return str_id

    @property
    def id(self) -> str:
        if not self._id:
            self._id = self._get_sheet_id()
        return self._id
