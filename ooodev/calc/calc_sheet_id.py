from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from com.sun.star.form import XForm

from ooodev.form.controls.form_ctl_hidden import FormCtlHidden
from ooodev.utils import gen_util as gUtil
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial

if TYPE_CHECKING:
    from com.sun.star.form.component import Form
    from ooodev.calc.calc_sheet import CalcSheet
    from ooodev.calc.calc_form import CalcForm


class CalcSheetId(TheDictionaryPartial):
    """
    Generates a unique id for the sheet if it does not exist and makes it available via the id property.

    Note:
        This class creates a hidden control on the sheet to store the unique id.
        When the sheet is copied the hidden control is not copied.
        This means that for copied and new sheets a new unique id will be generated.
    """

    def __init__(self, sheet: CalcSheet) -> None:
        TheDictionaryPartial.__init__(self)
        self._sheet = sheet
        self._hidden_name = "SheetUniqueID"
        self._id = ""
        self._form_id = "Form_SheetCustomProperties"

    def _add_default_form(self) -> None:

        if len(self._sheet.draw_page.forms) == 0:
            self._sheet.draw_page.forms.add_form("Form1")
        frm = self._sheet.draw_page.forms[0]
        if not frm.has_by_name(self._hidden_name):
            return

    def _0_45_0_upgrade(self) -> str:
        # upgrade to version 0.45.0
        if self._sheet.lo_inst.version >= (0, 45, 0):
            return ""
        if len(self._sheet.draw_page.forms) == 0:
            return ""
        frm = self._sheet.draw_page.forms[0]
        if not frm.has_by_name(self._hidden_name):
            return ""
        ctl = FormCtlHidden(frm.get_by_name(self._hidden_name), self._sheet.lo_inst)

        unique_id = ctl.hidden_value
        if not unique_id:
            frm.remove_by_name(self._hidden_name)
            return ""

        if not self._sheet.draw_page.forms.has_by_name("Form1"):
            frm = cast(
                "Form",
                self._sheet.lo_inst.create_instance_mcf(XForm, "stardiv.one.form.component.Form", raise_err=True),
            )
            frm.Name = "Form1"
            self._sheet.draw_page.forms.insert_by_index(0, frm)

        return unique_id

    def _get_hidden_control_form(self) -> CalcForm:
        key = self._form_id
        if self._sheet.draw_page.forms.has_by_name(key):
            return self._sheet.draw_page.forms.get_by_name(key)
        return self._sheet.draw_page.forms.add_form(key)

    def _get_sheet_id(self) -> str:
        # need to get a unique id for the sheet.

        str_id = self._0_45_0_upgrade()
        self._add_default_form()

        if self._sheet.draw_page.forms.has_by_name(self._form_id):
            frm = self._sheet.draw_page.forms.get_by_name(self._form_id)
        else:
            frm = self._sheet.draw_page.forms.add_form(self._form_id)

        if frm.has_by_name(self._hidden_name):
            ctl = FormCtlHidden(frm.get_by_name(self._hidden_name), self._sheet.lo_inst)
            return ctl.hidden_value
            # sheet_id = cast(str, ctl.get_property("sheet_id"))
        if not str_id:
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
