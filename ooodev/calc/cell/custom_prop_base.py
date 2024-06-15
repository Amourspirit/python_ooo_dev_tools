from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from com.sun.star.drawing import XControlShape
from com.sun.star.form import XForm
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial

if TYPE_CHECKING:
    from com.sun.star.form.component import Form
    from com.sun.star.form.component import HiddenControl
    from ooodev.calc.calc_sheet import CalcSheet
    from ooodev.calc.spreadsheet_draw_page import SpreadsheetDrawPage


class CustomPropBase(TheDictionaryPartial):
    def __init__(self, sheet: CalcSheet) -> None:
        TheDictionaryPartial.__init__(self)
        self._shape_prefix = "_cprop_"
        self._shape_suffix = "_id"  # suffix is important for ensure shape duplicates are removed.
        self._sheet = sheet
        self._form_name = "CellCustomProperties"
        self._cache = {}
        self._draw_page = self._sheet.draw_page

    def _get_hidden_control_simple(self, name: str) -> HiddenControl | None:
        frm = self._get_form()
        if not frm.hasByName(name):
            return None
        return frm.getByName(name)

    def _get_hidden_control_name_from_shape(self, shape: XControlShape) -> str:
        # Name is in format of _cprop_idofhsvtcky1hgom_id
        name = shape.Name  # type: ignore
        prefix_len = len(self.shape_prefix)
        suffix_len = len(self.shape_suffix)
        name = name[prefix_len:]
        name = name[:-suffix_len]
        return name

    def _get_form(self) -> Form:

        key = self._form_name
        if key in self._cache:
            return self._cache[key]
        forms = self.draw_page.forms.component
        if len(forms) == 0:  # type: ignore
            # insert a default form1.
            # The reason for this is many users many working in forms[0].
            # This way there will be a from to work with that is not for properties.
            # This is not critical but it is a good practice.
            # If the user deletes Forms[0] it will not wipe the property forms.
            # Also if the user draws control on the spreadsheet or other document it will use this form.
            frm = cast(
                "Form",
                self.sheet.lo_inst.create_instance_mcf(XForm, "stardiv.one.form.component.Form", raise_err=True),
            )
            frm.Name = "Form1"
            forms.insertByName("Form1", frm)

        if forms.hasByName(key):
            frm = forms.getByName(key)
        else:
            frm = cast(
                "Form",
                self.sheet.lo_inst.create_instance_mcf(XForm, "stardiv.one.form.component.Form", raise_err=True),
            )
            frm.Name = key
            forms.insertByName(key, frm)
        self._cache[key] = frm
        return frm

    def _reset(self) -> None:
        self._cache.clear()

    def __del__(self) -> None:
        pass

    # region Properties
    @property
    def cache(self) -> dict:
        return self._cache

    @property
    def draw_page(self) -> SpreadsheetDrawPage:
        return self._draw_page

    @property
    def form_name(self) -> str:
        return self._form_name

    @property
    def shape_prefix(self) -> str:
        return self._shape_prefix

    @property
    def shape_suffix(self) -> str:
        return self._shape_suffix

    @property
    def sheet(self) -> CalcSheet:
        return self._sheet

    # endregion Properties
