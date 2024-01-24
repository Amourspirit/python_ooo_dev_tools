from __future__ import annotations
from typing import overload, TYPE_CHECKING
import uno
from com.sun.star.form import XForm

from ooodev.adapter.form.forms_comp import FormsComp
from ooodev.exceptions import ex as mEx
from ooodev.utils import gen_util as mGenUtil
from ooodev.utils import lo as mLo
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.inst.lo.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.qi_partial import QiPartial
from .write_form import WriteForm

if TYPE_CHECKING:
    from com.sun.star.form import XForms
    from .write_draw_page import WriteDrawPage


class WriteForms(LoInstPropsPartial, FormsComp, QiPartial):
    """
    Class for managing Writer Forms.

    This class is Enumerable and returns ``WriteForm`` instance on iteration.

    .. code-block:: python

        for sheet in doc.sheets:
            sheet["A1"].set_val("test")
            assert sheet["A1"].get_val() == "test"

    This class also as index access and returns ``CalcSheet`` instance.

    .. code-block:: python

        sheet = doc.sheets["Sheet1"]
        # or set the value of cell A2 to TEST
        doc.sheets[0]["A2"].set_val("TEST")

        # get the last sheet of the document
        last_sheet = doc.sheets[-1]

        # get the second last sheet of the document
        second_last_sheet = doc.sheets[-2]

        # get the number of sheets
        num_sheets = len(doc.sheets)

    .. versionadded:: 0.18.3
    """

    def __init__(self, owner: WriteDrawPage, forms: XForms, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (WriteDrawPage): Owner Component
            forms (XForms): Forms instance.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        FormsComp.__init__(self, forms)  # type: ignore
        QiPartial.__init__(self, component=forms, lo_inst=self.lo_inst)

    def __next__(self) -> WriteForm:
        return WriteForm(owner=self, component=super().__next__(), lo_inst=self.lo_inst)

    def __getitem__(self, index: str | int) -> WriteForm:
        if isinstance(index, int):
            return self.get_by_index(index)
        return self.get_by_name(index)

    def __len__(self) -> int:
        return self.component.getCount()

    def _get_index(self, idx: int, allow_greater: bool = False) -> int:
        """
        Gets the index.

        Args:
            idx (int): Index of sheet. Can be a negative value to index from the end of the list.
            allow_greater (bool, optional): If True and index is greater then the number of
                sheets then the index becomes the next index if sheet were appended. Defaults to False.

        Returns:
            int: Index value.
        """
        count = len(self)
        return mGenUtil.Util.get_index(idx, count, allow_greater)

    def _create_name(self, name: str) -> str:
        used_name = True
        i = 1
        nm = f"{name}{i}"
        while used_name:
            used_name = self.has_by_name(nm)
            if used_name:
                i += 1
                nm = f"{name}{i}"
        return nm

    # region add_form
    @overload
    def add_form(self) -> WriteForm:
        """
        Adds a new form at the end.

        Returns:
            WriteForm: Form
        """
        ...

    @overload
    def add_form(self, idx: int) -> WriteForm:
        """
        Adds a new form.

        Args:
            idx (int): Index of form.

        Returns:
            WriteForm: Form
        """
        ...

    @overload
    def add_form(self, name: str) -> WriteForm:
        """
        Adds a new form.

        Args:
            name (str): Name of form.

        Returns:
            WriteForm: Form
        """
        ...

    def add_form(self, *args, **kwargs) -> WriteForm:
        """
        Adds a new form.

        Args:
            name (str): Name of form.
            idx (int): Index of form.

        Raises:
            NameClashError: If name already exists.

        Returns:
            WriteForm: Form
        """
        all_args = [arg for arg in args]
        all_args.extend(kwargs.values())
        if len(all_args) == 0:
            # add a new form at the end
            all_args.append(-1)

        if len(all_args) != 1:
            raise TypeError("add_form() takes 1 argument but {} were given".format(len(all_args)))
        arg1 = all_args[0]
        if isinstance(arg1, int):
            idx = self._get_index(arg1, allow_greater=True)
            with LoContext(self.lo_inst):
                frm = mLo.Lo.create_instance_mcf(XForm, "stardiv.one.form.component.Form", raise_err=True)
            frm.Name = self._create_name("Form")  # type: ignore
            self.insert_by_index(idx, frm)
            return self.get_by_index(idx)
        elif isinstance(arg1, str):
            if self.has_by_name(arg1):
                raise mEx.NameClashError(f"Name '{arg1}' already exists")
            with LoContext(self.lo_inst):
                frm = mLo.Lo.create_instance_mcf(XForm, "stardiv.one.form.component.Form", raise_err=True)
            self.insert_by_name(arg1, frm)
            return self.get_by_name(arg1)
        else:
            raise TypeError("add_form() argument must be int or str")

    # endregion add_form

    # region XIndexAccess overrides

    def get_by_index(self, idx: int) -> WriteForm:
        """
        Gets the element at the specified index.

        Args:
            idx (int): The Zero-based index of the element. Idx can be a negative value to index from the end of the list.
                For example, -1 will return the last element.

        Returns:
            WriteForm: The element at the specified index.
        """
        idx = self._get_index(idx, True)
        result = super().get_by_index(idx)
        return WriteForm(owner=self, component=result, lo_inst=self.lo_inst)

    # endregion XIndexAccess overrides

    # region XNameAccess overrides

    def get_by_name(self, name: str) -> WriteForm:
        """
        Gets the element with the specified name.

        Args:
            name (str): The name of the element.

        Raises:
            MissingNameError: If sheet is not found.

        Returns:
            WriteForm: The element with the specified name.
        """
        if not self.has_by_name(name):
            raise mEx.MissingNameError(f"Unable to find sheet with name '{name}'")
        result = super().get_by_name(name)
        return WriteForm(owner=self, component=result, lo_inst=self.lo_inst)

    # endregion XNameAccess overrides

    # region Properties
    @property
    def owner(self) -> WriteDrawPage:
        """
        Returns:
            WriteDrawPage: Writer Draw Page
        """
        return self._owner

    # endregion Properties
