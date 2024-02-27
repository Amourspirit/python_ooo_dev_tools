from __future__ import annotations
from typing import overload, TYPE_CHECKING
import uno
from com.sun.star.form import XForm

from ooodev.adapter.form.forms_comp import FormsComp
from ooodev.exceptions import ex as mEx
from ooodev.utils import gen_util as mGenUtil
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write.write_form import WriteForm

if TYPE_CHECKING:
    from com.sun.star.form import XForms
    from ooodev.write.write_draw_page import WriteDrawPage


class WriteForms(LoInstPropsPartial, FormsComp, WriteDocPropPartial, QiPartial):
    """
    Class for managing Writer Forms.

    This class is Enumerable and returns ``WriteForm`` instance on iteration.
    """

    def __init__(self, owner: WriteDrawPage, forms: XForms, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (WriteDrawPage): Owner Component
            forms (XForms): Forms instance.
            lo_inst (LoInst, optional): Lo instance. Used when creating multiple documents. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        FormsComp.__init__(self, forms)  # type: ignore
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)
        QiPartial.__init__(self, component=forms, lo_inst=self.lo_inst)

    def __next__(self) -> WriteForm:
        """
        Gets the next form.

        Returns:
            WriteForm: The next form.
        """
        return WriteForm(owner=self, component=super().__next__(), lo_inst=self.lo_inst)

    def __getitem__(self, key: str | int) -> WriteForm:
        """
        Gets the form at the specified index or name.

        This is short hand for ``get_by_index()`` or ``get_by_name()``.

        Args:
            key (key, str, int): The index or name of the form. When getting by index can be a negative value to get from the end.

        Returns:
            WriteForm: The form with the specified index or name.

        See Also:
            - :py:meth:`~ooodev.write.WriteForms.get_by_index`
            - :py:meth:`~ooodev.write.WriteForms.get_by_name`
        """
        if isinstance(key, int):
            return self.get_by_index(key)
        return self.get_by_name(key)

    def __len__(self) -> int:
        """
        Gets the number of forms in the document.

        Returns:
            int: Number of forms in the document.
        """
        return self.component.getCount()

    def __delitem__(self, _item: int | str) -> None:
        """
        Removes a form from the document.

        Args:
            _item (int | str): Index, or name, of the form.

        Raises:
            TypeError: If the item is not a supported type.
        """
        # Delete slide by index, name, or object
        if isinstance(_item, int):
            self.remove_by_index(_item)
        elif isinstance(_item, str):
            self.remove_by_name(_item)

            raise TypeError(f"Unsupported type: {type(_item)}")

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
            frm = self.lo_inst.create_instance_mcf(XForm, "stardiv.one.form.component.Form", raise_err=True)
            frm.Name = self._create_name("Form")  # type: ignore
            self.insert_by_index(idx, frm)
            return self.get_by_index(idx)
        elif isinstance(arg1, str):
            if self.has_by_name(arg1):
                raise mEx.NameClashError(f"Name '{arg1}' already exists")
            frm = self.lo_inst.create_instance_mcf(XForm, "stardiv.one.form.component.Form", raise_err=True)
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
            MissingNameError: If form is not found.

        Returns:
            WriteForm: The element with the specified name.
        """
        if not self.has_by_name(name):
            raise mEx.MissingNameError(f"Unable to find form with name '{name}'")
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
