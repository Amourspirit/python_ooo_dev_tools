from __future__ import annotations
from typing import Any, TypeVar, Generic, Sequence, TYPE_CHECKING
import uno

from ooodev.adapter.text.text_tables_comp import TextTablesComp
from ooodev.loader import lo as mLo
from ooodev.utils import gen_util as mGenUtil
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write.table.write_table import WriteTable
from ooodev.exceptions import ex as mEx


if TYPE_CHECKING:
    from ooodev.proto.component_proto import ComponentT

T = TypeVar("T", bound="ComponentT")


class WriteTables(LoInstPropsPartial, WriteDocPropPartial, TextTablesComp, QiPartial, Generic[T]):
    """
    Represents writer text tables.

    Contains Enumeration Access.

    A table can be added to a document using a cursor.

    Example:

        .. code-block:: python

            doc = WriteDoc.create_doc(loader=loader, visible=True)
            tbl_data = read_table(fnm) # get the table data from a file.
            cursor = doc.get_cursor()
            cursor.append_para("Table of Bond Movies")
            cursor.append_para('The following table comes form "bondMovies.txt"\\n')

            # with doc locks the controllers while the table is being added to the document.
            with doc:
                cursor.add_table(table_data=tbl_data))
                cursor.end_paragraph()

            my_table = doc.tables[-1] # get the last table added to the document.
            cell = my_table["D3"]
            # ...
    """

    def __init__(self, owner: T, component: Any, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XText): UNO object that supports ``com.sun.star.text.Text`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        if not isinstance(owner, WriteDocPropPartial):
            raise TypeError("WriteDocPropPartial is not inherited by owner.")
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        TextTablesComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore

    # region Overrides

    def __next__(self) -> WriteTable[WriteTables[T]]:
        """
        Gets the next element.

        Returns:
            WriteTable[WriteTables[T]]: Next element.
        """
        result = super().__next__()
        return WriteTable(owner=self, component=result, lo_inst=self.lo_inst)

    def __getitem__(self, key: str | int) -> WriteTable[WriteTables[T]]:
        """
        Gets the table at the specified index or name.

        This is short hand for ``get_by_index()`` or ``get_by_name()``.

        Args:
            key (key, str, int): The index or name of the form. When getting by index can be a negative value to get from the end.

        Returns:
            WriteTable[WriteTables[T]]: The table with the specified index or name.

        See Also:
            - :py:meth:`~ooodev.write.table.write_tables.get_by_index`
            - :py:meth:`~ooodev.write.table.write_tables.get_by_name`
        """
        if isinstance(key, int):
            return self.get_by_index(key)
        return self.get_by_name(key)

    # endregion Overrides

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

    def get_by_index(self, idx: int) -> WriteTable[WriteTables[T]]:
        """
        Gets the element at the specified index.

        Args:
            idx (int): The Zero-based index of the element. Idx can be a negative value to index from the end of the list.
                For example, -1 will return the last element.

        Returns:
            WriteTable[WriteTables[T]]: The element at the specified index.
        """
        idx = self._get_index(idx, True)
        result = super().get_by_index(idx)
        return WriteTable(owner=self, component=result, lo_inst=self.lo_inst)

    def get_by_name(self, name: str) -> WriteTable[WriteTables[T]]:
        """
        Gets the element with the specified name.

        Args:
            name (str): The name of the element.

        Raises:
            MissingNameError: If form is not found.

        Returns:
            WriteTable[WriteTables[T]]: The element with the specified name.
        """
        if not self.has_by_name(name):
            raise mEx.MissingNameError(f"Unable to find form with name '{name}'")
        result = super().get_by_name(name)
        return WriteTable(owner=self, component=result, lo_inst=self.lo_inst)

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self._owner

    # endregion Properties
