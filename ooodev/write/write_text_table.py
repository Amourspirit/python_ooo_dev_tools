from __future__ import annotations
from typing import cast, TYPE_CHECKING, TypeVar, Generic
import uno

from ooodev.mock import mock_g
from ooodev.adapter.text.text_table_comp import TextTableComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write import write_text_portions as mWriteTextPortions
from ooodev.proto.component_proto import ComponentT

if TYPE_CHECKING:
    from com.sun.star.text import XTextTable
    from com.sun.star.container import XEnumerationAccess
    from ooodev.write.table.write_table_rows import WriteTableRows

T = TypeVar("T", bound="ComponentT")


class WriteTextTable(Generic[T], LoInstPropsPartial, WriteDocPropPartial, TextTableComp, QiPartial, StylePartial):
    """Represents writer text content."""

    # this class can be used to wrap the table created by
    # ooodev.office.write.Write.add_table() method.

    # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextTable.html

    def __init__(self, owner: T, component: XTextTable, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XTextTable): UNO object that supports ``com.sun.star.text.TextContent`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        if not isinstance(owner, WriteDocPropPartial):
            raise TypeError("WriteDocPropPartial is not inherited by owner.")
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        TextTableComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        StylePartial.__init__(self, component=component)

    def get_text_portions(self) -> mWriteTextPortions.WriteTextPortions[T]:
        """Returns the text portions of this paragraph."""
        return mWriteTextPortions.WriteTextPortions(
            owner=self.owner, component=cast("XEnumerationAccess", self.component), lo_inst=self.lo_inst
        )

    # region TextTablePartial overrides
    def get_rows(self) -> WriteTableRows:
        """
        Gets the rows of this table.

        Returns:
            WriteTableRows: Table Rows
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.write.table.write_table_rows import WriteTableRows

        rows = self.component.getRows()
        return WriteTableRows(owner=self, component=rows)

    # endregion TextTablePartial overrides

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self._owner

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.write.table.write_table_rows import WriteTableRows
