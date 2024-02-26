from __future__ import annotations
from typing import Tuple
from ooodev.formatters.table_item_kind import TableItemKind as TableItemKind


# if include indexes are available but exclude is not then apply to all includes only
# if exclude indexes available but include is not then apply to all includes that are not in excludes
# if both include and exclude are available then any index in both include and exclude will NOT be formatted.


class FormatTableItem:
    def __init__(
        self,
        format: str | Tuple[str, ...],
        idxs_inc: Tuple[int, ...] | None = None,
        row_idxs_exc: Tuple[int, ...] | None = None,
    ) -> None:
        self._format = format
        self._idxs_inc: Tuple[int, ...] = idxs_inc or ()
        self._row_idxs_exc: Tuple[int, ...] = row_idxs_exc or ()
        self._item_kind: TableItemKind = TableItemKind.NONE

    """Custom Format. Applies to specific indexes of a list"""

    def is_index(self, idx: int) -> bool:
        """
        Gets if instance includes ``idx`` for formatting.

        Args:
            idx (int): Index value

        Returns:
            bool: ``True`` if index is found; Otherwise, ``False``
        """
        return idx in self._idxs_inc

    def is_row_exc_index(self, idx: int) -> bool:
        """
        Gets if instance excludes row ``idx`` from formatting.

        Args:
            idx (int): Index value

        Returns:
            bool: ``True`` if index is found; Otherwise, ``False``
        """
        return idx in self._row_idxs_exc

    @property
    def format(self) -> str | Tuple[str, ...]:
        """
        Gets Format

        Format option such as ``.2f``

        Multiple formats can be added such as ``(".2f", "<10")``.
        Formats are applied in the order they are added.
        In this case first float is formatted as string with two decimal places, and
        then value is padded to the right with spaces.
        """
        return self._format

    @property
    def idxs_inc(self) -> Tuple[int, ...]:
        """Gets indexes that formatting applies to"""
        return self._idxs_inc

    @property
    def row_idxs_exc(self) -> Tuple[int, ...]:
        """Gets indexes of rows that are excluded from formatted"""
        return self._row_idxs_exc

    @property
    def item_kind(self) -> TableItemKind:
        """Gets/Sets Item Kind"""
        return self._item_kind

    @item_kind.setter
    def item_kind(self, value: TableItemKind):
        self._item_kind = value


__all__ = ["FormatTableItem"]
