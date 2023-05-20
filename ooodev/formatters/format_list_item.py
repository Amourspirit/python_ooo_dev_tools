from __future__ import annotations
from typing import Sequence, Tuple


class FormatListItem:
    def __init__(self, format: str | Tuple[str, ...], idxs: Sequence[int]) -> None:
        self._format = format
        self._idxs = idxs

    """Custom Format. Applies to specific indexes of a list"""

    def has_index(self, idx: int) -> bool:
        """
        Gets if instance contains ``idx``

        Args:
            idx (int): Index value

        Returns:
            bool: ``True`` if index is found; Otherwise, ``False``
        """
        return idx in self._idxs

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
    def idxs(self) -> Sequence[int]:
        """Gets indexes of a list that this formatting applies to"""
        return self._idxs


__all__ = ["FormatListItem"]
