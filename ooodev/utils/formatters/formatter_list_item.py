from __future__ import annotations
from typing import Tuple
from dataclasses import dataclass
from ..validation import check


@dataclass
class FormatterListItem:
    """Custom Format. Applies to specific indexes of a list"""

    format: str | Tuple[str, ...]
    """
    Format option such as ``.2f``
    
    Multiple formats can be added such as ``(".2f", "<10")``.
    Formats are applied in the order they are added.
    In this case first float is formated as string with two decimal places, and
    then value is padded to the right with spaces.
    """
    idxs: Tuple[int, ...]
    """Specific index of a list that this formatting applies to"""

    def __post_init__(self) -> None:
        s_str = f"{self}"
        msg = "indexes can only be of type int"
        for index in self.idxs:
            check(isinstance(index, int), s_str, msg)

    def has_index(self, idx: int) -> bool:
        """
        Gets if instance contains ``idx``

        Args:
            idx (int): Index value

        Returns:
            bool: ``True`` if index is found; Otherwise, ``False``
        """
        return idx in self.idxs
