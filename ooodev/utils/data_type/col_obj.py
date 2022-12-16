from __future__ import annotations
from dataclasses import dataclass, field
from ..decorator import enforce
from .. import table_helper as mTb


@enforce.enforce_types
@dataclass(frozen=True)
class ColObj:
    """
    Column info.

    .. versionadded:: 0.8.2
    """

    name: str
    """Column such as ``A``"""
    index: int = field(init=False, repr=False, hash=False)

    def __post_init__(self):
        object.__setattr__(self, "name", self.name.upper())
        idx = mTb.TableHelper.col_name_to_int(name=self.name, zero_index=True)
        object.__setattr__(self, "index", idx)

    @staticmethod
    def from_str(name: str) -> ColObj:
        """
        Gets a ``ColObj`` instance from a string

        Args:
            name (str): Cell Name such ``A1``

        Returns:
            CellObj: Cell Object
        """
        num = mTb.TableHelper.col_name_to_int(name=name)
        return ColObj(mTb.TableHelper.make_column_name(num))

    @staticmethod
    def from_int(num: int, zero_index: bool = False) -> ColObj:
        """
        GEts a ``ColObj`` instance from an interger.

        Args:
            num (int): Column number.
            zero_index (bool, optional): Determines if the column number is treated as zero index. Defaults to ``False``.

        Returns:
            ColObj: Cell Object
        """
        ColObj(mTb.TableHelper.make_column_name(num=num, zero_index=zero_index))

    def __str__(self) -> str:
        return self.name

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, ColObj):
            return False
        return str(self) == str(other)
