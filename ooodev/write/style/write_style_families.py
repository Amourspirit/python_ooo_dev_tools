from __future__ import annotations
from typing import Any, List, TYPE_CHECKING
import uno

from ooodev.adapter.style.style_families_comp import StyleFamiliesComp
from ooodev.adapter.container.index_access_partial import IndexAccessPartial
from ooodev.loader import lo as mLo
from ooodev.exceptions import ex as mEx
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.format.writer.style.family_names_kind import FamilyNamesKind
from ooodev.loader.inst.lo_inst import LoInst
from .write_style_family import WriteStyleFamily

if TYPE_CHECKING:
    from ooodev.write.write_doc import WriteDoc


class WriteStyleFamilies(
    StyleFamiliesComp,
    IndexAccessPartial,
    QiPartial,
):
    """
    Represents writer paragraph content.

    Contains Enumeration Access.
    """

    def __init__(self, owner: WriteDoc, component: Any, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (Any): UNO object that supports ``com.sun.star.text.Paragraph`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        self._lo_inst = mLo.Lo.current_lo if lo_inst is None else lo_inst
        self.__owner = owner
        StyleFamiliesComp.__init__(self, component)
        IndexAccessPartial.__init__(self, component=self.component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self._lo_inst)  # type: ignore
        # self.__doc = doc

    def get_names(self) -> List[str]:
        """
        Gets a sorted list of style family names

        Raises:
            Exception: If unable to names.

        Returns:
            List[str]: List of style names
        """
        # sourcery skip: raise-specific-error
        try:
            names = list(self.get_element_names())
            return sorted(names)
        except Exception as e:
            raise Exception("Unable to get family style names") from e

    def get_style_family(self, name: str | FamilyNamesKind) -> WriteStyleFamily:
        """
        Gets a style family by name.

        Args:
            name (str | FamilyNamesKind): Name of style family.

        Raises:
            Exception: If unable to get style family.

        Returns:
            WriteStyleFamily: Style family.
        """
        try:
            result = self.get_by_name(str(name))
            return WriteStyleFamily(owner=self.owner, component=result, lo_inst=self._lo_inst)
        except Exception as e:
            raise mEx.StyleError(f"Unable to get style family: {name}") from e

    # region Properties
    @property
    def owner(self) -> WriteDoc:
        """Owner of this component."""
        return self.__owner

    # endregion Properties
