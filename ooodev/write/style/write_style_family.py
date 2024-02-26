from __future__ import annotations
from typing import Any, List, TypeVar, TYPE_CHECKING
import uno

from ooodev.adapter.container.index_access_partial import IndexAccessPartial
from ooodev.adapter.container.name_container_partial import NameContainerPartial
from ooodev.adapter.style.style_family_comp import StyleFamilyComp
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.write.style.write_style import WriteStyle
from ooodev.proto.component_proto import ComponentT

if TYPE_CHECKING:
    from ooodev.write.write_doc import WriteDoc

T = TypeVar("T", bound="ComponentT")


class WriteStyleFamily(
    StyleFamilyComp,
    IndexAccessPartial,
    NameContainerPartial,
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
            owner (WriteDoc): Owner of this component.
            component (Any): UNO object that supports ``com.sun.star.text.Paragraph`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            self._lo_inst = mLo.Lo.current_lo
        else:
            self._lo_inst = lo_inst
        self.__owner = owner
        StyleFamilyComp.__init__(self, component)
        IndexAccessPartial.__init__(self, component=self.component)  # type: ignore
        NameContainerPartial.__init__(self, component=self.component)  # type: ignore
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

    def get_style(self, name: str) -> WriteStyle[WriteDoc]:
        """
        Gets a style by name.

        Args:
            name (str): Name of style to get.

        Raises:
            Exception: If unable to get style.

        Returns:
            WriteStyle: Style
        """
        # sourcery skip: raise-specific-error
        try:
            style = self.get_by_name(name)
            return WriteStyle(owner=self.owner, component=style, lo_inst=self._lo_inst)
        except Exception as e:
            raise mEx.StyleError(f"Unable to get style '{name}'") from e

    # region Properties
    @property
    def owner(self) -> WriteDoc:
        """Owner of this component."""
        return self.__owner

    # endregion Properties
