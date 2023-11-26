from __future__ import annotations
import contextlib
from typing import TypeVar, Type
import uno
from ooo.dyn.awt.size import Size as UnoSize
from ...proto.size_obj import SizeObj

from .generic_size import GenericSize

_TSize = TypeVar(name="_TSize", bound="Size")


class Size(GenericSize[int]):
    """
    Represents a size with positive values.

    See Also:
        :ref:`proto_size_obj`
    """

    def get_uno_size(self) -> UnoSize:
        """Gets UNO instance from current values"""
        return UnoSize(self.width, self.height)

    @classmethod
    def from_size(cls: Type[_TSize], sz: SizeObj) -> _TSize:
        """
        Gets instance from Size.

        Args:
            sz (Size): Size object, Can be UNO Size.

        Returns:
            Size: Size instance from Size values.
        """
        inst = super(Size, cls).__new__(cls)
        inst.__init__(sz.Width, sz.Height)
        return inst
