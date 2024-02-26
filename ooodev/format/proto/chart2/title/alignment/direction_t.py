# pylint: disable=ungrouped-imports
from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.common.abstract.abstract_writing_mode_t import AbstractWritingModeT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from ooodev.format.inner.direct.chart2.title.alignment.direction import DirectionModeKind
else:
    Protocol = object
    DirectionModeKind = Any


# see ooodev.format.inner.direct.calc.numbers.numbers.Numbers
class DirectionT(AbstractWritingModeT, Protocol):
    """Writing Mode"""

    def __init__(self, mode: DirectionModeKind | None = None) -> None:
        """
        Constructor

        Args:
            mode (DirectionModeKind, optional): Determines the writing direction.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> DirectionT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> DirectionT:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Keyword Args:
            component (XComponent): Calc document. Default is current document.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Numbers: Instance that represents numbers format.
        """
        ...

    # region Instance Properties

    # region style methods
    def fmt_mode(self, value: DirectionModeKind) -> DirectionT:
        """
        Gets copy of instance with writing mode set or removed

        Args:
            value (DirectionModeKind | None): mode value

        Returns:
            Direction: ``Direction`` instance.
        """
        ...

    # endregion style methods

    # region Style Properties
    @property
    def lr_tb(self) -> DirectionT:
        """
        Gets instance.

        Text within lines is written left-to-right.
        Lines and blocks are placed top-to-bottom.
        Typically, this is the writing mode for normal ``alphabetic`` text.
        """
        ...

    @property
    def rl_tb(self) -> DirectionT:
        """
        Gets instance.

        Text within a line are written right-to-left.
        Lines and blocks are placed top-to-bottom.
        Typically, this writing mode is used in Arabic and Hebrew text.
        """
        ...

    @property
    def page(self) -> DirectionT:
        """
        Gets instance.

        Use super-ordinate object settings.
        """
        ...

    # endregion Style Properties

    # endregion Properties
