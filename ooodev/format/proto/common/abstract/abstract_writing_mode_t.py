from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING


from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Self
    from typing_extensions import Protocol
    from ooo.dyn.text.writing_mode2 import WritingMode2Enum
else:
    Protocol = object
    Self = Any
    WritingMode2Enum = Any


# see ooodev.format.inner.direct.calc.numbers.numbers.Numbers
class AbstractWritingModeT(StyleT, Protocol):
    """Paragraph Writing Mode"""

    def __init__(self, mode: WritingMode2Enum | None = None) -> None:
        """
        Constructor

        Args:
            mode (WritingMode2Enum, optional): Determines the writing direction.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> AbstractWritingModeT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> AbstractWritingModeT:
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

    @property
    def prop_mode(self) -> WritingMode2Enum | None:
        """Gets/Sets writing mode of a paragraph."""
        ...

    @prop_mode.setter
    def prop_mode(self, value: WritingMode2Enum | None): ...

    @property
    def default(self) -> Self:
        """Gets ``WritingMode`` default."""
        ...

    # endregion Properties
