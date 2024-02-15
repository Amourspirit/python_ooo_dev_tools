from __future__ import annotations
from enum import Enum
from typing import Tuple, cast
import uno
from ooo.dyn.text.writing_mode2 import WritingMode2
from ooo.dyn.text.writing_mode2 import WritingMode2Enum

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.common.abstract.abstract_writing_mode import AbstractWritingMode


class DirectionModeKind(Enum):
    """
    Describes different text directions.
    """

    LR_TB = WritingMode2.LR_TB  # keep
    """
    Left-to-right (LTR)
    
    Text within lines is written left-to-right.
    Lines and blocks are placed top-to-bottom.
    Typically, this is the writing mode for normal ``alphabetic`` text.
    """
    RL_TB = WritingMode2.RL_TB  # keep
    """
    Right-to-left (RTL).
    
    text within a line are written right-to-left.
    Lines and blocks are placed top-to-bottom.
    Typically, this writing mode is used in Arabic and Hebrew text.
    """
    PAGE = WritingMode2.PAGE  # keep, use superordinate object settings
    """
    Use super-ordinate object settings
    
    Obtain writing mode from the current page.
    May not be used in page styles.
    """

    def __int__(self) -> int:
        return self.value


class Direction(AbstractWritingMode):
    """
    Title Text Direction.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. seealso::

        - :ref:`help_chart2_format_direct_title_alignment`

    .. versionadded:: 0.9.4
    """

    def __init__(self, mode: DirectionModeKind = DirectionModeKind.PAGE) -> None:
        """
        Constructor

        Args:
            mode (DirectionModeKind, optional): Specifies the writing direction.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_title_alignment`
        """
        super().__init__(mode=WritingMode2Enum(mode.value))

    # region overrides
    def _container_get_service_name(self) -> str:
        # keep type hinting happy
        raise NotImplementedError

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.beans.PropertySet",)
        return self._supported_services_values

    # endregion overrides

    # region style methods
    def fmt_mode(self, value: DirectionModeKind) -> Direction:
        """
        Gets copy of instance with writing mode set or removed

        Args:
            value (DirectionModeKind | None): mode value

        Returns:
            Direction: ``Direction`` instance.
        """
        cp = self.copy()
        cp.prop_mode = value
        return cp

    # endregion style methods

    # region Style Properties
    @property
    def lr_tb(self) -> Direction:
        """
        Gets instance.

        Text within lines is written left-to-right.
        Lines and blocks are placed top-to-bottom.
        Typically, this is the writing mode for normal ``alphabetic`` text.
        """
        cp = self.copy()
        cp.prop_mode = DirectionModeKind.LR_TB
        return cp

    @property
    def rl_tb(self) -> Direction:
        """
        Gets instance.

        Text within a line are written right-to-left.
        Lines and blocks are placed top-to-bottom.
        Typically, this writing mode is used in Arabic and Hebrew text.
        """
        cp = self.copy()
        cp.prop_mode = DirectionModeKind.RL_TB
        return cp

    @property
    def page(self) -> Direction:
        """
        Gets instance.

        Use super-ordinate object settings.
        """
        cp = self.copy()
        cp.prop_mode = DirectionModeKind.PAGE
        return cp

    # endregion Style Properties

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.UNKNOWN
        return self._format_kind_prop

    @property
    def prop_mode(self) -> DirectionModeKind:
        """Gets/Sets writing mode."""
        pv = cast(int, self._get(self._get_property_name()))
        return DirectionModeKind(pv)

    @prop_mode.setter
    def prop_mode(self, value: DirectionModeKind):
        self._set(self._get_property_name(), value.value)

    # endregion Properties
