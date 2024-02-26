"""
Module for managing paragraph Writing Mode.

.. versionadded:: 0.9.0
"""

from __future__ import annotations
from typing import TypeVar

from ooo.dyn.text.writing_mode2 import WritingMode2Enum
from ooodev.format.inner.common.abstract.abstract_writing_mode import AbstractWritingMode


_TWritingMode = TypeVar(name="_TWritingMode", bound="WritingMode")


class WritingMode(AbstractWritingMode):
    """
    Paragraph Writing Mode

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region style methods
    def fmt_mode(self: _TWritingMode, value: WritingMode2Enum | None) -> _TWritingMode:
        """
        Gets copy of instance with writing mode set or removed

        Args:
            value (WritingMode2Enum | None): mode value

        Returns:
            WritingMode: ``WritingMode`` instance
        """
        cp = self.copy()
        cp.prop_mode = value
        return cp

    # endregion style methods

    # region Style Properties
    @property
    def bt_lr(self: _TWritingMode) -> _TWritingMode:
        """
        Gets instance.

        Text within a line is written bottom-to-top.
        Lines and blocks are placed left-to-right.
        """
        cp = self.copy()
        cp.prop_mode = WritingMode2Enum.BT_LR
        return cp

    @property
    def lr_tb(self: _TWritingMode) -> _TWritingMode:
        """
        Gets instance.

        Text within lines is written left-to-right.
        Lines and blocks are placed top-to-bottom.
        Typically, this is the writing mode for normal ``alphabetic`` text.
        """
        cp = self.copy()
        cp.prop_mode = WritingMode2Enum.LR_TB
        return cp

    @property
    def rl_tb(self: _TWritingMode) -> _TWritingMode:
        """
        Gets instance.

        Text within a line are written right-to-left.
        Lines and blocks are placed top-to-bottom.
        Typically, this writing mode is used in Arabic and Hebrew text.
        """
        cp = self.copy()
        cp.prop_mode = WritingMode2Enum.RL_TB
        return cp

    @property
    def tb_rl(self: _TWritingMode) -> _TWritingMode:
        """
        Gets instance.

        Text within a line is written top-to-bottom.
        Lines and blocks are placed right-to-left.
        Typically, this writing mode is used in Chinese and Japanese text.
        """
        cp = self.copy()
        cp.prop_mode = WritingMode2Enum.TB_RL
        return cp

    @property
    def tb_lr(self: _TWritingMode) -> _TWritingMode:
        """
        Gets instance.

        Text within a line is written top-to-bottom.
        Lines and blocks are placed left-to-right.
        Typically, this writing mode is used in Mongolian text.
        """
        cp = self.copy()
        cp.prop_mode = WritingMode2Enum.TB_LR
        return cp

    @property
    def page(self: _TWritingMode) -> _TWritingMode:
        """
        Gets instance.

        Obtain writing mode from the current page.
        May not be used in page styles.
        """
        cp = self.copy()
        cp.prop_mode = WritingMode2Enum.PAGE
        return cp

    @property
    def context(self: _TWritingMode) -> _TWritingMode:
        """
        Gets instance.

        Obtain actual writing mode from the context of the object.
        """
        cp = self.copy()
        cp.prop_mode = WritingMode2Enum.CONTEXT
        return cp

    # endregion Style Properties
