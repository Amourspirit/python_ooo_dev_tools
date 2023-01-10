"""
Modele for managing paragraph Writing Mode.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, cast, overload

from ...exceptions import ex as mEx
from ...meta.static_prop import static_prop
from ...utils import lo as mLo
from ...utils import props as mProps
from ..kind.style_kind import StyleKind
from ..style_base import StyleBase

from ooo.dyn.text.writing_mode2 import WritingMode2Enum as WritingMode2Enum


class WritingMode(StyleBase):
    """
    Paragraph Writeing Mode

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``style_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    # region init

    def __init__(self, mode: WritingMode2Enum | None = None) -> None:
        """
        Constructor

        Args:
            mode (WritingMode2Enum, optional): Determines the writing direction

        Returns:
            None:
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        init_vals = {}

        if not mode is None:
            init_vals["WritingMode"] = mode.value

        super().__init__(**init_vals)

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.ParagraphPropertiesComplex``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.ParagraphPropertiesComplex",)

    @overload
    def apply_style(self, obj: object) -> None:
        ...

    def apply_style(self, obj: object, **kwargs) -> None:
        """
        Applies writing mode to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphPropertiesComplex`` service.

        Returns:
            None:
        """
        try:
            super().apply_style(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__}.apply_style(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    @staticmethod
    def from_obj(obj: object) -> WritingMode:
        """
        Gets instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphPropertiesComplex`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support  ``com.sun.star.style.ParagraphPropertiesComplex`` service.

        Returns:
            WritingMode: ``WritingMode`` instance that represents ``obj`` writing mode.
        """
        inst = WritingMode()
        if inst._is_valid_service(obj):
            inst._set("WritingMode", int(mProps.Props.get(obj, "WritingMode")))
        else:
            raise mEx.NotSupportedServiceError(inst._supported_services()[0])
        return inst

    # endregion methods

    # region style methods
    def style_mode(self, value: WritingMode2Enum | None) -> WritingMode:
        """
        Gets copy of instance with writing mode set or removed

        Args:
            value (ParagraphAdjust | None): mode value

        Returns:
            WritingMode: ``WritingMode`` instance
        """
        cp = self.copy()
        cp.prop_align = value
        return cp

    # endregion style methods

    # region Style Properties
    @property
    def bt_lr(self) -> WritingMode:
        """
        Gets instance.

        Text within a line is written bottom-to-top.
        Lines and blocks are placed left-to-right.
        """
        cp = self.copy()
        cp.prop_mode = WritingMode2Enum.BT_LR
        return cp

    @property
    def lr_tb(self) -> WritingMode:
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
    def rl_tb(self) -> WritingMode:
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
    def tb_rl(self) -> WritingMode:
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
    def tb_lr(self) -> WritingMode:
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
    def page(self) -> WritingMode:
        """
        Gets instance.

        Obtain writing mode from the current page.
        May not be used in page styles.
        """
        cp = self.copy()
        cp.prop_mode = WritingMode2Enum.PAGE
        return cp

    @property
    def context(self) -> WritingMode:
        """
        Gets instance.

        Obtain actual writing mode from the context of the object.
        """
        cp = self.copy()
        cp.prop_mode = WritingMode2Enum.CONTEXT
        return cp

    # endregion Style Properties

    # region properties
    @property
    def prop_style_kind(self) -> StyleKind:
        """Gets the kind of style"""
        return StyleKind.PARA | StyleKind.PARA_COMPLEX

    @property
    def prop_mode(self) -> WritingMode2Enum | None:
        """Gets/Sets wrighting mode of a paragraph."""
        pv = cast(int, self._get("WritingMode"))
        if pv is None:
            return None
        return WritingMode2Enum(pv)

    @prop_mode.setter
    def prop_mode(self, value: WritingMode2Enum | None):
        if value is None:
            self._remove("WritingMode")
            return
        self._set("WritingMode", value)

    @static_prop
    def default(cls) -> WritingMode:  # type: ignore[misc]
        """Gets ``WritingMode`` default. Static Property."""
        if cls._DEFAULT is None:
            cls._DEFAULT = WritingMode(WritingMode2Enum.PAGE)
        return cls._DEFAULT

    # endregion properties
