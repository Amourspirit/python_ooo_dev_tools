"""
Modele for managing paragraph padding.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import cast, overload

from ...exceptions import ex as mEx
from ...meta.static_prop import static_prop
from ...utils import info as mInfo
from ...utils import lo as mLo
from ...utils import props as mProps
from ..style_base import StyleBase

from ooo.dyn.text.writing_mode2 import WritingMode2Enum as WritingMode2Enum


class WritingMode(StyleBase):
    """
    Paragraph Alignment

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``style_`` can be used to chain together Padding properties.

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
        if mInfo.Info.support_service(obj, "com.sun.star.style.ParagraphPropertiesComplex"):
            try:
                super().apply_style(obj)
            except mEx.MultiError as e:
                mLo.Lo.print(f"{self.__class__}.apply_style(): Unable to set Property")
                for err in e.errors:
                    mLo.Lo.print(f"  {err}")
        else:
            mLo.Lo.print('Padding.apply_style(): "com.sun.star.style.ParagraphPropertiesComplex" not supported')
        return None

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
        if not mInfo.Info.support_service(obj, "com.sun.star.style.ParagraphPropertiesComplex"):
            raise mEx.NotSupportedServiceError("com.sun.star.style.ParagraphPropertiesComplex")
        wm = WritingMode()
        wm._set("WritingMode", int(mProps.Props.get(obj, "WritingMode")))
        return wm

    # endregion methods

    # region style methods
    def style_mode(self, value: WritingMode2Enum | None) -> WritingMode:
        """
        Gets copy of instance with writing mode set or removed

        Args:
            value (ParagraphAdjust | None): Alignment value

        Returns:
            Alignment: Alignment instance
        """
        cp = self.copy()
        cp.prop_align = value
        return cp

    # endregion style methods

    # region properties

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
    def default(cls) -> WritingMode:
        """Gets ``WritingMode`` default. Static Property."""
        if cls._DEFAULT is None:
            cls._DEFAULT = WritingMode(WritingMode2Enum.PAGE)
        return cls._DEFAULT

    # endregion properties
