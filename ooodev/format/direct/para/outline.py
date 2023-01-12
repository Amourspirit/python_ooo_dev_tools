"""
Modele for managing paragraph Outline.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, cast, overload
from enum import IntEnum

from ....exceptions import ex as mEx
from ....meta.static_prop import static_prop
from ....utils import lo as mLo
from ....utils import props as mProps
from ...kind.style_kind import StyleKind
from ...style_base import StyleBase


class LevelKind(IntEnum):
    """Outline Level"""

    TEXT_BODY = 0
    LEVEL_01 = 1
    LEVEL_02 = 2
    LEVEL_03 = 3
    LEVEL_04 = 4
    LEVEL_05 = 5
    LEVEL_06 = 6
    LEVEL_07 = 7
    LEVEL_08 = 8
    LEVEL_09 = 9
    LEVEL_10 = 10


class Outline(StyleBase):
    """
    Paragraph Outline

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``style_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    # region init

    def __init__(self, level: LevelKind = LevelKind.TEXT_BODY) -> None:
        """
        Constructor

        Args:
            level (LevelKind): Outline level.

        Returns:
            None:
        """
        super().__init__(**{"OutlineLevel": level.value})

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.ParagraphProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.ParagraphProperties",)

    @overload
    def apply_style(self, obj: object) -> None:
        ...

    def apply_style(self, obj: object, **kwargs) -> None:
        """
        Applies break properties to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

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
    def from_obj(obj: object) -> Outline:
        """
        Gets instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support  ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            Outline: ``Outline`` instance that represents ``obj`` break properties.
        """
        inst = Outline(level=LevelKind.TEXT_BODY)
        if not inst._is_valid_service(obj):
            raise mEx.NotSupportedServiceError(inst._supported_services()[0])

        level = int(mProps.Props.get(obj, "OutlineLevel"))
        inst._set("OutlineLevel", level)
        return inst

    # endregion methods
    # region style methods
    def style_above(self, value: LevelKind | None) -> Outline:
        """
        Gets a copy of instance with level set.

        Args:
            value (LevelKind | None): Level value

        Returns:
            Spacing: Outline instance
        """
        cp = self.copy()
        cp.prop_level = value
        return cp

    # endregion style methods

    # region Style Properties
    @property
    def text_body(self) -> Outline:
        """GEts copy of instance set to outline level text body."""
        cp = self.copy()
        cp.prop_level = LevelKind.TEXT_BODY
        return cp

    @property
    def level_01(self) -> Outline:
        """GEts copy of instance set to outline level ``1``."""
        cp = self.copy()
        cp.prop_level = LevelKind.LEVEL_01
        return cp

    @property
    def level_02(self) -> Outline:
        """GEts copy of instance set to outline level ``2``."""
        cp = self.copy()
        cp.prop_level = LevelKind.LEVEL_02
        return cp

    @property
    def level_03(self) -> Outline:
        """GEts copy of instance set to outline level ``3``."""
        cp = self.copy()
        cp.prop_level = LevelKind.LEVEL_03
        return cp

    @property
    def level_04(self) -> Outline:
        """GEts copy of instance set to outline level ``4``."""
        cp = self.copy()
        cp.prop_level = LevelKind.LEVEL_04
        return cp

    @property
    def level_05(self) -> Outline:
        """GEts copy of instance set to outline level ``5``."""
        cp = self.copy()
        cp.prop_level = LevelKind.LEVEL_05
        return cp

    @property
    def level_06(self) -> Outline:
        """GEts copy of instance set to outline level ``6``."""
        cp = self.copy()
        cp.prop_level = LevelKind.LEVEL_06
        return cp

    @property
    def level_07(self) -> Outline:
        """GEts copy of instance set to outline level ``7``."""
        cp = self.copy()
        cp.prop_level = LevelKind.LEVEL_07
        return cp

    @property
    def level_08(self) -> Outline:
        """GEts copy of instance set to outline level ``8``."""
        cp = self.copy()
        cp.prop_level = LevelKind.LEVEL_08
        return cp

    @property
    def level_09(self) -> Outline:
        """GEts copy of instance set to outline level ``9``."""
        cp = self.copy()
        cp.prop_level = LevelKind.LEVEL_09
        return cp

    @property
    def level_10(self) -> Outline:
        """GEts copy of instance set to outline level ``10``."""
        cp = self.copy()
        cp.prop_level = LevelKind.LEVEL_10
        return cp

    # endregion Style Properties

    # region properties
    @property
    def prop_style_kind(self) -> StyleKind:
        """Gets the kind of style"""
        return StyleKind.PARA

    @property
    def prop_level(self) -> LevelKind:
        """Gets/Sets level"""
        pv = cast(int, self._get("OutlineLevel"))
        return LevelKind(pv)

    @prop_level.setter
    def prop_level(self, value: LevelKind) -> LevelKind:
        self._set("OutlineLevel", value.value)

    @static_prop
    def default() -> Outline:  # type: ignore[misc]
        """Gets ``Outline`` default. Static Property."""
        if Outline._DEFAULT is None:
            Outline._DEFAULT = Outline(level=LevelKind.TEXT_BODY)
        return Outline._DEFAULT

    # endregion properties
