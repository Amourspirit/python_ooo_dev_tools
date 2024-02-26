"""
Module for managing paragraph Outline.

.. seealso::

    - :ref:`help_writer_format_direct_para_outline_and_list`

.. versionadded:: 0.9.0
"""

from __future__ import annotations
from typing import Any, Tuple, cast, overload, Type, TypeVar
from enum import IntEnum

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase

_TOutline = TypeVar(name="_TOutline", bound="Outline")


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

    .. seealso::

        - :ref:`help_writer_format_direct_para_outline_and_list`

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(self, level: LevelKind = LevelKind.TEXT_BODY) -> None:
        """
        Constructor

        Args:
            level (LevelKind): Outline level.

        Returns:
            None:

        See Also:

            - :ref:`help_writer_format_direct_para_outline_and_list`
        """
        super().__init__(OutlineLevel=level.value)

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.ParagraphProperties",
                "com.sun.star.style.ParagraphStyle",
            )
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    # region apply()
    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies break properties to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            None:
        """
        try:
            super().apply(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TOutline], obj: Any) -> _TOutline: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TOutline], obj: Any, **kwargs) -> _TOutline: ...

    @classmethod
    def from_obj(cls: Type[_TOutline], obj: Any, **kwargs) -> _TOutline:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Outline: ``Outline`` instance that represents ``obj`` break properties.
        """
        # pylint: disable=protected-access
        inst = cls(level=LevelKind.TEXT_BODY, **kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        level = int(mProps.Props.get(obj, "OutlineLevel"))
        inst._set("OutlineLevel", level)
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()

    # endregion methods
    # region style methods
    def style_above(self: _TOutline, value: LevelKind) -> _TOutline:
        """
        Gets a copy of instance with level set.

        Args:
            value (LevelKind): Level value

        Returns:
            Spacing: Outline instance
        """
        cp = self.copy()
        cp.prop_level = value
        return cp

    # endregion style methods

    # region Style Properties
    @property
    def text_body(self: _TOutline) -> _TOutline:
        """Gets copy of instance set to outline level text body."""
        cp = self.copy()
        cp.prop_level = LevelKind.TEXT_BODY
        return cp

    @property
    def level_01(self: _TOutline) -> _TOutline:
        """Gets copy of instance set to outline level ``1``."""
        cp = self.copy()
        cp.prop_level = LevelKind.LEVEL_01
        return cp

    @property
    def level_02(self: _TOutline) -> _TOutline:
        """Gets copy of instance set to outline level ``2``."""
        cp = self.copy()
        cp.prop_level = LevelKind.LEVEL_02
        return cp

    @property
    def level_03(self: _TOutline) -> _TOutline:
        """Gets copy of instance set to outline level ``3``."""
        cp = self.copy()
        cp.prop_level = LevelKind.LEVEL_03
        return cp

    @property
    def level_04(self: _TOutline) -> _TOutline:
        """Gets copy of instance set to outline level ``4``."""
        cp = self.copy()
        cp.prop_level = LevelKind.LEVEL_04
        return cp

    @property
    def level_05(self: _TOutline) -> _TOutline:
        """Gets copy of instance set to outline level ``5``."""
        cp = self.copy()
        cp.prop_level = LevelKind.LEVEL_05
        return cp

    @property
    def level_06(self: _TOutline) -> _TOutline:
        """Gets copy of instance set to outline level ``6``."""
        cp = self.copy()
        cp.prop_level = LevelKind.LEVEL_06
        return cp

    @property
    def level_07(self: _TOutline) -> _TOutline:
        """Gets copy of instance set to outline level ``7``."""
        cp = self.copy()
        cp.prop_level = LevelKind.LEVEL_07
        return cp

    @property
    def level_08(self: _TOutline) -> _TOutline:
        """Gets copy of instance set to outline level ``8``."""
        cp = self.copy()
        cp.prop_level = LevelKind.LEVEL_08
        return cp

    @property
    def level_09(self: _TOutline) -> _TOutline:
        """Gets copy of instance set to outline level ``9``."""
        cp = self.copy()
        cp.prop_level = LevelKind.LEVEL_09
        return cp

    @property
    def level_10(self: _TOutline) -> _TOutline:
        """Gets copy of instance set to outline level ``10``."""
        cp = self.copy()
        cp.prop_level = LevelKind.LEVEL_10
        return cp

    # endregion Style Properties

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA
        return self._format_kind_prop

    @property
    def prop_level(self) -> LevelKind:
        """Gets/Sets level"""
        pv = cast(int, self._get("OutlineLevel"))
        return LevelKind(pv)

    @prop_level.setter
    def prop_level(self, value: LevelKind) -> None:
        self._set("OutlineLevel", value.value)

    @property
    def default(self: _TOutline) -> _TOutline:
        """Gets ``Outline`` default."""
        # pylint: disable=protected-access
        # pylint: disable=unexpected-keyword-arg
        try:
            return self._default_inst
        except AttributeError:
            self._default_inst = self.__class__(level=LevelKind.TEXT_BODY, _cattribs=self._get_internal_cattribs())
            self._default_inst._is_default_inst = True
        return self._default_inst

    # endregion properties
