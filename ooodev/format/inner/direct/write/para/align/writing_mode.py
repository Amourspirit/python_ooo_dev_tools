"""
Modele for managing paragraph Writing Mode.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Any, Tuple, cast, overload, Type, TypeVar

import uno
from ooo.dyn.text.writing_mode2 import WritingMode2Enum as WritingMode2Enum

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo
from ooodev.utils import props as mProps
from ooodev.format.kind.format_kind import FormatKind
from ooodev.format.style_base import StyleBase


_TWritingMode = TypeVar(name="_TWritingMode", bound="WritingMode")


class WritingMode(StyleBase):
    """
    Paragraph Writeing Mode

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

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

        if mode is not None:
            init_vals[self._get_property_name()] = mode.value

        super().__init__(**init_vals)

    # endregion init

    # region methods
    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "WritingMode"
        return self._property_name

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.ParagraphPropertiesComplex",
                "com.sun.star.style.ParagraphStyle",
            )
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    # region apply()

    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies writing mode to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphPropertiesComplex`` service.

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
    def from_obj(cls: Type[_TWritingMode], obj: object) -> _TWritingMode:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TWritingMode], obj: object, **kwargs) -> _TWritingMode:
        ...

    @classmethod
    def from_obj(cls: Type[_TWritingMode], obj: object, **kwargs) -> _TWritingMode:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            WritingMode: ``WritingMode`` instance that represents ``obj`` writing mode.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        inst._set("WritingMode", int(mProps.Props.get(obj, inst._get_property_name())))
        return inst

    # endregion from_obj()

    # endregion methods

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
        cp.prop_align = value
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

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA | FormatKind.PARA_COMPLEX
        return self._format_kind_prop

    @property
    def prop_mode(self) -> WritingMode2Enum | None:
        """Gets/Sets wrighting mode of a paragraph."""
        pv = cast(int, self._get(self._get_property_name()))
        if pv is None:
            return None
        return WritingMode2Enum(pv)

    @prop_mode.setter
    def prop_mode(self, value: WritingMode2Enum | None):
        if value is None:
            self._remove(self._get_property_name())
            return
        self._set(self._get_property_name(), value.value)

    @property
    def default(self: _TWritingMode) -> _TWritingMode:
        """Gets ``WritingMode`` default."""
        try:
            return self._default_inst
        except AttributeError:
            self._default_inst = self.__class__(mode=WritingMode2Enum.PAGE, _cattribs=self._get_internal_cattribs())
            self._default_inst._is_default_inst = True
        return self._default_inst

    # endregion properties
