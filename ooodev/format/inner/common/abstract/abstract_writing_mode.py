from __future__ import annotations
from typing import Any, Tuple, cast, overload, Type, TypeVar

from ooo.dyn.text.writing_mode2 import WritingMode2Enum as WritingMode2Enum

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase


_TAbstractWritingMode = TypeVar("_TAbstractWritingMode", bound="AbstractWritingMode")


class AbstractWritingMode(StyleBase):
    """
    Paragraph Writing Mode

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.4
    """

    # region init

    def __init__(self, mode: WritingMode2Enum | None = None) -> None:
        """
        Constructor

        Args:
            mode (WritingMode2Enum, optional): Determines the writing direction.

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
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:
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
    def from_obj(cls: Type[_TAbstractWritingMode], obj: Any) -> _TAbstractWritingMode: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TAbstractWritingMode], obj: Any, **kwargs) -> _TAbstractWritingMode: ...

    @classmethod
    def from_obj(cls: Type[_TAbstractWritingMode], obj: Any, **kwargs) -> _TAbstractWritingMode:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            WritingMode: ``WritingMode`` instance that represents ``obj`` writing mode.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        inst._set("WritingMode", int(mProps.Props.get(obj, inst._get_property_name())))
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()

    # endregion methods

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
        """Gets/Sets writing mode of a paragraph."""
        pv = cast(int, self._get(self._get_property_name()))
        return None if pv is None else WritingMode2Enum(pv)

    @prop_mode.setter
    def prop_mode(self, value: WritingMode2Enum | None):
        if value is None:
            self._remove(self._get_property_name())
            return
        self._set(self._get_property_name(), value.value)

    @property
    def default(self: _TAbstractWritingMode) -> _TAbstractWritingMode:
        """Gets ``WritingMode`` default."""
        # pylint: disable=unexpected-keyword-arg
        # pylint: disable=protected-access
        try:
            return self._default_inst
        except AttributeError:
            self._default_inst = self.__class__(mode=WritingMode2Enum.PAGE, _cattribs=self._get_internal_cattribs())
            self._default_inst._is_default_inst = True
        return self._default_inst

    # endregion properties
