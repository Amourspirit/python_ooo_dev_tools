"""
This module does not currently work.
It is not clear what property needs to be set to change the writing mode.

The shape has a property called ``TextWritingMode``. But changing it does not seem to have any effect.
There is also a property called ``WritingMode``. But also changing it does not seem to have any effect.
When a shapes changes the writing mode via the dialog it seems to change the shape.Start.WritingMode property;
however it seems not possible to set this property manually.
"""
from __future__ import annotations
from typing import Any, Tuple, overload, Type, TypeVar
import uno

from ooo.dyn.text.writing_mode import WritingMode

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase


_TTextDirection = TypeVar(name="_TTextDirection", bound="TextDirection")


class TextDirection(StyleBase):
    """
    Shape Paragraph Text Direction

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.17.8
    """

    # region init

    def __init__(self, mode: WritingMode | None = None) -> None:
        """
        Constructor

        Args:
            mode (WritingMode, optional): Determines the writing direction

        Returns:
            None:
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        init_vals = {}

        if mode is not None:
            init_vals[self._get_property_name()] = mode

        super().__init__(**init_vals)

    # endregion init

    # region methods
    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "TextWritingMode"
        return self._property_name

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.drawing.TextProperties",)
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TTextDirection], obj: Any) -> _TTextDirection:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TTextDirection], obj: Any, **kwargs) -> _TTextDirection:
        ...

    @classmethod
    def from_obj(cls: Type[_TTextDirection], obj: Any, **kwargs) -> _TTextDirection:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TextDirection: ``WritingMode`` instance that represents ``obj`` writing mode.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        inst._set("TextWritingMode", mProps.Props.get(obj, inst._get_property_name()))
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
            self._format_kind_prop = FormatKind.PARA | FormatKind.SHAPE
        return self._format_kind_prop

    @property
    def prop_mode(self) -> WritingMode | None:
        """Gets/Sets writing mode of a paragraph."""
        return self._get(self._get_property_name())

    @prop_mode.setter
    def prop_mode(self, value: WritingMode | None):
        if value is None:
            self._remove(self._get_property_name())
            return
        self._set(self._get_property_name(), value)

    @property
    def default(self: _TTextDirection) -> _TTextDirection:
        """Gets ``TextDirection`` default."""
        try:
            return self._default_inst
        except AttributeError:
            self._default_inst = self.__class__(mode=WritingMode.LR_TB, _cattribs=self._get_internal_cattribs())
            self._default_inst._is_default_inst = True
        return self._default_inst

    # endregion properties
