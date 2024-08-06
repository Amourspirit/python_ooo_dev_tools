# region Import
from __future__ import annotations
from typing import Any, Tuple, Type, TypeVar, overload
from ooo.dyn.text.wrap_text_mode import WrapTextMode

from ooodev.exceptions import ex as mEx
from ooodev.format.inner.common.props.frame_wrap_settings_props import FrameWrapSettingsProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.utils import props as mProps

# endregion Import

_TSettings = TypeVar("_TSettings", bound="Settings")


class Settings(StyleBase):
    """
    Frame Vertical Alignment

    .. versionadded:: 0.9.0
    """

    def __init__(self, mode: WrapTextMode = WrapTextMode.PARALLEL) -> None:
        """
        Constructor

        Args:
            mode (WrapTextMode): Specifies Wrap mode. Default ``WrapTextMode.PARALLEL``
        """
        super().__init__()
        self.prop_mode = mode

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.Style",
                "com.sun.star.text.BaseFrame",
                "com.sun.star.text.TextEmbeddedObject",
                "com.sun.star.text.TextFrame",
                "com.sun.star.text.TextGraphicObject",
            )
        return self._supported_services_values

    def _is_valid_obj(self, obj: Any) -> bool:
        """
        Gets if ``obj`` supports one of the services required by style class

        Args:
            obj (Any): UNO object that must have requires service

        Returns:
            bool: ``True`` if has a required service; Otherwise, ``False``
        """
        if super()._is_valid_obj(obj):
            return True
        # check if obj has matching property
        # Some objects such as 'com.sun.star.drawing.shape' sometime support this style.
        # Such is the case when a shape is added to a Writer drawing page.
        # Assume if on attribute matches then it is a match.
        return hasattr(obj, self._props.name)

    # endregion Overrides
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TSettings], obj: object) -> _TSettings: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TSettings], obj: object, **kwargs) -> _TSettings: ...

    @classmethod
    def from_obj(cls: Type[_TSettings], obj: object, **kwargs) -> _TSettings:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Settings: Instance that represents Frame Wrap Settings.
        """
        # this nu is only used to get Property Name

        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        inst.prop_mode = mProps.Props.get(obj, inst._props.name)
        return inst

    # endregion from_obj()

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE
        return self._format_kind_prop

    @property
    def prop_mode(self) -> WrapTextMode:
        """Gets/Sets Wrap mode value"""
        return self._get(self._props.name)

    @prop_mode.setter
    def prop_mode(self, value: WrapTextMode) -> None:
        self._set(self._props.name, value)

    @property
    def _props(self) -> FrameWrapSettingsProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = FrameWrapSettingsProps(name="Surround")
        return self._props_internal_attributes
