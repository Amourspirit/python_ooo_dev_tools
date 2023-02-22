from __future__ import annotations
from typing import Any, Tuple, Type, TypeVar, overload
import uno
from ooo.dyn.text.wrap_text_mode import WrapTextMode as WrapTextMode

from .....exceptions import ex as mEx
from .....utils import lo as mLo
from .....utils import props as mProps
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase
from ...common.props.frame_wrap_settings_props import FrameWrapSettingsProps

_TSettings = TypeVar(name="_TSettings", bound="Settings")


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
            self._supported_services_values = ("com.sun.star.style.Style",)
        return self._supported_services_values

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion Overrides
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TSettings], obj: object) -> _TSettings:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TSettings], obj: object, **kwargs) -> _TSettings:
        ...

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
