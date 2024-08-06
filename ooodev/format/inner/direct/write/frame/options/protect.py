from __future__ import annotations
from typing import overload
from typing import Any, Tuple, Type, TypeVar
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.format.inner.common.props.frame_options_protect_props import FrameOptionsProtectProps

_TProtect = TypeVar("_TProtect", bound="Protect")


class Protect(StyleBase):
    """
    Frame Protections

    .. versionadded:: 0.9.0
    """

    def __init__(self, size: bool | None = None, position: bool | None = None, content: bool | None = None) -> None:
        """
        Constructor

        Args:
            size (bool, optional): Specifies size protection.
            position (bool, optional): Specifies position protection.
            content (bool, optional): Specifies content protection.
        """
        super().__init__()
        self.prop_size = size
        self.prop_position = position
        self.prop_content = content

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
    def from_obj(cls: Type[_TProtect], obj: object) -> _TProtect: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TProtect], obj: object, **kwargs) -> _TProtect: ...

    @classmethod
    def from_obj(cls: Type[_TProtect], obj: object, **kwargs) -> _TProtect:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Protect: Instance that represents Frame Protection.
        """
        # this nu is only used to get Property Name

        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        inst.prop_size = mProps.Props.get(obj, inst._props.size)
        inst.prop_position = mProps.Props.get(obj, inst._props.pos)
        inst.prop_content = mProps.Props.get(obj, inst._props.content)
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
    def prop_size(self) -> bool | None:
        """Gets/Sets size"""
        return self._get(self._props.size)

    @prop_size.setter
    def prop_size(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.size)
            return
        self._set(self._props.size, value)

    @property
    def prop_position(self) -> bool | None:
        """Gets/Sets position"""
        return self._get(self._props.pos)

    @prop_position.setter
    def prop_position(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.pos)
            return
        self._set(self._props.pos, value)

    @property
    def prop_content(self) -> bool | None:
        """Gets/Sets content"""
        return self._get(self._props.content)

    @prop_content.setter
    def prop_content(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.content)
            return
        self._set(self._props.content, value)

    @property
    def _props(self) -> FrameOptionsProtectProps:
        try:
            return self._props_frame_opts_protect
        except AttributeError:
            self._props_frame_opts_protect = FrameOptionsProtectProps(
                size="SizeProtected", pos="PositionProtected", content="ContentProtected"
            )
        return self._props_frame_opts_protect
