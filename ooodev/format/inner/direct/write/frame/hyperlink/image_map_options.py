# region imports
from __future__ import annotations
from typing import Any, Tuple, overload, Type, TypeVar

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps

# endregion imports

_TImageMapOptions = TypeVar(name="_TImageMapOptions", bound="ImageMapOptions")


class ImageMapOptions(StyleBase):
    """
    Image Map Options

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(self, *, server_map: bool = False) -> None:
        """
        Constructor

        Args:
            server_map (str, optional): Server side image map. Defaults to ``False``.

        Returns:
            None:
        """

        super().__init__()
        self.prop_server_map = server_map

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.text.TextFrame",
                "com.sun.star.text.TextGraphicObject",
                "com.sun.star.text.BaseFrame",
                "com.sun.star.text.TextEmbeddedObject",
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
        Applies padding to ``obj``

        Args:
            obj (Any): UNO object that supports ``com.sun.star.style.CharacterProperties`` service.
            kwargs (Any, optional): Expandable list of key value pairs that may be used in child classes.

        Returns:
            None:
        """
        try:
            super().apply(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")
        return None

    # endregion apply()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TImageMapOptions], obj: object) -> _TImageMapOptions: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TImageMapOptions], obj: object, **kwargs) -> _TImageMapOptions: ...

    @classmethod
    def from_obj(cls: Type[_TImageMapOptions], obj: object, **kwargs) -> _TImageMapOptions:
        """
        Gets Image Map info instance from object.

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            ImageMapOptions: Instance that represents ``obj`` Image map options.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        inst._set("ServerMap", mProps.Props.get(obj, "ServerMap", False))
        return inst

    # endregion from_obj()
    # endregion methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.FRAME
        return self._format_kind_prop

    @property
    def prop_server_map(self) -> bool:
        """Gets/Sets server side"""
        return self._get("ServerMap")

    @prop_server_map.setter
    def prop_server_map(self, value: bool):
        self._set("ServerMap", value)

    # endregion Properties
