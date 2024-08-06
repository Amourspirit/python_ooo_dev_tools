# region imports
from __future__ import annotations
from typing import Any, Tuple, overload, Type, TypeVar
from enum import Enum

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.common.props.hyperlink_props import HyperlinkProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps

# endregion imports

_TLinkTo = TypeVar("_TLinkTo", bound="LinkTo")


class TargetKind(Enum):
    """
    Hyperlink Target

    .. versionadded:: 0.9.0
    """

    NONE = ""
    """No target"""
    BLANK = "_blank"
    """Blank target"""
    TOP = "_top"
    """Top target"""
    PARENT = "_parent"
    """Parent target"""
    SELF = "_self"
    """Self target"""

    def __str__(self) -> str:
        return self.value


class LinkTo(StyleBase):
    """
    Link to.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self, *, name: str | None = None, url: str | None = None, target: TargetKind | str = TargetKind.NONE
    ) -> None:
        """
        Constructor

        Args:
            name (str, optional): Link name.
            url (str, optional): Link URL.
            target (TargetKind, str, optional): Link target. Defaults to ``TargetKind.NONE``.

        Returns:
            None:
        """
        init_vals = {self._props.target: str(target)}
        if name is not None:
            init_vals[self._props.name] = name
        if url is not None:
            init_vals[self._props.url] = url

        super().__init__(**init_vals)

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
    def from_obj(cls: Type[_TLinkTo], obj: Any) -> _TLinkTo: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TLinkTo], obj: Any, **kwargs) -> _TLinkTo: ...

    @classmethod
    def from_obj(cls: Type[_TLinkTo], obj: Any, **kwargs) -> _TLinkTo:
        """
        Gets hyperlink instance from object

        Args:
            obj (Any): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            LinkTo: ``LinkTo`` that represents ``obj`` Hyperlink.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        inst._set(inst._props.name, mProps.Props.get(obj, inst._props.name))
        inst._set(inst._props.url, mProps.Props.get(obj, inst._props.url))
        inst._set(inst._props.target, mProps.Props.get(obj, inst._props.target))

        return inst

    # endregion from_obj()
    # endregion methods

    # region Properties

    @property
    def prop_name(self) -> str | None:
        """Gets/Sets name"""
        return self._get(self._props.name)

    @prop_name.setter
    def prop_name(self, value: str | None):
        if value is None:
            if self._has(self._props.name):
                self._remove(self._props.name)
        else:
            self._set(self._props.name, value)

    @property
    def prop_url(self) -> str | None:
        """Gets/Sets url"""
        return self._get(self._props.url)

    @prop_url.setter
    def prop_url(self, value: str | None):
        if value is None:
            if self._has(self._props.url):
                self._remove(self._props.url)
        else:
            self._set(self._props.url, value)

    @property
    def prop_target(self) -> str:
        """Gets/Sets target"""
        return self._get(self._props.target)

    @prop_target.setter
    def prop_target(self, value: TargetKind | str):
        self._set(self._props.target, str(value))

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.FRAME
        return self._format_kind_prop

    @property
    def _props(self) -> HyperlinkProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = HyperlinkProps(
                name="HyperLinkName", target="HyperLinkTarget", url="HyperLinkURL", visited="", unvisited=""
            )
        return self._props_internal_attributes

    # endregion Properties
