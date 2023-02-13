"""
Module for creating hyperlinks

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import Tuple, overload, Type, TypeVar
from enum import Enum

from .....events.args.cancel_event_args import CancelEventArgs
from .....exceptions import ex as mEx
from .....meta.static_prop import static_prop
from .....utils import lo as mLo
from .....utils import props as mProps
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase

# endregion imports

_THyperlink = TypeVar(name="_THyperlink", bound="Hyperlink")


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


class Hyperlink(StyleBase):
    """
    Hyperlink

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        *,
        name: str | None = None,
        url: str | None = None,
        target: TargetKind = TargetKind.NONE,
        visited_style: str = "Visited Internet Link",
        unvisited_style: str = "Internet link",
    ) -> None:
        """
        Constructor

        Args:
            name (str, optional): Link name.
            url (str, optional): Link Url.
            target (TargetKind, optional): Link target. Defaults to ``TargetKind.NONE``.
            visited_style (str, optional): Link visited style. Defaults to ``Internet link``.
            unvisited_style (str, optional): Link unvisited style. Defaults to ``Visited Internet Link``.

        Returns:
            None:
        """
        init_vals = {
            "HyperLinkTarget": target.value,
            "VisitedCharStyleName": visited_style,
            "UnvisitedCharStyleName": unvisited_style,
        }
        if not name is None:
            init_vals["HyperLinkName"] = name
        if not url is None:
            init_vals["HyperLinkURL"] = url

        super().__init__(**init_vals)

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        return ("com.sun.star.style.CharacterProperties", "com.sun.star.style.CharacterStyle")

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

    # region apply()

    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.CharacterProperties`` service.
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

    @classmethod
    def from_obj(cls: Type[_THyperlink], obj: object) -> _THyperlink:
        """
        Gets hyperlink instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Hyperlink: Hyperlink that represents ``obj`` Hyperlink.
        """
        inst = super(Hyperlink, cls).__new__(cls)
        inst.__init__()
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        inst._set("HyperLinkName", mProps.Props.get(obj, "HyperLinkName"))
        inst._set("HyperLinkURL", mProps.Props.get(obj, "HyperLinkURL"))
        inst._set("HyperLinkTarget", mProps.Props.get(obj, "HyperLinkTarget"))
        inst._set("VisitedCharStyleName", mProps.Props.get(obj, "VisitedCharStyleName"))
        inst._set("UnvisitedCharStyleName", mProps.Props.get(obj, "UnvisitedCharStyleName"))

        return inst

    # endregion methods

    # region Properties

    @property
    def prop_name(self) -> str | None:
        """Gets/Sets name"""
        return self._get("HyperLinkName")

    @prop_name.setter
    def prop_name(self, value: str | None):
        if value is None:
            if self._has("HyperLinkName"):
                self._remove("HyperLinkName")
        else:
            self._set("HyperLinkName", value)

    @property
    def prop_url(self) -> str | None:
        """Gets/Sets url"""
        return self._get("HyperLinkURL")

    @prop_url.setter
    def prop_url(self, value: str | None):
        if value is None:
            if self._has("HyperLinkURL"):
                self._remove("HyperLinkURL")
        else:
            self._set("HyperLinkURL", value)

    @property
    def prop_target(self) -> TargetKind:
        """Gets/Sets target"""
        return TargetKind(self._get("HyperLinkTarget"))

    @prop_target.setter
    def prop_target(self, value: TargetKind):
        self._set("HyperLinkTarget", value.value)

    @property
    def prop_visited_style(self) -> str:
        """Gets/Sets visited style"""
        return self._get("VisitedCharStyleName")

    @prop_visited_style.setter
    def prop_visited_style(self, value: str):
        self._set("VisitedCharStyleName", value)

    @property
    def prop_unvisited_style(self) -> str:
        """Gets/Sets style for links that have not yet been visited"""
        return self._get("UnvisitedCharStyleName")

    @prop_unvisited_style.setter
    def prop_unvisited_style(self, value: str):
        self._set("UnvisitedCharStyleName", value)

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.CHAR

    @static_prop
    def empty() -> Hyperlink:  # type: ignore[misc]
        """Gets Highlight empty. Static Property."""
        try:
            return Hyperlink._EMPTY_INST
        except AttributeError:
            Hyperlink._EMPTY_INST = Hyperlink(name="", url="")
            Hyperlink._EMPTY_INST._is_default_inst = True
        return Hyperlink._EMPTY_INST

    # endregion Properties
