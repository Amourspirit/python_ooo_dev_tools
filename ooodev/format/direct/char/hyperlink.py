"""
Module for creating hyperlinks

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import Tuple, overload
from enum import Enum


from ....exceptions import ex as mEx
from ....meta.static_prop import static_prop
from ....utils import lo as mLo
from ....utils import props as mProps
from ...kind.style_kind import StyleKind
from ...style_base import StyleBase

# endregion imports


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

    _EMPTY = None

    # region init

    def __init__(
        self,
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
        """
        Gets a tuple of supported services (``com.sun.star.style.CharacterProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.CharacterProperties",)

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
            mLo.Lo.print(f"{self.__class__}.apply_style(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")
        return None

    # endregion apply()

    @staticmethod
    def from_obj(obj: object) -> Hyperlink:
        """
        Gets hyperlink instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.CharacterProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support  ``com.sun.star.style.CharacterProperties`` service.

        Returns:
            Hyperlink: Hyperlink that represents ``obj`` Hyperlink.
        """
        inst = Hyperlink()
        if inst._supported_services():
            inst._set("HyperLinkName", mProps.Props.get(obj, "HyperLinkName"))
            inst._set("HyperLinkURL", mProps.Props.get(obj, "HyperLinkURL"))
            inst._set("HyperLinkTarget", mProps.Props.get(obj, "HyperLinkTarget"))
            inst._set("VisitedCharStyleName", mProps.Props.get(obj, "VisitedCharStyleName"))
            inst._set("UnvisitedCharStyleName", mProps.Props.get(obj, "UnvisitedCharStyleName"))
        else:
            raise mEx.NotSupportedServiceError(inst._supported_services()[0])
        return inst

    # endregion methods

    # region Properties

    @property
    def name(self) -> str | None:
        """Gets/Sets name"""
        return self._get("HyperLinkName")

    @name.setter
    def name(self, value: str | None):
        if value is None:
            if self._has("HyperLinkName"):
                self._remove("HyperLinkName")
        else:
            self._set("HyperLinkName", value)

    @property
    def url(self) -> str | None:
        """Gets/Sets url"""
        return self._get("HyperLinkURL")

    @url.setter
    def url(self, value: str | None):
        if value is None:
            if self._has("HyperLinkURL"):
                self._remove("HyperLinkURL")
        else:
            self._set("HyperLinkURL", value)

    @property
    def target(self) -> TargetKind:
        """Gets/Sets target"""
        return TargetKind(self._get("HyperLinkTarget"))

    @target.setter
    def target(self, value: TargetKind):
        self._set("HyperLinkTarget", value.value)

    @property
    def visited_style(self) -> str:
        """Gets/Sets visited style"""
        return self._get("VisitedCharStyleName")

    @visited_style.setter
    def visited_style(self, value: str):
        self._set("VisitedCharStyleName", value)

    @property
    def unvisited_style(self) -> str:
        """Gets/Sets style for links that have not yet been visited"""
        return self._get("UnvisitedCharStyleName")

    @unvisited_style.setter
    def unvisited_style(self, value: str):
        self._set("UnvisitedCharStyleName", value)

    @property
    def prop_style_kind(self) -> StyleKind:
        """Gets the kind of style"""
        return StyleKind.CHAR

    @static_prop
    def empty() -> Hyperlink:  # type: ignore[misc]
        """Gets Highlight empty. Static Property."""
        if Hyperlink._EMPTY is None:
            Hyperlink._EMPTY = Hyperlink(name="", url="")
        return Hyperlink._EMPTY

    # endregion Properties
