from __future__ import annotations
from typing import TYPE_CHECKING, overload
from typing import Any, Tuple, Type, TypeVar
from enum import Enum
from ooodev.exceptions import ex as mEx
from ooodev.meta.deleted_enum_meta import DeletedUnoEnumMeta
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.format.inner.common.props.frame_type_anchor_props import FrameTypeAnchorProps

_TAnchor = TypeVar("_TAnchor", bound="Anchor")

if TYPE_CHECKING:
    # this class is only available at design time.
    # When python is running this class will not exist
    class AnchorKind(Enum):
        """
        Anchor Position Enum.
        """

        @property
        def typeName(self) -> str:
            return "com.sun.star.text.TextContentAnchorType"

        AS_CHARACTER = "AS_CHARACTER"
        """
        The object is anchored instead of a character.
        """
        AT_CHARACTER = "AT_CHARACTER"
        """
        The object is anchored to a character.
        """
        AT_PAGE = "AT_PAGE"
        """
        The object is anchored to the page.
        """
        AT_PARAGRAPH = "AT_PARAGRAPH"
        """
        The anchor of the object is set at the top left position of the paragraph.
        """

else:
    # Class takes the place of the above class at runtime.
    # The reason for this to make sure 'AT_FRAME' enum value is excluded
    # Also future proof enum, if later version add new enum values.
    class AnchorKind(
        metaclass=DeletedUnoEnumMeta,
        type_name="com.sun.star.text.TextContentAnchorType",
        name_space="com.sun.star.text",
    ):
        @staticmethod
        def _get_deleted_attribs() -> Tuple[str]:
            return ("AT_FRAME",)


class Anchor(StyleBase):
    """
    Anchor Position Class.

    .. versionadded:: 0.9.0
    """

    def __init__(self, anchor: AnchorKind = AnchorKind.AT_PARAGRAPH) -> None:
        """
        Constructor

        Args:
            anchor (AnchorKind): Specifies the anchor position. Default is ``AnchorKind.AT_PARAGRAPH``
        """
        super().__init__()
        self.prop_anchor = anchor

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.text.TextContent", "com.sun.star.text.Shape")
        return self._supported_services_values

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

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
        # check if obj has AnchorType property
        # Although reported services are com.sun.star.text.TextContent and com.sun.star.text.Shape in the api,
        # some objects do implement the AnchorType property. Such as Drawing Shapes inserted into a Write document draw page.
        return hasattr(obj, "AnchorType")

    # endregion Overrides
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TAnchor], obj: Any) -> _TAnchor: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TAnchor], obj: Any, **kwargs) -> _TAnchor: ...

    @classmethod
    def from_obj(cls: Type[_TAnchor], obj: Any, **kwargs) -> _TAnchor:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Returns:
            Anchor: Instance that represents Frame Anchor Position.
        """
        # this nu is only used to get Property Name

        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        val = mProps.Props.get(obj, inst._props.name)
        inst.prop_anchor = val
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
    def prop_anchor(self) -> AnchorKind:
        """Gets/Sets Anchor position"""
        return self._get(self._props.name)

    @prop_anchor.setter
    def prop_anchor(self, value: AnchorKind) -> None:
        self._set(self._props.name, value)

    @property
    def _props(self) -> FrameTypeAnchorProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = FrameTypeAnchorProps(name="AnchorType")
        return self._props_internal_attributes
