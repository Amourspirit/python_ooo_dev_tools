from __future__ import annotations
from typing import TYPE_CHECKING
from typing import Any, Tuple, Type, TypeVar
from enum import Enum
import uno
from .....exceptions import ex as mEx
from .....meta.deleted_enum_meta import DeletedEnumMeta
from .....utils import lo as mLo
from .....utils import props as mProps
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase
from ...common.props.frame_type_anchor_props import FrameTypeAnchorProps

_TAnchor = TypeVar(name="_TAnchor", bound="Anchor")

if TYPE_CHECKING:
    # this class is only available at design time.
    # When python is running this class will not exist
    class AnchorKind(Enum):
        """
        Anchor Positon Enum.
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
    class AnchorKind(
        metaclass=DeletedEnumMeta, type_name="com.sun.star.text.TextContentAnchorType", name_space="com.sun.star.text"
    ):
        @staticmethod
        def _get_deleted_attribs() -> Tuple[str]:
            return ("AT_FRAME",)


class Anchor(StyleBase):
    """
    Fill Transparency

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
        return ("com.sun.star.style.Style",)

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion Overrides

    @classmethod
    def from_obj(cls: Type[_TAnchor], obj: object) -> _TAnchor:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Returns:
            Anchor: Instance that represents Frame Anchor Position.
        """
        # this nu is only used to get Property Name

        inst = super(Anchor, cls).__new__(cls)
        inst.__init__()
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        val = mProps.Props.get(obj, inst._props.name)
        inst.prop_anchor = val
        return inst

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.DOC | FormatKind.STYLE

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
            return self._props_frame_type_anchor
        except AttributeError:
            self._props_frame_type_anchor = FrameTypeAnchorProps(name="AnchorType")
        return self._props_frame_type_anchor
