from __future__ import annotations
from typing import TYPE_CHECKING
from typing import Any, Tuple, Type, TypeVar
from enum import Enum
import uno
from .....exceptions import ex as mEx
from .....meta.deleted_enum_meta import DeletedUnoEnumMeta
from .....utils import lo as mLo
from .....utils import props as mProps
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase
from ...common.props.frame_options_align_props import FrameOptionsAlignProps

_TAlign = TypeVar(name="_TAlign", bound="Align")

if TYPE_CHECKING:
    # this class is only available at design time.
    # When python is running this class will not exist
    class VertAdjustKind(Enum):
        """
        This enumeration specifies the vertical position of text inside a shape in relation to the shape.
        """

        @property
        def typeName(self) -> str:
            return "com.sun.star.text.TextContentAnchorType"

        BOTTOM = "BOTTOM"
        """
        The connection line leaves the connected object from the bottom,
        The text is positioned below the main line.
        The bottom edge of the text is adjusted to the bottom edge of the shape.
        """
        CENTER = "CENTER"
        """
        The text is centered inside the shape.
        """
        TOP = "TOP"
        """
        The connection line leaves the connected object from the top,
        The text is positioned above the main line.
        The top edge of the text is adjusted to the top edge of the shape.
        """

else:
    # Class takes the place of the above class at runtime.
    # The reason for this to make sure 'BLOCK' enum value is excluded
    class VertAdjustKind(
        metaclass=DeletedUnoEnumMeta,
        type_name="com.sun.star.drawing.TextVerticalAdjust",
        name_space="com.sun.star.drawing",
    ):
        @staticmethod
        def _get_deleted_attribs() -> Tuple[str]:
            return ("BLOCK",)


class Align(StyleBase):
    """
    Frame Vertical Alignment

    .. versionadded:: 0.9.0
    """

    def __init__(self, adjust: VertAdjustKind = VertAdjustKind.TOP) -> None:
        """
        Constructor

        Args:
            adjust (VertAdjustKindl): Specifies Verticial Adjustment. Default ``VertAdjustKind.TOP``
        """
        super().__init__()
        self.prop_adjust = adjust

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
    def from_obj(cls: Type[_TAlign], obj: object) -> _TAlign:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Returns:
            Align: Instance that represents Frame Protection.
        """
        # this nu is only used to get Property Name

        inst = super(Align, cls).__new__(cls)
        inst.__init__()
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        inst.prop_adjust = VertAdjustKind(mProps.Props.get(obj, inst._props.name))
        return inst

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.DOC | FormatKind.STYLE

    @property
    def prop_adjust(self) -> VertAdjustKind:
        """Gets/Sets Adjust value"""
        return self._get(self._props.name)

    @prop_adjust.setter
    def prop_adjust(self, value: VertAdjustKind) -> None:
        self._set(self._props.name, value)

    @property
    def _props(self) -> FrameOptionsAlignProps:
        try:
            return self._props_frame_opts_align
        except AttributeError:
            self._props_frame_opts_align = FrameOptionsAlignProps(name="TextVerticalAdjust")
        return self._props_frame_opts_align
