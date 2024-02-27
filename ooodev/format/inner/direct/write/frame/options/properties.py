from __future__ import annotations
from typing import cast, overload, TYPE_CHECKING
from typing import Any, Tuple, Type, TypeVar
from enum import Enum
from ooo.dyn.text.writing_mode2 import WritingMode2
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.format.inner.direct.write.para.align.writing_mode import WritingMode
from ooo.dyn.text.writing_mode2 import WritingMode2Enum
from ooodev.meta.deleted_attrib import DeletedAttrib
from ooodev.format.inner.common.props.frame_options_properties import FrameOptionsProperties

if TYPE_CHECKING:
    from ooodev.format.inner.direct.write.para.align.writing_mode import _TWritingMode

_TProperties = TypeVar(name="_TProperties", bound="Properties")


class TextDirectionKind(Enum):
    """
    Describes different writing directions
    """

    LR_TB = WritingMode2.LR_TB  # keep
    """
    Left-to-right (LTR)
    
    Text within lines is written left-to-right.
    Lines and blocks are placed top-to-bottom.
    Typically, this is the writing mode for normal ``alphabetic`` text.
    """
    RL_TB = WritingMode2.RL_TB  # keep
    """
    Right-to-left (RTL).
    
    text within a line are written right-to-left.
    Lines and blocks are placed top-to-bottom.
    Typically, this writing mode is used in Arabic and Hebrew text.
    """
    TB_RL = WritingMode2.TB_RL  # keep
    """
    Right-to-left (vertical).
    
    Text within a line is written top-to-bottom.
    Lines and blocks are placed right-to-left.
    Typically, this writing mode is used in Chinese and Japanese text.
    """
    TB_LR = WritingMode2.TB_LR  # keep
    """
    Left-to-right (vertical).
    
    Text within a line is written top-to-bottom.
    Lines and blocks are placed left-to-right.
    Typically, this writing mode is used in Mongolian text.
    """
    PAGE = WritingMode2.PAGE  # keep, use superordinate object settings
    """
    Use super-ordinate object settings
    
    Obtain writing mode from the current page.
    May not be used in page styles.
    """
    BT_LR = WritingMode2.BT_LR  # keep
    """
    Bottom-to-top, left-to-right (vertical)
    
    Text within a line is written bottom-to-top.
    Lines and blocks are placed left-to-right.
    """

    def __int__(self) -> int:
        return self.value


class TextDirectionMode(WritingMode):
    context = DeletedAttrib()  # type: ignore[misc]

    def __init__(self, mode: TextDirectionKind | None = None) -> None:
        """
        Constructor

        Args:
            mode (TextDirectionKind, optional): Determines the writing direction

        Returns:
            None:
        """
        if mode is None:
            super().__init__()
        else:
            super().__init__(mode=WritingMode2Enum(mode.value))

    # region style methods
    def fmt_mode(self: _TWritingMode, value: TextDirectionKind | None) -> _TWritingMode:
        """
        Gets copy of instance with writing mode set or removed

        Args:
            value (TextDirectionKind | None): mode value

        Returns:
            TextDirectionMode: ``TextDirectionMode`` instance
        """
        if value is None:
            return super().fmt_mode(value=None)
        return super().fmt_mode(value=WritingMode2Enum(value.value))

    # endregion style methods

    @property
    def prop_mode(self) -> TextDirectionKind | None:
        """Gets/Sets writing mode of a paragraph."""
        pv = cast(int, self._get(self._get_property_name()))
        return None if pv is None else TextDirectionKind(pv)

    @prop_mode.setter
    def prop_mode(self, value: TextDirectionKind | None):
        if value is None:
            self._remove(self._get_property_name())
            return
        self._set(self._get_property_name(), value)

    @property
    def default(self) -> TextDirectionMode:  # type: ignore[misc]
        """Gets ``WritingMode`` default."""
        # pylint: disable=unexpected-keyword-arg
        # pylint: disable=protected-access
        try:
            return self._default_inst
        except AttributeError:
            self._default_inst = self.__class__(mode=TextDirectionKind.LR_TB, _cattribs=self._get_internal_cattribs())
            self._default_inst._is_default_inst = True
        return self._default_inst


class Properties(StyleMulti):
    """
    Frame Frame Options Properties.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        editable: bool | None = None,
        printable: bool | None = None,
        txt_direction: TextDirectionKind | None = None,
    ) -> None:
        """
        Constructor

        Args:
            editable (bool, optional): Specifies if Frame is editable in read-only document.
            printable (bool, optional): Specifies if Frame can be printed.
            txt_direction (TextDirectionKind, optional): Specifies text direction.
        """
        super().__init__()
        if editable is not None:
            self.prop_editable = editable
        if printable is not None:
            self.prop_printable = printable
        if txt_direction is not None:
            self._set_txt_direction_mode(txt_direction)

    # region internal methods
    def _set_txt_direction_mode(self, txt_direction: TextDirectionKind | None) -> None:
        # pylint: disable=unexpected-keyword-arg
        # pylint: disable=protected-access
        self._remove_style("text_mode")
        self._del_attribs("_inner_writing_mode")
        if txt_direction is None:
            return
        mode = WritingMode(
            mode=txt_direction,
            _cattribs={
                "_property_name": self._props.write_mode,
                "_supported_services_values": self._supported_services(),
            },
        )
        mode._prop_parent = self
        self._set_style("text_mode", mode, *mode.get_attrs())  # type: ignore

    # endregion internal methods

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.Style", "com.sun.star.text.TextFrame")
        return self._supported_services_values

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion Overrides

    # region Static Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TProperties], obj: object) -> _TProperties: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TProperties], obj: object, **kwargs) -> _TProperties: ...

    @classmethod
    def from_obj(cls: Type[_TProperties], obj: object, **kwargs) -> _TProperties:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Properties: Instance that represents Frame Option Properties.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        inst.prop_txt_direction = TextDirectionKind(
            mProps.Props.get(obj, inst._props.write_mode, TextDirectionKind.LR_TB)
        )
        inst.prop_editable = bool(mProps.Props.get(obj, inst._props.editable, False))
        inst.prop_printable = bool(mProps.Props.get(obj, inst._props.printable, False))
        return inst

    # endregion from_obj()

    # endregion Static Methods

    # region Properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE
        return self._format_kind_prop

    @property
    def prop_editable(self) -> bool | None:
        """Gets/Sets editable value"""
        return self._get(self._props.editable)

    @prop_editable.setter
    def prop_editable(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.editable)
            return
        self._set(self._props.editable, value)

    @property
    def prop_printable(self) -> bool | None:
        """Gets/Sets print value"""
        return self._get(self._props.printable)

    @prop_printable.setter
    def prop_printable(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.printable)
            return
        self._set(self._props.printable, value)

    @property
    def prop_txt_direction(self) -> TextDirectionKind | None:
        """Gets/Sets text direction value"""
        if self.prop_inner_writing_mode is None:
            return None
        pv = self.prop_inner_writing_mode.prop_mode
        return None if pv is None else TextDirectionKind(int(pv))

    @prop_txt_direction.setter
    def prop_txt_direction(self, value: TextDirectionKind | None) -> None:
        self._set_txt_direction_mode(value)

    @property
    def prop_inner_writing_mode(self) -> WritingMode | None:
        try:
            return self._inner_writing_mode  # type: ignore
        except AttributeError:
            self._inner_writing_mode = self._get_style_inst("text_mode")
        return self._inner_writing_mode  # type: ignore

    @property
    def _props(self) -> FrameOptionsProperties:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = FrameOptionsProperties(
                editable="EditInReadonly", printable="Print", write_mode="WritingMode"
            )
        return self._props_internal_attributes

    # endregion Properties
