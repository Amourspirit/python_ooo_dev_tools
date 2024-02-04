from __future__ import annotations
from typing import overload
from typing import Any, cast, Tuple, Type, TypeVar, TYPE_CHECKING

# from ....style_base import StyleBase
from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.common.abstract.abstract_document import AbstractDocument
from ooodev.format.inner.common.props.frame_options_names_props import FrameOptionsNamesProps

if TYPE_CHECKING:
    from com.sun.star.text import ChainedTextFrame  # service

_TNames = TypeVar(name="_TNames", bound="Names")


class Names(AbstractDocument):
    """
    Frame Name options

    .. versionadded:: 0.9.0
    """

    def __init__(
        self, *, name: str | None = None, desc: str | None = None, prev: str | None = None, next: str | None = None
    ) -> None:
        """
        Constructor

        Args:
            name (str, optional): Specifies frame name.
            desc (str, optional): Specifies frame description.
            prev (str, optional): Specifies previous link.
            next (str, optional): Specifies next link.

        Returns:
            None:

        See Also:
            LibreOffice Help <Inserting, Editing and Linking Frames `https://help.libreoffice.org/latest/en-GB/text/swriter/guide/text_frame.html?DbPAR=WRITER#bm_id3149487`>__

        Note:
            Flowing text from one text frame to another, via ``prev`` and ``next`` required the text frame
            being flow to not contain text. If the frame to flow to contains text then ``prev`` and ``next``
            do not have any effect.
        """
        # Now Working, able to get prev and next working on "com.sun.star.text.TextFrame" service.
        # Posted issue on Ask: https://ask.libreoffice.org/t/how-to-link-text-frames-in-writer-via-macro/88497
        # If Text is added to both frames that are to be chained together then
        # LO will not chain them.
        # The Frame.ChainNextName and Frame.ChainPrevName properties cannot be set. and do not raise any error.
        # This is the default behavior and makes sense.
        # If the frame to flow to has text already then previous frame cannot flow to it.

        super().__init__()
        if name is not None:
            self.prop_name = name
        if desc is not None:
            self.prop_desc = desc
        if prev is not None:
            self.prop_prev = prev
        if next is not None:
            self.prop_next = next

    # region Overrides
    def apply(self, obj: object, **kwargs) -> None:
        super().apply(obj, **kwargs)
        # for some reason setting Name property raises "UnknownPropertyException" when "setPropertyValue()" is used (Which Props.set() uses).
        # However, setting Name via setattr() works fine.
        # for this reason this class cancels setting of Name property and sets it via setattr() here.
        name = self.prop_name
        if not name:
            return
        obj_name = cast(str, getattr(obj, self._props.name, None))
        if obj_name is None:
            return
        if obj_name != name:
            # only set name is it is different
            setattr(obj, self._props.name, name)

        try:
            p_prev = self.prop_prev
            p_next = self.prop_next
        except mEx.DeletedAttributeError:
            # attributes not used in a child class.
            return

        if p_prev is None and p_next is None:
            return

        frames = self.get_text_frames()
        if frames is None:
            return
        if self.prop_name is None:
            return
        if not frames.hasByName(self.prop_name):
            return
        this_frame = cast("ChainedTextFrame", frames.getByName(self.prop_name))
        if p_prev is not None:
            this_frame.ChainPrevName = p_prev
        if p_next is not None:
            this_frame.ChainNextName = p_next

    def on_property_setting(self, source: Any, event_args: KeyValCancelArgs) -> None:
        skip = (self._props.name, self._props.prev, self._props.next)
        if event_args.key in skip:
            event_args.cancel = True
            event_args.handled = True
            # see bug specified in apply() method.
        super().on_property_setting(source, event_args)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.text.TextFrame", "com.sun.star.text.ChainedTextFrame")
            # This should be ChainedTextFrame only, however there seems to be a bug in LibreOffice.
            # Even though an object is passed in that has ChainedTextFrame properties it does not support ChainedTextFrame service,
            # but does support TextFrame service.
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
    def from_obj(cls: Type[_TNames], obj: Any) -> _TNames: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TNames], obj: Any, **kwargs) -> _TNames: ...

    @classmethod
    def from_obj(cls: Type[_TNames], obj: Any, **kwargs) -> _TNames:
        """
        Gets instance from object

        Args:
            obj (Any): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Names: Instance that represents Frame Name options.
        """
        # this nu is only used to get Property Name

        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        for prop_name in inst._props:
            if prop_name:
                try:
                    inst._set(prop_name, mProps.Props.get(obj, prop_name))
                except mEx.PropertyNotFoundError:
                    # there is a bug. See apply()
                    if hasattr(obj, prop_name):
                        inst._set(prop_name, getattr(obj, prop_name))
                    else:
                        raise

        # prev, next not currently working
        return inst

    # endregion from_obj()
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.FRAME
        return self._format_kind_prop

    @property
    def prop_name(self) -> str | None:
        """Gets/Sets name"""
        return self._get(self._props.name)

    @prop_name.setter
    def prop_name(self, value: str | None) -> None:
        if value is None:
            self._remove(self._props.name)
            return
        self._set(self._props.name, value)

    @property
    def prop_desc(self) -> str | None:
        """Gets/Sets description"""
        return self._get(self._props.desc)

    @prop_desc.setter
    def prop_desc(self, value: str | None) -> None:
        if value is None:
            self._remove(self._props.desc)
            return
        self._set(self._props.desc, value)

    # Prev and Next not currently working
    @property
    def prop_prev(self) -> str | None:
        """Gets/Sets frame previous link"""
        return self._get(self._props.prev)

    @prop_prev.setter
    def prop_prev(self, value: str | None) -> None:
        if value is None:
            self._remove(self._props.prev)
            return
        self._set(self._props.prev, value)

    @property
    def prop_next(self) -> str | None:
        """Gets/Sets frame next link"""
        return self._get(self._props.next)

    @prop_next.setter
    def prop_next(self, value: str | None) -> None:
        if value is None:
            self._remove(self._props.next)
            return
        self._set(self._props.next, value)

    @property
    def _props(self) -> FrameOptionsNamesProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = FrameOptionsNamesProps(
                name="Name",
                desc="Description",
                prev="ChainPrevName",
                next="ChainNextName",
            )
        return self._props_internal_attributes
