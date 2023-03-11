from __future__ import annotations
from typing import overload
from typing import Any, Tuple, Type, TypeVar
import uno

from .....events.args.key_val_cancel_args import KeyValCancelArgs
from .....exceptions import ex as mEx
from .....utils import lo as mLo
from .....utils import props as mProps
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase
from ...common.props.frame_options_names_props import FrameOptionsNamesProps

_TNames = TypeVar(name="_TNames", bound="Names")


class Names(StyleBase):
    """
    Frame Name options

    .. versionadded:: 0.9.0
    """

    def __init__(self, name: str | None = None, desc: str | None = None) -> None:
        """
        Constructor

        Args:
            name (str, optional): Specifies frame name.
            desc (str, optional): Specifies frame description.
        """
        # TODO: Implement prev and next on Frame options Names class.
        # prev (str, optional): Specifies previous link.
        # next (str, optional): Specifies next link.
        # prev: str | None = None, next: str | None = None

        # Not able to get prev and next working on "com.sun.star.text.TextFrame" service.
        # Posted issue on Ask: https://ask.libreoffice.org/t/how-to-link-text-frames-in-writer-via-macro/88497
        # The "com.sun.star.text.TextFrame" has the expected properties of "ChainPrevName" and "ChainNextName" but not able to set them.
        # It may be that a "com.sun.star.text.ChainedTextFrame" service needs to be created (see Write.add_text_frame() for ref).
        # However creating "com.sun.star.text.ChainedTextFrame" results in a error.
        # bug posted: https://bugs.documentfoundation.org/show_bug.cgi?id=153825

        super().__init__()
        if name is not None:
            self.prop_name = name
        if desc is not None:
            self.prop_desc = desc
        # if prev is not None:
        #     self.prop_prev = prev
        # if next is not None:
        #     self.prop_next = next

    # region Overrides
    def apply(self, obj: object, **kwargs) -> None:
        super().apply(obj, **kwargs)
        # for some reason setting Name property raises "UnknownPropertyException" when "setPropertyValue()" is used (Which Props.set() uses).
        # However, setting Name via setattr() works fine.
        # for this reason this class cancels setting of Name property and sets it via setattr() here.
        if hasattr(obj, self._props.name):
            setattr(obj, self._props.name, self.prop_name)

    def on_property_setting(self, event_args: KeyValCancelArgs) -> None:
        if event_args.key == self._props.name:
            event_args.cancel = True
            event_args.handled = True
            # see bug specified in apply() method.
        super().on_property_setting(event_args)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.text.TextFrame", "com.sun.star.text.ChainedTextFrame")
            # This should be ChainedTextFrame only, however there seems to be a bug in LibreOffice.
            # Even though an object is passed in that has ChainedTextFrame properties it does not support ChainedTextFrame service,
            # but does suppor TextFrame service.
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
    def from_obj(cls: Type[_TNames], obj: object) -> _TNames:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TNames], obj: object, **kwargs) -> _TNames:
        ...

    @classmethod
    def from_obj(cls: Type[_TNames], obj: object, **kwargs) -> _TNames:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

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
    def prop_name(self) -> SystemError | None:
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
    # @property
    # def prop_prev(self) -> bool | None:
    #     """Gets/Sets frame previous link"""
    #     return self._get(self._props.prev)

    # @prop_prev.setter
    # def prop_prev(self, value: bool | None) -> None:
    #     if value is None:
    #         self._remove(self._props.prev)
    #         return
    #     self._set(self._props.prev, value)

    # @property
    # def prop_next(self) -> bool | None:
    #     """Gets/Sets frame next link"""
    #     return self._get(self._props.next)

    # @prop_next.setter
    # def prop_next(self, value: bool | None) -> None:
    #     if value is None:
    #         self._remove(self._props.next)
    #         return
    #     self._set(self._props.next, value)

    @property
    def _props(self) -> FrameOptionsNamesProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = FrameOptionsNamesProps(
                name="Name",
                desc="Description",
                prev="",  # ChainPrevName not working
                next="",  # ChainNextName not working
            )
        return self._props_internal_attributes
