# region Import
from __future__ import annotations
from typing import Any, Tuple, Type, TypeVar, cast, overload

from ooodev.exceptions import ex as mEx
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.format.inner.common.props.frame_wrap_options_props import FrameWrapOptionsProps

# endregion Import

_TOptions = TypeVar("_TOptions", bound="Options")


class Options(StyleBase):
    """
    Frame Vertical Alignment

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        first: bool | None = None,
        background: bool | None = None,
        contour: bool | None = None,
        outside: bool | None = None,
        overlap: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            first (bool , optional): Specifies first paragraph.
            background (bool , optional): Specifies in background.
            contour (bool , optional): Specifies contour.
            outside (bool , optional): Specifies contour outside only. ``contour`` must be ``True`` for this parameter to be effective.
            overlap (bool , optional): Specifies allow overlap.
        """
        super().__init__()
        if first is not None:
            self.prop_first = first
        if background is not None:
            self.prop_background = background
        if contour is not None:
            self.prop_contour = contour
        if outside is not None:
            self.prop_outside = outside
        if overlap is not None:
            self.prop_overlap = overlap

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.Style",
                "com.sun.star.text.BaseFrame",
                "com.sun.star.text.TextEmbeddedObject",
                "com.sun.star.text.TextFrame",
                "com.sun.star.text.TextGraphicObject",
            )
        return self._supported_services_values

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
        # check if obj has matching property
        # Some objects such as 'com.sun.star.drawing.shape' sometime support this style.
        # Such is the case when a shape is added to a Writer drawing page.
        # Assume if on attribute matches then it all is a match.
        return hasattr(obj, self._props.contour)

    # endregion Overrides
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TOptions], obj: Any) -> _TOptions: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TOptions], obj: Any, **kwargs) -> _TOptions: ...

    @classmethod
    def from_obj(cls: Type[_TOptions], obj: Any, **kwargs) -> _TOptions:
        """
        Gets instance from object

        Args:
            obj (Any): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Options: Instance that represents Frame Wrap Settings.
        """
        # this nu is only used to get Property Name

        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        for prop in inst._props:
            inst._set(prop, mProps.Props.get(obj, prop, False))
        return inst

    # endregion from_obj()

    # region Style Properties
    @property
    def first(self: _TOptions) -> _TOptions:
        """Gets copy of instance with first paragraph set to ``True``"""
        opt = self.copy()
        opt.prop_first = True
        return opt

    @property
    def background(self: _TOptions) -> _TOptions:
        """Gets copy of instance with in background set to ``True``"""
        opt = self.copy()
        opt.prop_background = True
        return opt

    @property
    def contour(self: _TOptions) -> _TOptions:
        """Gets copy of instance with contour set to ``True``"""
        opt = self.copy()
        opt.prop_contour = True
        return opt

    @property
    def outside(self: _TOptions) -> _TOptions:
        """Gets copy of instance with contour outside only set to ``True``"""
        opt = self.copy()
        opt.prop_outside = True
        return opt

    @property
    def overlap(self: _TOptions) -> _TOptions:
        """Gets copy of instance with allow overlap set to ``True``"""
        opt = self.copy()
        opt.prop_overlap = True
        return opt

    # endregion Style Properties

    # region Prop Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE
        return self._format_kind_prop

    @property
    def prop_first(self) -> bool | None:
        """Gets/Sets first paragraph value"""
        return self._get(self._props.first_para)

    @prop_first.setter
    def prop_first(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.first_para)
            return
        self._set(self._props.first_para, value)

    @property
    def prop_background(self) -> bool | None:
        """Gets/Sets background value"""
        # background (Opaque) is stored in toggled mode When background is True, Opaque is False
        pv = cast(bool, self._get(self._props.background))
        return None if pv is None else not pv

    @prop_background.setter
    def prop_background(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.background)
            return
        self._set(self._props.background, not value)

    @property
    def prop_contour(self) -> bool | None:
        """Gets/Sets contour value"""
        return self._get(self._props.contour)

    @prop_contour.setter
    def prop_contour(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.contour)
            return
        self._set(self._props.contour, value)

    @property
    def prop_outside(self) -> bool | None:
        """Gets/Sets contour outside only value"""
        return self._get(self._props.contour)

    @prop_outside.setter
    def prop_outside(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.contour)
            return
        self._set(self._props.contour, value)

    @property
    def prop_overlap(self) -> bool | None:
        """Gets/Sets allow overlap value"""
        return self._get(self._props.outside)

    @prop_overlap.setter
    def prop_overlap(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.outside)
            return
        self._set(self._props.outside, value)

    @property
    def _props(self) -> FrameWrapOptionsProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            # background (Opaque) is stored in toggled mode When background is True, Opaque is False
            self._props_internal_attributes = FrameWrapOptionsProps(
                first_para="SurroundAnchorOnly",
                background="Opaque",
                contour="SurroundContour",
                outside="ContourOutside",
                overlap="AllowOverlap",
            )
        return self._props_internal_attributes

    # endregion Prop Properties
