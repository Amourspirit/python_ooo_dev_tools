from __future__ import annotations
from typing import Any, Tuple, Type, TypeVar, cast, overload
import uno
from ooo.dyn.text.wrap_text_mode import WrapTextMode as WrapTextMode

from .....exceptions import ex as mEx
from .....utils import lo as mLo
from .....utils import props as mProps
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase
from ...common.props.frame_wrap_options_props import FrameWrapOptionsProps

_TOptions = TypeVar(name="_TOptions", bound="Options")


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
        countour: bool | None = None,
        outside: bool | None = None,
        overlap: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            first (bool , optional): Specifies first paragraph.
            background (bool , optional): Specifies in background.
            countour (bool , optional): Specifies contour.
            outside (bool , optional): Specifies contour outside only. ``countour`` must be ``True`` for this parameter to be effective.
            overlap (bool , optional): Specifies allow overlap.
        """
        super().__init__()
        if first is not None:
            self.prop_first = first
        if background is not None:
            self.prop_background = background
        if countour is not None:
            self.prop_coutour = countour
        if outside is not None:
            self.prop_outside = outside
        if overlap is not None:
            self.prop_overlap = overlap

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
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TOptions], obj: object) -> _TOptions:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TOptions], obj: object, **kwargs) -> _TOptions:
        ...

    @classmethod
    def from_obj(cls: Type[_TOptions], obj: object, **kwargs) -> _TOptions:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

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
        opt.prop_coutour = True
        return opt

    @property
    def outisde(self: _TOptions) -> _TOptions:
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
        if pv is None:
            return None
        return not pv

    @prop_background.setter
    def prop_background(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.background)
            return
        self._set(self._props.background, not value)

    @property
    def prop_coutour(self) -> bool | None:
        """Gets/Sets contour value"""
        return self._get(self._props.countour)

    @prop_coutour.setter
    def prop_coutour(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.countour)
            return
        self._set(self._props.countour, value)

    @property
    def prop_outside(self) -> bool | None:
        """Gets/Sets contour outside only value"""
        return self._get(self._props.countour)

    @prop_outside.setter
    def prop_outside(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.countour)
            return
        self._set(self._props.countour, value)

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
                countour="SurroundContour",
                outside="ContourOutside",
                overlap="AllowOverlap",
            )
        return self._props_internal_attributes

    # endregion Prop Properties
