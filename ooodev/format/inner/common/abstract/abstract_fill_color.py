"""Abstract Fill Color Module"""

# pylint: disable=broad-exception-raised
# pylint: disable=unused-import
# region Import
from __future__ import annotations
from typing import Any, overload, Type, TypeVar

from ooo.dyn.drawing.fill_style import FillStyle

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.exceptions import ex as mEx
from ooodev.utils import color as mColor
from ooodev.utils.color import StandardColor
from ooodev.format.inner.style_base import StyleBase
from ooodev.format.inner.common.props.fill_color_props import FillColorProps

# endregion Import

# LibreOffice seems to have an unresolved bug with Background color.
# https://bugs.documentfoundation.org/show_bug.cgi?id=99125
# see Also: https://forum.openoffice.org/en/forum/viewtopic.php?p=417389&sid=17b21c173e4a420b667b45a2949b9cc5#p417389
# The solution to these issues is to apply FillColor to Paragraph cursors TextParagraph.

_TAbstractColor = TypeVar(name="_TAbstractColor", bound="AbstractColor")


class AbstractColor(StyleBase):
    """
    Paragraph Fill Coloring

    .. versionadded:: 0.9.0
    """

    def __init__(self, color: mColor.Color = StandardColor.AUTO_COLOR) -> None:
        """
        Constructor

        Args:
            color (:py:data:`~.utils.color.Color`, optional): FillColor Color.

        Returns:
            None:
        """
        super().__init__()
        self.prop_color = color

    # region Overrides

    def _on_modifying(self, source: Any, event_args: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event_args)

    # region apply()

    @overload
    def apply(self, obj: Any) -> None:  # type: ignore
        ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): UNO object.

        Returns:
            None:
        """
        super().apply(obj, **kwargs)

    # endregion apply()

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as multi_err:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in multi_err.errors:
                mLo.Lo.print(f"  {err}")

    # endregion Overrides

    # region Static Methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TAbstractColor], obj: Any) -> _TAbstractColor: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TAbstractColor], obj: Any, **kwargs) -> _TAbstractColor: ...

    @classmethod
    def from_obj(cls: Type[_TAbstractColor], obj: Any, **kwargs) -> _TAbstractColor:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Color: ``Color`` instance that represents ``obj`` Color properties.
        """
        new_class = cls(**kwargs)

        # pylint: disable=protected-access
        if not new_class._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        color = mProps.Props.get(obj, new_class._props.color, None)

        result = cls(**kwargs) if color is None else cls(color=color, **kwargs)
        result.set_update_obj(obj)
        return result

    # endregion from_obj()

    # endregion Static Methods

    # region Properties

    @property
    def prop_color(self) -> mColor.Color:
        """Gets/Sets color"""
        return self._get(self._props.color)

    @prop_color.setter
    def prop_color(self, value: mColor.Color):
        if value >= 0:
            self._set(self._props.color, value)
            if self._props.style:
                self._set(self._props.style, FillStyle.SOLID)
            if self._props.bg:
                self._set(self._props.bg, False)
        else:
            self._set(self._props.color, -1)
            if self._props.style:
                self._set(self._props.style, FillStyle.NONE)
            if self._props.bg:
                self._set(self._props.bg, True)

    @property
    def _props(self) -> FillColorProps:
        raise NotImplementedError

    @property
    def default(self: _TAbstractColor) -> _TAbstractColor:
        """Gets Color empty."""
        try:
            return self._default_inst
        except AttributeError:
            # pylint: disable=unexpected-keyword-arg
            self._default_inst = self.__class__(color=-1, _cattribs=self._get_internal_cattribs())  # type: ignore
            # pylint: disable=protected-access
            self._default_inst._is_default_inst = True
        return self._default_inst

    # endregion Properties
