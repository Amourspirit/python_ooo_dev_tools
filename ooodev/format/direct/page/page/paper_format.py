"""
Module for Fill Transparency.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Any, Tuple, cast, Type, TypeVar, NamedTuple

import uno
from .....events.args.cancel_event_args import CancelEventArgs
from .....exceptions import ex as mEx
from .....utils import lo as mLo
from .....utils import props as mProps
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase
from ...common.size_mm import SizeMM as SizeMM
from ....preset.preset_paper_format import PaperFormatKind as PaperFormatKind

from ooo.dyn.awt.size import Size as Size

_TPaperFormat = TypeVar(name="_TPaperFormat", bound="PaperFormat")


class PaperFormat(StyleBase):
    """
    Fill Transparency

    .. versionadded:: 0.9.0
    """

    def __init__(self, size: SizeMM = SizeMM(215.9, 279.4)) -> None:
        """
        Constructor

        Args:
            size (SizeMM, optional): Width and height in ``mm`` units. Defaults to Letter size in Portrait mode.
        """

        super().__init__()
        self.prop_size = size

    # region Overrides
    def _supported_services(self) -> Tuple[str, ...]:
        return ("com.sun.star.style.PageProperties", "com.sun.star.style.PageStyle")

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion Overrides

    # region internal methods

    # endregion internal methods

    # region Static Methods
    @classmethod
    def from_obj(cls: Type[_TPaperFormat], obj: object) -> _TPaperFormat:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Returns:
            Margins: Instance that represents object margins.
        """
        # this nu is only used to get Property Name

        inst = super(PaperFormat, cls).__new__(cls)
        inst.__init__()
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        def set_prop(key: str, clazz: PaperFormat):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if not val is None:
                clazz._set(key, val)

        set_prop("Width", inst)
        set_prop("Height", inst)
        set_prop("Size", inst)
        set_prop("IsLandscape", inst)

        return inst

    @classmethod
    def from_preset(cls: Type[_TPaperFormat], preset: PaperFormatKind, landscape: bool = False) -> _TPaperFormat:
        """
        Gets instance from preset

        Args:
            preset (PaperFormatKind): Preset kind
            landscape (bool, optional): Specifies if the preset is in landscape mode. Defaults to ``False``.

        Returns:
            PaperFormat: Format from preset
        """
        sz = preset.get_size()
        if landscape:
            sz = sz.swap()
        inst = super(PaperFormat, cls).__new__(cls)
        inst.__init__(SizeMM.from_size_mm100(sz))
        return inst

    # endregion Static Methods

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PAGE

    @property
    def prop_size(self) -> SizeMM:
        """Gets/Sets Size value"""
        width = cast(int, self._get("Width"))
        height = cast(int, self._get("Height"))
        return SizeMM.from_mm100(width, height)

    @prop_size.setter
    def prop_size(self, value: SizeMM) -> None:
        size = value.get_size_mm100()
        self._set("Width", size.width)
        self._set("Height", size.height)
        self._set("Size", Size(Width=size.width, Height=size.height))
        if size.width > size.height:
            self._set("IsLandscape", True)
        else:
            self._set("IsLandscape", False)

    @property
    def prop_landscape(self) -> bool:
        """Gets Landscape value"""
        return cast(bool, self._get("IsLandscape"))
