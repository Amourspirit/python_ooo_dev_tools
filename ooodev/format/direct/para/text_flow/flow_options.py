"""
Modele for managing paragraph Text Flow options.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, cast, overload, Type, TypeVar

from .....events.args.cancel_event_args import CancelEventArgs
from .....exceptions import ex as mEx
from .....meta.static_prop import static_prop
from .....utils import lo as mLo
from .....utils import props as mProps
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase
from .....events.args.key_val_cancel_args import KeyValCancelArgs

_TFlowOptions = TypeVar(name="_TFlowOptions", bound="FlowOptions")


class FlowOptions(StyleBase):
    """
    Paragraph Text Flow Options

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        *,
        orphans: int | None = None,
        widows: int | None = None,
        keep: bool | None = None,
        no_split: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            orphans (int, optional): Number of Orphan Control Lines.
            widows (int, optional): Number Widow Control Lines.
            keep (bool, optional): Keep with next paragraph.
            no_split (bool, optional): Do not split paragraph.
        Returns:
            None:

        Note:
            When ``orphans`` or ``Windows`` argument is presenent then the ``no_split`` has no effect.
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        init_vals = {}

        if not orphans is None:
            if orphans < 2 or orphans > 9:
                raise ValueError("orphans must be from 2 to 9")
            init_vals["ParaOrphans"] = orphans

        if not widows is None:
            if widows < 2 or widows > 9:
                raise ValueError("windows must be from 2 to 9")
            init_vals["ParaWidows"] = widows

        if not keep is None:
            init_vals["ParaKeepTogether"] = keep

        # no split requires orphans and  windows be none
        if orphans is None and widows is None and not no_split is None:
            init_vals["ParaSplit"] = not no_split

        super().__init__(**init_vals)

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        return ("com.sun.star.style.ParagraphProperties", "com.sun.star.style.ParagraphStyle")

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies writing mode to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            None:
        """
        try:
            super().apply(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()

    @classmethod
    def from_obj(cls: Type[_TFlowOptions], obj: object) -> _TFlowOptions:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            WritingMode: ``FlowOptions`` instance that represents ``obj`` writing mode.
        """
        inst = super(FlowOptions, cls).__new__(cls)
        inst.__init__()
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        def set_prop(key: str, indent: FlowOptions):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if not val is None:
                indent._set(key, val)

        set_prop("ParaOrphans", inst)
        set_prop("ParaWidows", inst)
        set_prop("ParaKeepTogether", inst)
        set_prop("ParaSplit", inst)
        return inst

    def on_property_setting(self, event_args: KeyValCancelArgs):
        """
        Subscribe to property setting events

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        if event_args.key == "ParaSplit" and self.prop_no_split is True:
            # ParaSplit default is True, prop_no_split controls this property and is inverse
            # prop_no_split is True so attempting to set ParaSplit False
            # only allow ParaSplit to be set to False if orphans and windows are not present.
            if self._has("ParaOrphans") or self._has("ParaWidows"):
                # orphans or windows are present
                event_args.value = True

    # endregion methods

    # region style methods
    def fmt_orphans(self: _TFlowOptions, value: int | None) -> _TFlowOptions:
        """
        Gets a copy of instance with orphans set or removed

        Args:
            value (float | None): orphans value.

        Returns:
            FlowOptions: FlowOptions instance
        """
        cp = self.copy()
        cp.prop_orphans = value
        return cp

    def fmt_widows(self: _TFlowOptions, value: int | None) -> _TFlowOptions:
        """
        Gets a copy of instance with widows set or removed

        Args:
            value (float | None): widows value (in mm units).

        Returns:
            FlowOptions: FlowOptions instance
        """
        cp = self.copy()
        cp.prop_widows = value
        return cp

    def fmt_keep(self: _TFlowOptions, value: bool | None) -> _TFlowOptions:
        """
        Gets a copy of instance with keep set or removed

        Args:
            value (bool | None): keep value.

        Returns:
            FlowOptions: FlowOptions instance
        """
        cp = self.copy()
        cp.prop_keep = value
        return cp

    def fmt_no_split(self: _TFlowOptions, value: bool | None) -> _TFlowOptions:
        """
        Gets a copy of instance with no split set or removed

        Args:
            value (bool | None): no split value.

        Returns:
            FlowOptions: FlowOptions instance
        """
        cp = self.copy()
        cp.prop_keep = value
        return cp

    # endregion style methods

    # region Style Properties
    @property
    def keep(self: _TFlowOptions) -> _TFlowOptions:
        """Gets copy of instance with keep set"""
        cp = self.copy()
        cp.prop_keep = True
        return cp

    @property
    def no_split(self: _TFlowOptions) -> _TFlowOptions:
        """Gets copy of instance with no split set"""
        cp = self.copy()
        cp.prop_no_split = True
        return cp

    # endregion Style Properties

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA

    @property
    def prop_orphans(self) -> int | None:
        """Gets/Sets Number of Orphan Control Lines."""
        return self._get("ParaOrphans")

    @prop_orphans.setter
    def prop_orphans(self, value: int | None):
        if value is None:
            self._remove("ParaOrphans")
            return
        if value < 2 or value > 9:
            raise ValueError("orphans must be from 2 to 9")
        self._set("ParaOrphans", value)

    @property
    def prop_widows(self) -> int | None:
        """Gets/Sets Number Widow Control Lines."""
        return self._get("ParaWidows")

    @prop_widows.setter
    def prop_widows(self, value: int | None):
        if value is None:
            self._remove("ParaWidows")
            return
        if value < 2 or value > 9:
            raise ValueError("windows must be from 2 to 9")
        self._set("ParaWidows", value)

    @property
    def prop_keep(self) -> bool | None:
        """Gets/Sets Keep with next paragraph."""
        return self._get("ParaKeepTogether")

    @prop_keep.setter
    def prop_keep(self, value: bool | None):
        if value is None:
            self._remove("ParaKeepTogether")
            return
        self._set("ParaKeepTogether", value)

    @property
    def prop_no_split(self) -> bool | None:
        """Gets/Sets Do not split paragraph"""
        pv = cast(bool, self._get("ParaSplit"))
        if pv is None:
            return None
        return not pv

    @prop_no_split.setter
    def prop_no_split(self, value: bool | None):
        if value is None:
            self._remove("ParaSplit")
            return
        self._set("ParaSplit", not value)

    @static_prop
    def default() -> FlowOptions:  # type: ignore[misc]
        """Gets ``FlowOptions`` default. Static Property."""
        try:
            return FlowOptions._DEFAULT_INST
        except AttributeError:
            flo = FlowOptions()
            flo._set("ParaOrphans", 2)
            flo._set("ParaWidows", 2)
            flo._set("ParaSplit", True)
            flo._set("ParaKeepTogether", False)
            flo._is_default_inst = True
            FlowOptions._DEFAULT_INST = flo
        return FlowOptions._DEFAULT_INST

    # endregion properties
