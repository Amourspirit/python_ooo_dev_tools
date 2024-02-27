"""
Module for managing paragraph Text Flow options.

.. versionadded:: 0.9.0
"""

# region Import
from __future__ import annotations
from typing import Any, Tuple, cast, overload, Type, TypeVar

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs

# endregion Import

_TFlowOptions = TypeVar(name="_TFlowOptions", bound="FlowOptions")


class FlowOptions(StyleBase):
    """
    Paragraph Text Flow Options

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. seealso::

        - :ref:`help_writer_format_direct_para_text_flow`

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
            When ``orphans`` or ``Windows`` argument is present then the ``no_split`` has no effect.

        See Also:

            - :ref:`help_writer_format_direct_para_text_flow`
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        init_vals = {}

        if orphans is not None:
            if orphans < 2 or orphans > 9:
                raise ValueError("orphans must be from 2 to 9")
            init_vals["ParaOrphans"] = orphans

        if widows is not None:
            if widows < 2 or widows > 9:
                raise ValueError("windows must be from 2 to 9")
            init_vals["ParaWidows"] = widows

        if keep is not None:
            init_vals["ParaKeepTogether"] = keep

        # no split requires orphans and  windows be none
        if orphans is None and widows is None and no_split is not None:
            init_vals["ParaSplit"] = not no_split

        super().__init__(**init_vals)

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.ParagraphProperties",
                "com.sun.star.style.ParagraphStyle",
            )
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    # region apply()
    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:
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

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TFlowOptions], obj: Any) -> _TFlowOptions: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TFlowOptions], obj: Any, **kwargs) -> _TFlowOptions: ...

    @classmethod
    def from_obj(cls: Type[_TFlowOptions], obj: Any, **kwargs) -> _TFlowOptions:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            WritingMode: ``FlowOptions`` instance that represents ``obj`` writing mode.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        def set_prop(key: str, indent: FlowOptions):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if val is not None:
                indent._set(key, val)

        set_prop("ParaOrphans", inst)
        set_prop("ParaWidows", inst)
        set_prop("ParaKeepTogether", inst)
        set_prop("ParaSplit", inst)
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()

    def on_property_setting(self, source: Any, event_args: KeyValCancelArgs):
        """
        Subscribe to property setting events

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        # sourcery skip: merge-nested-ifs
        if event_args.key == "ParaSplit" and self.prop_no_split is True:
            # ParaSplit default is True, prop_no_split controls this property and is inverse
            # prop_no_split is True so attempting to set ParaSplit False
            # only allow ParaSplit to be set to False if orphans and windows are not present.
            if self._has("ParaOrphans") or self._has("ParaWidows"):
                # orphans or windows are present
                event_args.value = True
        super().on_property_setting(source, event_args)

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
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA
        return self._format_kind_prop

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
        return None if pv is None else not pv

    @prop_no_split.setter
    def prop_no_split(self, value: bool | None):
        if value is None:
            self._remove("ParaSplit")
            return
        self._set("ParaSplit", not value)

    @property
    def default(self: _TFlowOptions) -> _TFlowOptions:
        """Gets ``FlowOptions`` default."""
        # pylint: disable=protected-access
        # pylint: disable=unexpected-keyword-arg
        try:
            return self._default_inst
        except AttributeError:
            flo = self.__class__(_cattribs=self._get_internal_cattribs())
            flo._set("ParaOrphans", 2)
            flo._set("ParaWidows", 2)
            flo._set("ParaSplit", True)
            flo._set("ParaKeepTogether", False)
            flo._is_default_inst = True
            self._default_inst = flo
        return self._default_inst

    # endregion properties
