"""
Module for managing paragraph Drop Caps.

.. versionadded:: 0.9.0
"""

# region Imports
from __future__ import annotations
from typing import Any, Tuple, cast, Type, TypeVar, overload

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.direct.structs.drop_cap_struct import DropCapStruct
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.format.writer.style.char.kind.style_char_kind import StyleCharKind
from ooodev.loader import lo as mLo
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_obj import UnitT
from ooodev.utils import props as mProps

# endregion Imports

_TDropCaps = TypeVar(name="_TDropCaps", bound="DropCaps")


class DropCaps(StyleMulti):
    """
    Paragraph Drop Caps

    .. seealso::

        - :ref:`help_writer_format_direct_para_drop_caps`

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        *,
        count: int = 0,
        spaces: float | UnitT = 0.0,
        lines: int = 3,
        style: StyleCharKind | str | None = None,
        whole_word: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            count (int): Specifies the number of characters in the drop cap. Must be from ``0`` to ``255``.
            spaces (float, UnitT): Specifies the distance between the drop cap in the following text
                (in ``mm`` units) or :ref:`proto_unit_obj`.
            lines (int): Specifies the number of lines used for a drop cap. Must be from ``0`` to ``255``.
            style (StyleCharKind, str, optional): Specifies the character style name for drop caps.
            whole_word (bool, optional): specifies if Drop Cap is applied to the whole first word.
        Returns:
            None:

        Note:
            If ``count==-1`` then only ``style`` can be updated.
            If ``count==0`` then all other arguments are ignored and instance set to remove drop caps when
            ``apply()`` is called.

        See Also:

            - :ref:`help_writer_format_direct_para_drop_caps`
        """
        # pylint: disable=unexpected-keyword-arg
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html

        # if count == -1 then do not include DropCapStruct. only update style
        # if count = 0 then default to no drop cap values
        dc = None
        init_vars = {}

        if count == -1:
            if style is not None:
                init_vars["DropCapCharStyleName"] = str(style)
        elif count == 0:
            # set defaults to apply no drop caps
            whole_word = False
            style = ""
            init_vars["DropCapWholeWord"] = False
            init_vars["DropCapCharStyleName"] = ""
            dc = DropCapStruct(count=0, distance=0, lines=0, _cattribs=self._get_cattribs())  # type: ignore
        elif count > 0:
            try:
                dist = spaces.get_value_mm100()  # type: ignore
            except AttributeError:
                dist = UnitConvert.convert_mm_mm100(spaces)  # type: ignore
            dc = DropCapStruct(count=count, distance=dist, lines=lines, _cattribs=self._get_cattribs())  # type: ignore
            if whole_word is not None:
                init_vars["DropCapWholeWord"] = whole_word
                if whole_word:
                    # when whole word is set the count must be 1 (for one word)
                    dc.prop_count = 1
            if style is not None:
                init_vars["DropCapCharStyleName"] = str(style)
        else:
            raise ValueError("Count must not be less then -1")

        super().__init__(**init_vars)
        if dc:
            self._set_style_dc(dc)

    # endregion init

    # region methods

    def dispatch_reset(self) -> None:
        """
        Resets the cursor at is current position/selection to remove any Drop Caps Formatting using a dispatch command.

        Returns:
            None:

        Example:

            .. code-block:: python

                dc = DropCaps(count=1, style=StyleCharKind.DROP_CAPS)
                Write.append_para(cursor=cursor, text="Hello World!", styles=(dc,))
                dc.dispatch_reset()
        """
        drop_cap_args = {
            "FormatDropcap.Lines": 1,
            "FormatDropcap.Count": 1,
            "FormatDropcap.Distance": 0,
            "FormatDropcap.WholeWord": False,
        }
        drop_cap_props = mProps.Props.make_props(**drop_cap_args)
        mLo.Lo.dispatch_cmd("FormatDropcap", drop_cap_props)
        # mLo.Lo.delay(300)
        mLo.Lo.dispatch_cmd("SetDropCapCharStyleName", mProps.Props.make_props(CharStyleName=""))
        # mLo.Lo.delay(300)

        # endregion methods

    # region Overrides
    def on_property_setting(self, source: Any, event_args: KeyValCancelArgs) -> None:
        """
        Triggers for each property that is set

        Args:
            source (Any): Event Source.
            event_args (KeyValueCancelArgs): Event Args
        """
        # sourcery skip: merge-nested-ifs
        if event_args.key == "DropCapCharStyleName":
            # DropCapCharStyleName will not allow itself to be set if it has empty string
            # as a value, even though it ia a string and will take a string value of any valid
            # character style.
            if event_args.value is None or event_args.value == "":
                # instruct Props.set to call set_default()
                event_args.default = True
        super().on_property_setting(source, event_args)

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.ParagraphProperties",
                "com.sun.star.text.TextContent",
                "com.sun.star.style.ParagraphStyle",
            )
        return self._supported_services_values

    # endregion Overrides

    # region internal methods
    def _set_style_dc(self, dc: DropCapStruct | None) -> None:
        if dc is None:
            self._remove_style("drop_cap")
            return
        dc._prop_parent = self
        self._set_style("drop_cap", dc, *dc.get_attrs())

    def _get_cattribs(self) -> dict:
        return {
            "_supported_services_values": self._supported_services(),
            "_property_name": "DropCapFormat",
            "_format_kind_prop": self.prop_format_kind,
        }

    # endregion internal methods

    # region Static Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TDropCaps], obj: Any) -> _TDropCaps: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TDropCaps], obj: Any, **kwargs) -> _TDropCaps: ...

    @classmethod
    def from_obj(cls: Type[_TDropCaps], obj: Any, **kwargs) -> _TDropCaps:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            DropCaps: ``DropCaps`` instance that represents ``obj`` Drop Caps.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        dc = DropCapStruct.from_obj(obj, _cattribs=inst._get_cattribs())
        inst._set_style_dc(dc)

        whole_word = cast(bool, mProps.Props.get(obj, "DropCapWholeWord"))
        style = cast(str, mProps.Props.get(obj, "DropCapCharStyleName"))
        if whole_word is not None:
            inst._set("DropCapWholeWord", whole_word)
        if style is not None:
            inst._set("DropCapCharStyleName", style)
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()
    # endregion Static Methods

    # region properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA | FormatKind.TXT_CONTENT
        return self._format_kind_prop

    @property
    def prop_inner(self) -> DropCapStruct | None:
        """Gets Drop Caps Format instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DropCapStruct, self._get_style_inst("drop_cap"))
        return self._direct_inner

    @property
    def default(self: _TDropCaps) -> _TDropCaps:
        """Gets ``DropCaps`` default."""
        try:
            return self._default_inst
        except AttributeError:
            inst = self.__class__(count=0, _cattribs=self._get_cattribs())  # type: ignore
            inst._is_default_inst = True
            self._default_inst = inst
        return self._default_inst

    # endregion properties
