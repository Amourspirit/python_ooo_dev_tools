"""
Module for managing paragraph Drop Caps.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, cast, overload

from ....exceptions import ex as mEx
from ....meta.static_prop import static_prop
from ....utils import props as mProps
from ....utils import lo as mLo
from ...kind.format_kind import FormatKind
from ...style_base import StyleMulti
from ...style.writer.kind.style_char_kind import StyleCharKind as StyleCharKind
from ..structs.drop_cap import DropCap
from ....events.args.key_val_cancel_args import KeyValCancelArgs


class DropCapFmt(DropCap):
    """
    Paragraph Drop Cap

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.ParagraphProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.ParagraphProperties",)

    def _get_property_name(self) -> str:
        return "DropCapFormat"


class DropCaps(StyleMulti):
    """
    Paragraph Drop Caps

    Warning:
        This class uses dispatch commands and is not suitable for use in headless mode.

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    # region init

    def __init__(
        self,
        *,
        count: int = 0,
        spaces: float = 0.0,
        lines: int = 3,
        style: StyleCharKind | str | None = None,
        whole_word: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            count (int): Specifies the number of characters in the drop cap. Must be from ``0`` to ``255``.
            spaces (float): Specifies the distance between the drop cap in the following text (in mm units)
            lines (int): Specifies the number of lines used for a drop cap. Must be from ``0`` to ``255``.
            style (StyleCharKind, str, optional): Specifies the character style name for drop caps.
            whole_word (bool, optional): specifies if Drop Cap is applied to the whole first word.
        Returns:
            None:

        Note:
            If ``count==-1`` then only ``style`` can be updated.
            If ``count==0`` then all other argumnets are ignored and instance set to remove drop caps when ``apply()`` is called.

        Warning:
            This class may uses dispatch commands and may not suitable for use in headless mode.

            Due to LibreOffice bug this class will use a dispatch command if ``style`` is set to empty string.
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html

        # Due seeemingly to a bug setting ``style`` result in style not being written.
        # If style is set to a valid Character Style Name
        # then it will be written successfully.

        # if count == -1 then do not include DropCapFmt. only update style
        # if count = 0 then default to no drop cap values
        if mLo.Lo.bridge_connector.headless:
            mLo.Lo.print("Warning! DropCaps class is not suitable in Headless mode.")
        dc = None
        init_vars = {}

        if count == -1:
            if not style is None:
                init_vars["DropCapCharStyleName"] = str(style)
        elif count == 0:
            # set defatults to apply no drop caps
            whole_word = False
            style = ""
            init_vars["DropCapWholeWord"] = False
            init_vars["DropCapCharStyleName"] = ""
            dc = DropCapFmt(0, 0, 0)
        elif count > 0:
            dc = DropCapFmt(count, round(spaces * 100), lines)
            if not whole_word is None:
                init_vars["DropCapWholeWord"] = whole_word
                if whole_word:
                    # when whole word is set the count must be 1 (for one word)
                    dc.prop_count = 1
            if not style is None:
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
        Resets the cursor at is current position/selection to remove any Drop Caps Formatting.

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

    def on_property_backing_up(self, event_args: KeyValCancelArgs):
        """
        Triggers before each property that is about to be backup up during backup

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        if event_args.key == "DropCapCharStyleName":
            # DropCapCharStyleName will not allow itself to be set if it has empyt string
            # as a value, even though it ia a string and will take a string value of any valid
            # character style.
            if event_args.value is None or event_args.value == "":
                event_args.cancel = True

    def _set_style_dc(self, dc: DropCapFmt | None) -> None:
        if dc is None:
            self._remove_style("drop_cap")
            return
        self._set_style("drop_cap", dc, *dc.get_attrs(), keys={"prop": "DropCapFormat"})

    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.ParagraphProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.ParagraphProperties",)

    @staticmethod
    def from_obj(obj: object) -> DropCaps:
        """
        Gets instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support  ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            DropCaps: ``DropCaps`` instance that represents ``obj`` Drop Caps.
        """
        inst = DropCaps()
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedServiceError(inst._supported_services()[0])
        dc = DropCapFmt.from_obj(obj)
        inst._set_style_dc(dc)

        whole_word = cast(bool, mProps.Props.get(obj, "DropCapWholeWord"))
        style = cast(str, mProps.Props.get(obj, "DropCapCharStyleName"))
        if not whole_word is None:
            inst.set("DropCapWholeWord", whole_word)
        if not style is None:
            inst.set("DropCapCharStyleName", style)
        return inst

    def backup(self, obj: object) -> None:
        """Overrides, No actions are taken"""
        pass

    @overload
    def restore(self, obj: object) -> None:
        ...

    def restore(self, obj: object, clear: bool = False) -> None:
        """
        Restores ``obj`` via dispatc command.

        Args:
            obj (object): Object to restore properties on.

        Returns:
            None:
        """
        self.dispatch_reset()

    # endregion methods

    # region properties
    @property
    def prop_has_backup(self) -> bool:
        """Gets If instantance has backup data. Overrides, Returns ``False``."""
        return True

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA

    @static_prop
    def default() -> DropCaps:  # type: ignore[misc]
        """Gets ``DropCaps`` default. Static Property."""
        if DropCaps._DEFAULT is None:
            inst = DropCaps(count=0)
            DropCaps._DEFAULT = inst
        return DropCaps._DEFAULT

    # endregion properties
