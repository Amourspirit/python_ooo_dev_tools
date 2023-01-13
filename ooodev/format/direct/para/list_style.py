"""
Modele for managing paragraph breaks.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, overload

from ....exceptions import ex as mEx
from ....meta.static_prop import static_prop
from ....utils import lo as mLo
from ....utils import props as mProps
from ...kind.style_kind import StyleKind
from ...style_base import StyleBase
from ...style.kind.style_list_kind import StyleListKind as StyleListKind

# from ...events.args.key_val_cancel_args import KeyValCancelArgs


class ListStyle(StyleBase):
    """
    Paragraph ListStyle

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    # region init

    def __init__(self, list_style: str | StyleListKind | None = None, num_start: int | None = None) -> None:
        """
        Constructor

        Args:
            list_style (str, StyleListKind, optional): List Style.
            num_start (int, optional): Starts with number.
                If ``-1`` then restart numbering at current paragraph is consider to be ``False``.
                If ``-2`` then restart numbering at current paragraph is consider to be ``True``.
                Otherewise, restart numbering is considered to be ``True``.

        Returns:
            None:

        Note:
            If argument ``list_style`` is ``StyleListKind.NONE`` or empty string then ``num_start`` is ignored.
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html

        init_vals = {}
        if not list_style is None:
            # if list_style is StyleListKind and it is StyleListKind.NONE then str will be empty string
            str_style = str(list_style)
            if str_style:
                init_vals["NumberingStyleName"] = str_style
            else:
                init_vals["NumberingStyleName"] = ""
                init_vals["NumberingStartValue"] = -1
                init_vals["ParaIsNumberingRestart"] = False

        if not num_start is None and not "NumberingStartValue" in init_vals:
            # ignore num_start if NumberingStartValue = -1 due to no style
            if num_start == -1:
                init_vals["NumberingStartValue"] = -1
                init_vals["ParaIsNumberingRestart"] = False
            elif num_start < -1:
                init_vals["NumberingStartValue"] = -1
                init_vals["ParaIsNumberingRestart"] = True
            elif num_start >= 0:
                init_vals["NumberingStartValue"] = num_start
                init_vals["ParaIsNumberingRestart"] = True

        super().__init__(**init_vals)

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.ParagraphProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.ParagraphProperties",)

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies break properties to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            None:
        """
        try:
            super().apply(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__}.apply_style(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()

    @staticmethod
    def from_obj(obj: object) -> ListStyle:
        """
        Gets instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support  ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            ListStyle: ``ListStyle`` instance that represents ``obj`` properties.
        """
        inst = ListStyle()
        if not inst._is_valid_service(obj):
            raise mEx.NotSupportedServiceError(inst._supported_services()[0])

        def set_prop(key: str, o: ListStyle):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if not val is None:
                o._set(key, val)

        set_prop("NumberingStyleName", inst)
        set_prop("NumberingStartValue", inst)
        set_prop("ParaIsNumberingRestart", inst)

        return inst

    # endregion methods

    # region Style Methods
    def fmt_list_style(self, value: str | StyleListKind | None) -> ListStyle:
        """
        Gets a copy of instance with before list style set or removed

        Args:
            value (str, StyleListKind, None): List style value.

        Returns:
            ListStyle: List Style instance
        """
        cp = self.copy()
        cp.prop_list_style = value
        return cp

    def fmt_num_start(self, value: int | None) -> ListStyle:
        """
        Gets a copy of instance with before list style set or removed

        Args:
            value (int | None): List style value.
                If ``-1`` then restart numbering at current paragraph is consider to be ``False``.
                If ``-2`` then restart numbering at current paragraph is consider to be ``True``.
                Otherewise, restart numbering is considered to be ``True``.

        Returns:
            ListStyle: List Style instance
        """
        cp = self.copy()
        cp.prop_num_start = value
        return cp

    # endregion Style Methods

    # region Style Properties
    @property
    def restart_numbers(self) -> ListStyle:
        """Gets instance with restart numbers set"""
        cp = self.copy()
        cp.prop_num_start = -1
        cp._set("ParaIsNumberingRestart", True)
        return cp

    # endregion Style Properties

    # region properties
    @property
    def prop_style_kind(self) -> StyleKind:
        """Gets the kind of style"""
        return StyleKind.PARA

    @property
    def prop_list_style(self) -> str | None:
        """Gets/Sets break type"""
        return self._get("NumberingStyleName")

    @prop_list_style.setter
    def prop_list_style(self, value: str | StyleListKind | None) -> None:
        if value is None:
            self._remove("NumberingStyleName")
            return
        str_val = str(value)
        if not str_val:
            # empty string
            self._remove("NumberingStyleName")
            return
        self._set("NumberingStyleName", value)

    @property
    def prop_num_start(self) -> int | None:
        """
        Gets/Sets Starts with number.

        If Less then zero then restart numbering at current paragraph is consider to be ``False``;
        Otherewise; restart numbering is considered to be ``True``.
        """
        return self._get("NumberingStartValue")

    @prop_num_start.setter
    def prop_num_start(self, value: int | None) -> None:
        if value is None:
            self._remove("NumberingStartValue")
            self._remove("ParaIsNumberingRestart")
            return
        if value == -1:
            self._set("NumberingStartValue", -1)
            self._set("ParaIsNumberingRestart", False)
            return
        if value < -1:
            self._set("NumberingStartValue", -1)
            self._set("ParaIsNumberingRestart", True)
        self._set("NumberingStartValue", value)
        self._set("ParaIsNumberingRestart", True)

    @static_prop
    def default() -> ListStyle:  # type: ignore[misc]
        """Gets ``ListStyle`` default. Static Property."""
        if ListStyle._DEFAULT is None:
            ls = ListStyle()
            ls._set("NumberingStyleName", "")
            ls._set("ParaIsNumberingRestart", False)
            ls._set("NumberingStartValue", -1)
            ListStyle._DEFAULT = ls
        return ListStyle._DEFAULT

    # endregion properties
