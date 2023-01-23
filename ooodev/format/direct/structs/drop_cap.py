"""
Modele for addin paragraph Drop Cap.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Dict, Tuple, cast, overload

import uno
from ....events.event_singleton import _Events
from ....exceptions import ex as mEx
from ....utils import props as mProps
from ....utils.data_type.byte import Byte
from ....utils.type_var import T
from ...kind.format_kind import FormatKind
from ...style_base import StyleBase, EventArgs, CancelEventArgs, FormatNamedEvent


from ooo.dyn.style.drop_cap_format import DropCapFormat


class DropCap(StyleBase):
    """
    Paragraph Drop Cap

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(self, count: int = 0, distance: int = 0, lines: int = 0) -> None:
        """
        Constructor

        Args:
            count (int): Specifies the number of characters in the drop cap. Must be from ``0`` to ``255``. Defaults to ``0``
            distance (int): Specifies the distance between the drop cap in the following text. Defaults to ``0``
            lines (int): Specifies the number of lines used for a drop cap. Must be from ``0`` to ``255``. Defaults to ``0``

        Returns:
            None:

        Note:
            If argument ``type`` is ``None`` then all other argument are ignored
        """
        bcount = Byte(count)
        count = bcount.value
        blines = Byte(lines)
        lines = blines.value
        if distance < 0:
            raise ValueError("distance arg must be a positive number.")

        init_vals = {"Count": count, "Distance": distance, "Lines": lines}

        super().__init__(**init_vals)

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services.
        This is an empty value for this class but may be different for child classes.

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ()

    def _get_property_name(self) -> str:
        return "DropCapFormat"

    def copy(self: T) -> T:
        nu = super(DropCap, self.__class__).__new__(self.__class__)
        nu.__init__()
        if self._dv:
            nu._update(self._dv)
        return nu

    def get_attrs(self) -> Tuple[str, ...]:
        return (self._get_property_name(),)

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    @overload
    def apply(self, obj: object, keys: Dict[str, str]) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies tab properties to ``obj``

        If a DropCap instance with the same position is existing it is updated;
        Otherwise, a new DropCap is added.

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.
            keys (Dict[str, str], optional): Property key, value items that map properties.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYED` :eventref:`src-docs-event`

        Returns:
            None:
        """
        if not self._is_valid_obj(obj):
            # will not apply on this class but may apply on child classes
            self._print_not_valid_obj("apply()")
            return

        cargs = CancelEventArgs(source=f"{self.apply.__qualname__}")
        cargs.event_data = self
        self.on_applying(cargs)
        if cargs.cancel:
            return
        _Events().trigger(FormatNamedEvent.STYLE_APPLYING, cargs)
        if cargs.cancel:
            return

        keys = {"prop": self._get_property_name()}
        if "keys" in kwargs:
            keys.update(kwargs["keys"])
        key = keys["prop"]
        dcf = self.get_drop_cap_format()
        mProps.Props.set(obj, **{key: dcf})
        eargs = EventArgs.from_args(cargs)
        self.on_applied(eargs)
        _Events().trigger(FormatNamedEvent.STYLE_APPLIED, eargs)

        # mProps.Props.set(obj, **{key: tuple(tss_lst)})

    # endregion apply()

    def get_drop_cap_format(self) -> DropCapFormat:
        """
        Gets drop cap foramt for instance

        Returns:
            DropCapFormat: ``DropCapFormat`` instance
        """
        return DropCapFormat(Lines=self._get("Lines"), Count=self._get("Count"), Distance=self._get("Distance"))

    @classmethod
    def from_obj(cls, obj: object) -> DropCap:
        """
        Gets instance from object

        Args:
            obj (object): UNO object

        Raises:
            PropertyNotFoundError: If ``obj`` does not have required property

        Returns:
            DropCap: ``DropCap`` instance that represents ``obj`` Drop cap format properties.
        """
        # this nu is only used to get Property Name
        nu = super(DropCap, cls).__new__(cls)
        nu.__init__()
        prop_name = nu._get_property_name()

        dcf = cast(DropCapFormat, mProps.Props.get(obj, prop_name), None)
        if dcf is None:
            raise mEx.PropertyNotFoundError(prop_name, f"from_obj() obj as no {prop_name} property")

        return cls.from_drop_cap_format(dcf)

    @classmethod
    def from_drop_cap_format(cls, dcf: DropCapFormat) -> DropCap:
        """
        Converts a DropCap Stop instance to a DropCap

        Args:
            ts (TabStop): DropCap stop

        Returns:
            DropCap: DropCap set with DropCap Stop properties
        """
        inst = super(DropCap, cls).__new__(cls)
        inst.__init__(count=dcf.Count, distance=dcf.Distance, lines=dcf.Lines)
        return inst

    # endregion methods

    # region dunder methods
    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, DropCap):
            return (
                self._get("Count") == oth._get("Count")
                and self._get("Distance") == oth._get("Distance")
                and self._get("Lines") == oth._get("Lines")
            )
        if hasattr(oth, "typeName") and getattr(oth, "typeName") == "com.sun.star.style.DropCapFormat":
            dcf = cast(DropCapFormat, oth)
            return (
                self._get("Count") == dcf.Count
                and self._get("Distance") == dcf.Distance
                and self._get("Lines") == dcf.Lines
            )
        return NotImplemented

    # endregion dunder methods

    # region format methods
    def fmt_count(self, value: int) -> DropCap:
        """
        Gets a copy of instance with count set.

        Args:
            value (int): Count value.

        Returns:
            DropCap: DropCap instance
        """
        cp = self.copy()
        cp.prop_count = value
        return cp

    def fmt_distance(self, value: int) -> DropCap:
        """
        Gets a copy of instance with distance set.

        Args:
            value (int): Distance value.

        Returns:
            DropCap: DropCap instance
        """
        cp = self.copy()
        cp.prop_distance = value
        return cp

    def fmt_lines(self, value: int) -> DropCap:
        """
        Gets a copy of instance with lines set.

        Args:
            value (int): Lines value.

        Returns:
            DropCap: DropCap instance
        """
        cp = self.copy()
        cp.prop_lines = value
        return cp

    # endregion format methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA

    @property
    def prop_count(self) -> int:
        """Gets/Sets the number of characters in the drop cap. Must be from 0 to 255."""
        return self._get("Count")

    @prop_count.setter
    def prop_count(self, value: int) -> None:
        val = Byte(value)
        self._set("Count", val.value)

    @property
    def prop_lines(self) -> int:
        """Gets/Sets the number of lines used for a drop cap. Must be from 0 to 255."""
        return self._get("Lines")

    @prop_lines.setter
    def prop_lines(self, value: int) -> None:
        val = Byte(value)
        self._set("Lines", val.value)

    @property
    def prop_distance(self) -> int:
        """Gets/Sets the distance between the drop cap in the following text."""
        return self._get("Distance")

    @prop_distance.setter
    def prop_distance(self, value: int) -> None:
        if value < 0:
            raise ValueError("value must be a positive number.")
        self._set("Distance", value)

    # endregion properties
