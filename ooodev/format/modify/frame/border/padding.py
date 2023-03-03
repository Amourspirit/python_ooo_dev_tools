from __future__ import annotations
from typing import cast
import uno

from .....proto.unit_obj import UnitObj
from ..frame_style_base_multi import FrameStyleBaseMulti
from ....writer.style.frame.style_frame_kind import StyleFrameKind as StyleFrameKind
from ....direct.para.border.padding import Padding as InnerPadding


class Padding(FrameStyleBaseMulti):
    """
    Frame Style Border Padding.

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(
        self,
        *,
        left: float | UnitObj | None = None,
        right: float | UnitObj | None = None,
        top: float | UnitObj | None = None,
        bottom: float | UnitObj | None = None,
        all: float | UnitObj | None = None,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> None:
        """
        Constructor

        Args:
            left (float, UnitObj, optional): Left (in ``mm`` units) or :ref:`proto_unit_obj`.
            right (float, UnitObj, optional): Right (in ``mm`` units)  or :ref:`proto_unit_obj`.
            top (float, UnitObj, optional): Top (in ``mm`` units)  or :ref:`proto_unit_obj`.
            bottom (float, UnitObj,  optional): Bottom (in ``mm`` units)  or :ref:`proto_unit_obj`.
            all (float, UnitObj, optional): Left, right, top, bottom (in ``mm`` units)  or :ref:`proto_unit_obj`. If argument is present then ``left``, ``right``, ``top``, and ``bottom`` arguments are ignored.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Deftult is Default Frame Style.
            style_family (str, optional): Style family. Defatult ``FrameStyles``.

        Returns:
            None:
        """

        direct = InnerPadding(
            left=left, right=right, top=top, bottom=bottom, all=all, _cattribs=self._get_inner_cattribs()
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    # endregion Init

    # region Static Methods
    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> Padding:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Deftult is Default Frame Style.
            style_family (str, optional): Style family. Defatult ``FrameStyles``.

        Returns:
            Padding: ``Padding`` instance from style properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerPadding.from_obj(obj=inst.get_style_props(doc), _cattribs=inst._get_inner_cattribs())
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    # endregion Static Methods

    # region internal methods
    def _get_inner_cattribs(self) -> dict:
        return {"_format_kind_prop": self.prop_format_kind, "_supported_services_values": self._supported_services()}

    # endregion internal methods

    # region Properties
    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleFrameKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerPadding:
        """Gets/Sets Inner Padding instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerPadding, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerPadding) -> None:
        if not isinstance(value, InnerPadding):
            raise TypeError(f'Expected type of InnerTransparency, got "{type(value).__name__}"')
        direct = value.__class__(_cattribs=self._get_inner_cattribs())
        vals = [value._get(pname) for pname in value._props]
        for key, val in zip(direct._props, vals):
            direct._set(key, val)
        self._del_attribs("_direct_inner")
        self._set_style("direct", direct, *direct.get_attrs())

    # endregion Properties
