from __future__ import annotations
from typing import cast
import uno
from ....writer.style.char.kind.style_char_kind import StyleCharKind as StyleCharKind
from ..char_style_base_multi import CharStyleBaseMulti
from ....direct.char.border.padding import Padding as DirectPadding


class Padding(CharStyleBaseMulti):
    """
    Character Style Border padding.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        left: float | None = None,
        right: float | None = None,
        top: float | None = None,
        bottom: float | None = None,
        padding_all: float | None = None,
        style_name: StyleCharKind | str = StyleCharKind.STANDARD,
        style_family: str = "CharacterStyles",
    ) -> None:
        """
        Constructor

        Args:
            left (float, optional): Character left padding (in mm units).
            right (float, optional): Character right padding (in mm units).
            top (float, optional): Character top padding (in mm units).
            bottom (float, optional): Character bottom padding (in mm units).
            padding_all (float, optional): Character left, right, top, bottom padding (in mm units). If argument is present then ``left``, ``right``, ``top``, and ``bottom`` arguments are ignored.
            style_name (StyleParaKind, str, optional): Specifies the Character Style that instance applies to. Deftult is Default Character Style.
            style_family (str, optional): Style family. Defatult ``CharacterStyles``.

        Returns:
            None:
        """

        direct = DirectPadding(left=left, right=right, top=top, bottom=bottom, padding_all=padding_all)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleCharKind | str = StyleCharKind.STANDARD,
        style_family: str = "CharacterStyles",
    ) -> Padding:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleCharKind, str, optional): Specifies the Character Style that instance applies to. Deftult is Default Character Style.
            style_family (str, optional): Style family. Defatult ``CharacterStyles``.

        Returns:
            Padding: ``Padding`` instance from document properties.
        """
        inst = super(Padding, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = DirectPadding.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleCharKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> DirectPadding:
        """Gets Inner Padding instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DirectPadding, self._get_style_inst("direct"))
        return self._direct_inner
