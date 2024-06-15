from __future__ import annotations
from typing import Sequence, TYPE_CHECKING
import uno

from ooodev.exceptions import ex as mEx
from ooodev.utils import gen_util as mGenUtil
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.adapter.text.text_frames import TextFramesComp
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write.write_text_frame import WriteTextFrame

if TYPE_CHECKING:
    from com.sun.star.container import XNameAccess
    from ooodev.units.unit_obj import UnitT
    from ooodev.proto.style_obj import StyleT
    from ooodev.utils.color import Color
    from ooodev.write.write_doc import WriteDoc  # circular import if not TYPE_CHECKING


class WriteTextFrames(
    LoInstPropsPartial, TextFramesComp, WriteDocPropPartial, QiPartial, ServicePartial, TheDictionaryPartial
):
    """
    Class for managing Writer Text Frames.

    This class is Enumerable and returns ``WriteTextFrame[WriteDoc]`` instance on iteration.
    """

    def __init__(self, owner: WriteDoc, frames: XNameAccess, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (WriteDoc): Owner Component
            frames (XNameAccess): Text Frames instance.
            lo_inst (LoInst, optional): Lo instance. Used when creating multiple documents. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)
        TextFramesComp.__init__(self, frames)  # type: ignore
        QiPartial.__init__(self, component=frames, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=frames, lo_inst=self.lo_inst)
        TheDictionaryPartial.__init__(self)

    def __next__(self) -> WriteTextFrame[WriteDoc]:
        """
        Gets the next Text Frame.

        Returns:
            WriteTextFrame[WriteDoc]: The next Text Frame.
        """
        return WriteTextFrame(owner=self.owner, component=super().__next__(), lo_inst=self.lo_inst)

    def __getitem__(self, index: str | int) -> WriteTextFrame[WriteDoc]:
        """
        Gets the Text Frame at the specified index or name.

        This is short hand for ``get_by_index()`` or ``get_by_name()``.

        Args:
            key (key, str, int): The index or name of the form. When getting by index can be a negative value to get from the end.

        Returns:
            WriteTextFrame[WriteDoc]: The from with the specified index or name.

        See Also:
            - :py:meth:`~ooodev.write.WriteTextFrames.get_by_index`
            - :py:meth:`~ooodev.write.WriteTextFrames.get_by_name`
        """
        if isinstance(index, int):
            return self.get_by_index(index)
        return self.get_by_name(index)

    def __len__(self) -> int:
        """
        Gets the number of Text Frames in the document.

        Returns:
            int: Number of Text Frames in the document.
        """
        return self.component.getCount()

    def _get_index(self, idx: int, allow_greater: bool = False) -> int:
        """
        Gets the index.

        Args:
            idx (int): Index of sheet. Can be a negative value to index from the end of the list.
            allow_greater (bool, optional): If True and index is greater then the number of
                sheets then the index becomes the next index if sheet were appended. Defaults to False.

        Returns:
            int: Index value.
        """
        count = len(self)
        return mGenUtil.Util.get_index(idx, count, allow_greater)

    def _create_name(self, name: str) -> str:
        used_name = True
        i = 1
        nm = f"{name}{i}"
        while used_name:
            used_name = self.has_by_name(nm)
            if used_name:
                i += 1
                nm = f"{name}{i}"
        return nm

    def add_text_frame(
        self,
        *,
        text: str = "",
        ypos: int | UnitT = 300,
        width: int | UnitT = 5000,
        height: int | UnitT = 5000,
        page_num: int = 1,
        border_color: Color | None = None,
        background_color: Color | None = None,
        styles: Sequence[StyleT] | None = None,
    ) -> WriteTextFrame[WriteDoc]:
        """
        Adds a text frame.

        Args:
            text (str, optional): Frame Text
            ypos (int, UnitT. optional): Frame Y pos in ``1/100th mm`` or :ref:`proto_unit_obj`. Default ``300``.
            width (int, UnitT, optional): Width in ``1/100th mm`` or :ref:`proto_unit_obj`.
            height (int, UnitT, optional): Height in ``1/100th mm`` or :ref:`proto_unit_obj`.
            page_num (int, optional): Page Number to add text frame. If ``0`` Then Frame is anchored to paragraph. Default ``1``.
            border_color (:py:data:`~.utils.color.Color`, optional):.color.Color`, optional): Border Color.
            background_color (:py:data:`~.utils.color.Color`, optional): Background Color.
            styles (Sequence[StyleT]): One or more styles to apply to frame. Only styles that support ``com.sun.star.text.TextFrame`` service are applied.

        Raises:
            CreateInstanceMsfError: If unable to create text.TextFrame
            CancelEventError: If ``WriteNamedEvent.TEXT_FRAME_ADDING`` event is cancelled
            Exception: If unable to add text frame

        Returns:
            WriteTextFrame: Text frame that is added to document.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.TEXT_FRAME_ADDING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.TEXT_FRAME_ADDED` :eventref:`src-docs-event`

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.writer.direct.frame </src/format/ooodev.format.writer.direct.frame>` subpackages.

        Note:
            Event args ``event_data`` is a dictionary containing all method args.

        See Also:
            - :py:class:`~.utils.color.CommonColor`
            - :py:class:`~.utils.color.StandardColor`
        """
        cursor = self.owner.get_cursor()
        return cursor.add_text_frame(
            text=text,
            ypos=ypos,
            width=width,
            height=height,
            page_num=page_num,
            border_color=border_color,
            background_color=background_color,
            styles=styles,
        )

    # region XIndexAccess overrides

    def get_by_index(self, idx: int) -> WriteTextFrame[WriteDoc]:
        """
        Gets the element at the specified index.

        Args:
            idx (int): The Zero-based index of the element. Idx can be a negative value to index from the end of the list.
                For example, -1 will return the last element.

        Returns:
            WriteTextFrame: The element at the specified index.
        """
        idx = self._get_index(idx, True)
        result = super().get_by_index(idx)
        return WriteTextFrame(owner=self.write_doc, component=result, lo_inst=self.lo_inst)

    # endregion XIndexAccess overrides

    # region XNameAccess overrides

    def get_by_name(self, name: str) -> WriteTextFrame[WriteDoc]:
        """
        Gets the element with the specified name.

        Args:
            name (str): The name of the element.

        Raises:
            MissingNameError: If text frame is not found.

        Returns:
            WriteTextFrame: The element with the specified name.
        """
        if not self.has_by_name(name):
            raise mEx.MissingNameError(f"Unable to find text frame with name '{name}'")
        result = super().get_by_name(name)
        return WriteTextFrame(owner=self.write_doc, component=result, lo_inst=self.lo_inst)

    # endregion XNameAccess overrides

    # region Properties
    @property
    def owner(self) -> WriteDoc:
        """
        Returns:
            WriteDoc: Writer Draw Page
        """
        return self.write_doc

    # endregion Properties
