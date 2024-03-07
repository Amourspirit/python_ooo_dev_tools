from __future__ import annotations
from typing import Sequence, overload, TYPE_CHECKING, TypeVar, Generic
import uno

from ooodev.mock import mock_g
from ooodev.adapter.drawing.graphic_object_shape_comp import GraphicObjectShapeComp
from ooodev.office import write as mWrite
from ooodev.loader import lo as mLo
from ooodev.utils import selection as mSelection
from ooodev.utils.color import Color, CommonColor
from ooodev.utils.context.lo_context import LoContext
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.proto.component_proto import ComponentT

if TYPE_CHECKING:
    from com.sun.star.text import XTextContent
    from com.sun.star.text import XTextCursor
    from ooo.dyn.text.control_character import ControlCharacterEnum
    from ooodev.write.table.write_table import WriteTable
    from ooodev.proto.style_obj import StyleT
    from ooodev.units.unit_obj import UnitT
    from ooodev.utils.type_var import PathOrStr, Table

_T = TypeVar("_T", bound="ComponentT")


class TextCursorPartial(Generic[_T]):
    """
    Represents a writer text cursor.

    This class implements ``__len__()`` method, which returns the number of characters in the range.
    """

    def __init__(self, owner: _T, component: XTextCursor, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (_T): Object that owns this component.
            component (XTextCursor): A UNO object that supports ``com.sun.star.text.TextCursor`` service.
        """
        if lo_inst is None:
            self.__lo_inst = mLo.Lo.current_lo
        else:
            self.__lo_inst = lo_inst
        self.__owner = owner
        self.__component = component

    def add_bookmark(self, name: str) -> bool:
        """
        Adds a bookmark with the specified name to the cursor.

        Args:
            name (str): Bookmark name

        Returns:
            bool: True if bookmark is added; Otherwise, False

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.BOOKMARK_ADDING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.BOOKMARK_ADDED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing ``name`` and ``cursor``.
        """
        with LoContext(self.__lo_inst):
            result = mWrite.Write.add_bookmark(self.__component, name)
        return result

    # region add_formula()
    @overload
    def add_formula(self, formula: str) -> mWriteTextContent.WriteTextContent[_T]:
        """
        Adds a formula

        Args:
            formula (str): formula

        Returns:
            WriteTextContent: Embedded Object.
        """
        ...

    @overload
    def add_formula(self, formula: str, styles: Sequence[StyleT]) -> mWriteTextContent.WriteTextContent[_T]:
        """
        Adds a formula

        Args:
            formula (str): formula
            styles (Sequence[StyleT]): One or more styles to apply to frame. Only styles that support ``com.sun.star.text.TextEmbeddedObject`` service are applied.

        Returns:
            WriteTextContent: Embedded Object.
        """
        ...

    def add_formula(
        self, formula: str, styles: Sequence[StyleT] | None = None
    ) -> mWriteTextContent.WriteTextContent[_T]:
        """
        Adds a formula

        Args:
            formula (str): formula
            styles (Sequence[StyleT]): One or more styles to apply to frame. Only styles that support ``com.sun.star.text.TextEmbeddedObject`` service are applied.

        Raises:
            CreateInstanceMsfError: If unable to create text.TextEmbeddedObject
            CancelEventError: If event ``WriteNamedEvent.FORMULA_ADDING`` is cancelled
            Exception: If unable to add formula

        Returns:
            WriteTextContent: Embedded Object.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.FORMULA_ADDING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.FORMULA_ADDED` :eventref:`src-docs-event`

        Hint:
            Styles that can be applied are found in the following packages.

            - :doc:`ooodev.format.writer.direct.obj </src/format/ooodev.format.writer.direct.obj>`

        Note:
            Event args ``event_data`` is a dictionary containing ``formula`` and ``cursor``.
        """
        with LoContext(self.__lo_inst):
            if styles:
                result = mWrite.Write.add_formula(self.__component, formula, styles)
            else:
                result = mWrite.Write.add_formula(self.__component, formula)

        return mWriteTextContent.WriteTextContent(self.__owner, result)

    # endregion add_formula()

    def add_hyperlink(self, label: str, url_str: str) -> bool:
        """
        Add a hyperlink

        Args:
            label (str): Hyperlink label
            url_str (str): Hyperlink url

        Raises:
            CreateInstanceMsfError: If unable to create TextField.URL instance
            Exception: If unable to create hyperlink

        Returns:
            bool: True if hyperlink is added; Otherwise, False

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.HYPER_LINK_ADDING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.HYPER_LINK_ADDED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing ``label``, ``url_str`` and ``cursor``.
        """
        with LoContext(self.__lo_inst):
            result = mWrite.Write.add_hyperlink(self.__component, label, url_str)
        return result

    # region add_image_link
    @overload
    def add_image_link(self, fnm: PathOrStr) -> mWriteTextContent.WriteTextContent[_T]:
        """
        Add Image Link.

        Args:
            fnm (PathOrStr): Image path.

        Returns:
            WriteTextContent: Image Link on success; Otherwise, ``None``.
        """
        ...

    # style: Sequence[StyleT] = None
    @overload
    def add_image_link(
        self, fnm: PathOrStr, *, width: int | UnitT, height: int | UnitT
    ) -> mWriteTextContent.WriteTextContent[_T]:
        """
        Add Image Link.

        Args:
            fnm (PathOrStr): Image path.
            width (int, UnitT): Width in ``1/100th mm`` or ``UnitT``.
            height (int, UnitT): Height in ``1/100th mm`` or ``UnitT``.

        Returns:
            WriteTextContent: Image Link on success; Otherwise, ``None``.
        """
        ...

    @overload
    def add_image_link(
        self,
        fnm: PathOrStr,
        *,
        styles: Sequence[StyleT],
    ) -> mWriteTextContent.WriteTextContent[_T]:
        """
        Add Image Link.

        Args:
            fnm (PathOrStr): Image path.
            styles (Sequence[StyleT]): One or more styles to apply to frame. Only styles that support ``com.sun.star.text.TextGraphicObject`` service are applied.

        Returns:
            WriteTextContent: Image Link on success; Otherwise, ``None``.
        """
        ...

    @overload
    def add_image_link(
        self,
        fnm: PathOrStr,
        *,
        width: int | UnitT,
        height: int | UnitT,
        styles: Sequence[StyleT],
    ) -> mWriteTextContent.WriteTextContent[_T]:
        """
        Add Image Link.

        Args:
            fnm (PathOrStr): Image path.
            width (int, UnitT): Width in ``1/100th mm`` or ``UnitT``.
            height (int, UnitT): Height in ``1/100th mm`` or ``UnitT``.
            styles (Sequence[StyleT]): One or more styles to apply to frame. Only styles that support ``com.sun.star.text.TextGraphicObject`` service are applied.

        Returns:
            WriteTextContent: Image Link on success; Otherwise, ``None``.
        """
        ...

    def add_image_link(
        self,
        fnm: PathOrStr,
        *,
        width: int | UnitT = 0,
        height: int | UnitT = 0,
        styles: Sequence[StyleT] | None = None,
    ) -> mWriteTextContent.WriteTextContent[_T]:
        """
        Add Image Link

        Args:
            fnm (PathOrStr): Image path
            width (int, UnitT): Width in ``1/100th mm`` or :ref:`proto_unit_obj`.
            height (int, UnitT): Height in ``1/100th mm`` or :ref:`proto_unit_obj`.
            styles (Sequence[StyleT]): One or more styles to apply to frame. Only styles that support ``com.sun.star.text.TextGraphicObject`` service are applied.

        Raises:
            CreateInstanceMsfError: If Unable to create text.TextGraphicObject.
            MissingInterfaceError: If unable to obtain XPropertySet interface.
            Exception: If unable to add image.
            CancelEventError: If ``IMAGE_LINK_ADDING`` event is canceled.

        Returns:
            XTextContent: Image Link on success; Otherwise, ``None``.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.IMAGE_LINK_ADDING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.IMAGE_LINK_ADDED` :eventref:`src-docs-event`

        Hint:
            Styles that can be applied are found in the following packages.

            - :doc:`ooodev.format.writer.direct.image </src/format/ooodev.format.writer.direct.image>`

        Note:
            Event args ``event_data`` is a dictionary containing ``doc``, ``cursor``, ``fnm``, ``width`` and ``height``.
        """
        if styles is None:
            styles = ()
        with LoContext(self.__lo_inst):
            result = mWrite.Write.add_image_link(
                doc=self.__owner.component,
                cursor=self.__component,
                fnm=fnm,
                width=width,
                height=height,
                styles=styles,
            )
        return mWriteTextContent.WriteTextContent(self.__owner, result)

    # endregion add_image_link

    # region add_image_shape()
    @overload
    def add_image_shape(self, fnm: PathOrStr) -> GraphicObjectShapeComp:
        """
        Add Image Shape.

        Args:
            fnm (PathOrStr): Image path.

        Returns:
            GraphicObjectShapeComp: Image Shape on success; Otherwise, ``None``.
        """
        ...

    @overload
    def add_image_shape(self, fnm: PathOrStr, width: int | UnitT, height: int | UnitT) -> GraphicObjectShapeComp:
        """
        Add Image Shape.

        Args:
            fnm (PathOrStr): Image path.
            width (int, UnitT): Width in ``1/100th mm`` or ``UnitT``.
            height (int, UnitT): Height in ``1/100th mm`` or ``UnitT``.

        Returns:
            GraphicObjectShapeComp: Image Shape on success; Otherwise, ``None``.
        """
        ...

    def add_image_shape(
        self, fnm: PathOrStr, width: int | UnitT = 0, height: int | UnitT = 0
    ) -> GraphicObjectShapeComp:
        """
        Add Image Shape.

        Args:
            fnm (PathOrStr): Image path.
            width (int, UnitT): Width in ``1/100th mm`` or :ref:`proto_unit_obj`.
            height (int, UnitT): Height in ``1/100th mm`` or :ref:`proto_unit_obj`.

        Raises:
            CreateInstanceMsfError: If unable to create drawing.GraphicObjectShape.
            ValueError: If unable to get image.
            MissingInterfaceError: If require interface cannot be obtained.
            Exception: If unable to add image shape.
            CancelEventError: if ``IMAGE_SHAPE_ADDING`` event is canceled.

        Returns:
            GraphicObjectShapeComp: Image Shape on success; Otherwise, ``None``.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.IMAGE_SHAPE_ADDING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.IMAGE_SHAPE_ADDED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing ``doc``, ``cursor``, ``fnm``, ``width`` and ``height``.
        """
        with LoContext(self.__lo_inst):
            result = mWrite.Write.add_image_shape(cursor=self.__component, fnm=fnm, width=width, height=height)
        return GraphicObjectShapeComp(result)

    # endregion add_image_shape()

    def add_line_divider(self, line_width: int) -> None:
        """
        Adds a line divider

        Args:
            line_width (int): Line width

        Raises:
            CreateInstanceMsfError: If unable to create drawing.LineShape instance
            MissingInterfaceError: If unable to obtain XShape interface
            Exception: If unable to add Line divider
        """
        with LoContext(self.__lo_inst):
            mWrite.Write.add_line_divider(self.__component, line_width)

    def add_table(
        self,
        table_data: Table,
        *,
        name="",
        header_bg_color: Color | None = None,
        header_fg_color: Color | None = None,
        tbl_bg_color: Color | None = None,
        tbl_fg_color: Color | None = None,
        first_row_header: bool = True,
        styles: Sequence[StyleT] | None = None,
    ) -> WriteTable[_T]:
        """
        Adds a table.

        Each row becomes a row of the table. The first row is treated as a header.

        Args:
            table_data (Table): 2D Table with the the first row containing column names.
            name (str, optional): Table name.
            header_bg_color (:py:data:`~.utils.color.Color`, optional): Table header background color.
                Set to None to ignore header color. Defaults to ``None``.
            header_fg_color (:py:data:`~.utils.color.Color`, optional): Table header foreground color.
                Set to None to ignore header color. Defaults to `Defaults to ``None``.
            tbl_bg_color (:py:data:`~.utils.color.Color`, optional): Table background color.
                Set to None to ignore background color. Defaults to ``None``.
            tbl_fg_color (:py:data:`~.utils.color.Color`, optional): Table background color.
                Set to None to ignore background color. Defaults to ``None``.
            first_row_header (bool, optional): If ``True`` First row is treated as header data. Default ``True``.
            styles (Sequence[StyleT], optional): One or more styles to apply to frame.
                Only styles that support ``com.sun.star.text.TextTable`` service are applied.

        Raises:
            ValueError: If table_data is empty
            CreateInstanceMsfError: If unable to create instance of text.TextTable
            CancelEventError:  If ``WriteNamedEvent.TABLE_ADDING`` event cancelled
            Exception: If unable to add table

        Returns:
            WriteTable: Table that is added to document.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.TABLE_ADDING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.TABLE_ADDED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing all method args.

        See Also:
            - :ref:`help_writer_format_direct_table`
            - :py:class:`~.utils.color.CommonColor`
            - :py:meth:`~.utils.table_helper.TableHelper.table_2d_to_dict`
            - :py:meth:`~.utils.table_helper.TableHelper.table_dict_to_table`

        Hint:
            Styles that can be applied are found in :doc:`ooodev.format.writer.direct.table </src/format/ooodev.format.writer.direct.table>` subpackages.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.write.table.write_table import WriteTable

        with LoContext(self.__lo_inst):
            result = mWrite.Write.add_table(
                cursor=self.__component,
                table_data=table_data,
                header_bg_color=header_bg_color,
                header_fg_color=header_fg_color,
                tbl_bg_color=tbl_bg_color,
                tbl_fg_color=tbl_fg_color,
                first_row_header=first_row_header,
                styles=styles,
            )
        tbl = WriteTable(self.__owner, result)
        if name:
            tbl.name = name
        return tbl

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
    ) -> mWriteTextFrame.WriteTextFrame[_T]:
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
        with LoContext(self.__lo_inst):
            result = mWrite.Write.add_text_frame(
                cursor=self.__component,
                text=text,
                ypos=ypos,
                width=width,
                height=height,
                page_num=page_num,
                border_color=border_color,
                background_color=background_color,
                styles=styles,
            )
        return mWriteTextFrame.WriteTextFrame(self.__owner, result)

    # region append()
    @overload
    def append(self, text: str) -> None:
        """
        Append content to cursor

        Args:
            text (str): Text to append.

        Returns:
            None:
        """
        ...

    @overload
    def append(self, text: str, styles: Sequence[StyleT]) -> None:
        """
        Append content to cursor

        Args:
            text (str): Text to append.
            styles (Sequence[StyleT]):One or more styles to apply to text.

        Returns:
            None:
        """
        ...

    @overload
    def append(self, ctl_char: ControlCharacterEnum) -> None:
        """
        Append content to cursor

        Args:
            ctl_char (int): Control Char (like a paragraph break or a hard space).

        Returns:
            None:
        """
        ...

    @overload
    def append(self, text_content: XTextContent) -> None:
        """
        Append content to cursor

        Args:
            text_content (XTextContent): Text content, such as a text table, text frame or text field.

        Returns:
            None:
        """
        ...

    def append(self, *args, **kwargs) -> None:
        """
        Append content to cursor

        Args:
            text (str): Text to append.
            styles (Sequence[StyleT]):One or more styles to apply to text.
            ctl_char (int): Control Char (like a paragraph break or a hard space).
            text_content (XTextContent): Text content, such as a text table, text frame or text field.

        Returns:
            None:

        :events:
            If using styles then the following events are triggered for each style.

            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLED` :eventref:`src-docs-event`

        Hint:
            Styles that can be applied are found in the following packages.

            - :doc:`ooodev.format.writer.direct.char </src/format/ooodev.format.writer.direct.char>`

        See Also:
            `API ControlCharacter <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1text_1_1ControlCharacter.html>`_
        """
        mWrite.Write.append(self.__component, *args, **kwargs)

    # endregion append()

    def append_date_time(self) -> None:
        """
        Append two DateTime fields, one for the date, one for the time

        Raises:
            MissingInterfaceError: If required interface cannot be obtained.
        """
        with LoContext(self.__lo_inst):
            mWrite.Write.append_date_time(self.__component)

    # region append_line()
    @overload
    def append_line(
        self,
    ) -> None: ...

    @overload
    def append_line(self, text: str) -> None:
        """
        Appends a new Line.

        Args:
            text (str, optional): text to append before new line is inserted.

        Returns:
            None:
        """
        ...

    @overload
    def append_line(self, text: str, styles: Sequence[StyleT]) -> None:
        """
        Appends a new Line.

        Args:
            text (str, optional): text to append before new line is inserted.
            styles (Sequence[StyleT]): One or more styles to apply to text. If ``text`` is omitted then this argument is ignored.

        Returns:
            None:
        """
        ...

    def append_line(self, text: str = "", styles: Sequence[StyleT] | None = None) -> None:
        """
        Appends a new Line.

        Args:
            text (str, optional): text to append before new line is inserted.
            styles (Sequence[StyleT]): One or more styles to apply to text. If ``text`` is omitted then this argument is ignored.

        Returns:
            None:

        :events:
            If using styles then the following events are triggered for each style.

            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLED` :eventref:`src-docs-event`
        """
        if styles:
            mWrite.Write.append_line(self.__component, text, styles)
        else:
            mWrite.Write.append_line(self.__component, text)

    # endregion append_line()

    # region append_para()
    @overload
    def append_para(self) -> None:
        """
        Appends text (if present) and then a paragraph break.

        Returns:
            None:
        """
        ...

    @overload
    def append_para(self, text: str) -> None:
        """
        Appends text (if present) and then a paragraph break.

            text (str, optional): Text to append

        Returns:
            None:
        """
        ...

    @overload
    def append_para(self, text: str, styles: Sequence[StyleT]) -> None:
        """
        Appends text (if present) and then a paragraph break.

        Args:
            text (str, optional): Text to append
            styles (Sequence[StyleT]): One or more styles to apply to text. If ``text`` is empty then this argument is ignored.

        Returns:
            None:
        """
        ...

    def append_para(self, text: str = "", styles: Sequence[StyleT] | None = None) -> None:
        """
        Appends text (if present) and then a paragraph break.

        Args:
            text (str, optional): Text to append
            styles (Sequence[StyleT]): One or more styles to apply to text. If ``text`` is empty then this argument is ignored.

        Returns:
            None:

        :events:
            If using styles then the following events are triggered for each style.

            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLED` :eventref:`src-docs-event`

        Hint:
            Styles that can be applied are found in the following packages.

            - :doc:`ooodev.format.writer.direct.char </src/format/ooodev.format.writer.direct.char>`
            - :doc:`ooodev.format.writer.direct.para </src/format/ooodev.format.writer.direct.para>`
        """
        # paragraph break (UNICODE 0x000D). \r
        # https://wiki.documentfoundation.org/Documentation/DevGuide/Text_Documents#Control_Characters
        if styles:
            mWrite.Write.append_para(self.__component, text, styles)
        else:
            mWrite.Write.append_para(self.__component, text)

    # endregion append_para()

    def column_break(self) -> None:
        """
        Inserts a column break

        """
        mWrite.Write.column_break(self.__component)

    def end_line(self) -> None:
        """
        Inserts a line break
        """
        mWrite.Write.end_line(self.__component)

    def end_paragraph(self) -> None:
        """
        Inserts a paragraph break
        """
        mWrite.Write.end_paragraph(self.__component)

    def get_all_text(self) -> str:
        """
        Gets the text part of the document

        Returns:
            str: text
        """
        # Note UNO cursor.getString() replaces the \r with \n automatically even
        # though \r is the paragraph break character.

        self.__component.gotoStart(False)
        self.__component.gotoEnd(True)
        text = self.__component.getString()
        self.__component.gotoEnd(False)  # to deselect everything
        return text

    def get_pos(self) -> int:
        """
        Gets position of the cursor

        Args:
            cursor (XTextCursor): Text Cursor

        Returns:
            int: Current Cursor Position

        Note:
            This method is not the most reliable.
            It attempts to read all the text in a document and move the cursor to the end
            and then get the position.

            It would be better to use cursors from relative positions in bigger documents.
        """
        with LoContext(self.__lo_inst):
            result = mSelection.Selection.get_position(self.__component)
        return result

    def insert_para(self, para: str, para_style: str) -> None:
        """
        Inserts a paragraph with a style applied

        Args:
            para (str): Paragraph text
            para_style (str): Style such as 'Heading 1'
        """
        mWrite.Write.insert_para(self.__component, para, para_style)

    def page_break(self) -> None:
        """
        Inserts a page break
        """
        mWrite.Write.page_break(self.__component)

    # region style()
    @overload
    def style(self, *, pos: int, length: int, styles: Sequence[StyleT]) -> None:
        """
        Styles. From position styles right by distance amount.

        Args:
            pos (int): Position style start.
            length (int): The distance from ``pos`` to apply style.
            styles (Sequence[StyleT]):One or more styles to apply to text.

        Returns:
            None:
        """
        ...

    @overload
    def style(self, *, pos: int, length: int, prop_name: str, prop_val: object) -> None:
        """
        Styles. From position styles right by distance amount.

        Args:
            pos (int): Position style start.
            length (int): The distance from ``pos`` to apply style.
            prop_name (str): Property Name such as ``CharHeight``
            prop_val (object): Property Value such as ``10``

        Returns:
            None:
        """
        ...

    def style(self, **kwargs) -> None:
        """
        Styles. From position styles right by distance amount.

        Args:
            pos (int): Position style start.
            length (int): The distance from ``pos`` to apply style.
            styles (Sequence[StyleT]):One or more styles to apply to text.
            prop_name (str): Property Name such as ``CharHeight``
            prop_val (object): Property Value such as ``10``

        Returns:
            None:

        See Also:
            :py:meth:`~.write_doc.WriteDoc.style_left`

        Note:
            Unlike :py:meth:`~.Write.style_left` this method does not restore any style properties after style is applied.
        """
        kwargs["cursor"] = self.__component
        with LoContext(self.__lo_inst):
            mWrite.Write.style(**kwargs)

    # endregion style()

    # region style_left()
    @overload
    def style_left(self, pos: int, styles: Sequence[StyleT]) -> None:
        """
        Styles left. From current cursor position to left by pos amount.

        Args:
            pos (int): Positions to style left
            styles (Sequence[StyleT]): One or more styles to apply to text.

        Returns:
            None:
        """
        ...

    @overload
    def style_left(self, pos: int, prop_name: str, prop_val: object) -> None:
        """
        Styles left. From current cursor position to left by pos amount.

        Args:
            pos (int): Positions to style left
            prop_name (str): Property Name such as ``CharHeight``
            prop_val (object): Property Value such as ``10``

        Returns:
            None:
        """
        ...

    def style_left(self, *args, **kwargs) -> None:
        """
        Styles left. From current cursor position to left by pos amount.

        Args:
            pos (int): Positions to style left
            styles (Sequence[StyleT]): One or more styles to apply to text.
            prop_name (str): Property Name such as ``CharHeight``
            prop_val (object): Property Value such as ``10``

        Returns:
            None:

        :events:
            If using styles then the following events are triggered for each style.

            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLED` :eventref:`src-docs-event`

            Otherwise, the following events are triggered once.

            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLING` :eventref:`src-docs-key-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLED` :eventref:`src-docs-key-event`

        See Also:
            :py:meth:`~.write_doc.WriteDoc.style`

        Note:
            This method restores the style properties to their original state after the style is applied.
            This is done so applied style properties are reset before next text is appended.
            This is not the case for :py:meth:`~.Write.style` method.
        """
        with LoContext(self.__lo_inst):
            mWrite.Write.style_left(self.__component, *args, **kwargs)

    # endregion style_left()
    def style_left_bold(self, pos: int) -> None:
        """
        Styles bold from current cursor position left by pos amount.

        Args:
            pos (int): Number of positions to go left
        """
        with LoContext(self.__lo_inst):
            mWrite.Write.style_left_bold(self.__component, pos)

    def style_left_code(self, pos: int) -> None:
        """
        Styles using a Mono font from current cursor position left by pos amount.
        Font Char Height is set to ``10``

        Args:
            pos (int): Number of positions to go left

        Returns:
            None:

        Note:
            The font applied is determined by :py:meth:`.Info.get_font_mono_name`
        """
        with LoContext(self.__lo_inst):
            mWrite.Write.style_left_code(self.__component, pos)

    def style_left_color(self, pos: int, color: Color) -> None:
        """
        Styles color from current cursor position left by pos amount.

        Args:
            pos (int): Number of positions to go left
            color (~ooodev.utils.color.Color): RGB color as int to apply

        Returns:
            None:

        See Also:
            :py:class:`~.utils.color.CommonColor`
        """
        with LoContext(self.__lo_inst):
            mWrite.Write.style_left_color(self.__component, pos, color)

    def style_left_italic(self, pos: int) -> None:
        """
        Styles italic from current cursor position left by pos amount.

        Args:
            pos (int): Number of positions to go left
        """
        with LoContext(self.__lo_inst):
            mWrite.Write.style_left_italic(self.__component, pos)


# avoid circular imports

from ooodev.write import write_text_content as mWriteTextContent
from ooodev.write import write_text_frame as mWriteTextFrame

if mock_g.FULL_IMPORT:
    from ooodev.write.table.write_table import WriteTable
