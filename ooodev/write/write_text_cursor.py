from __future__ import annotations
from typing import Sequence, overload, TYPE_CHECKING, TypeVar, Generic
import uno


if TYPE_CHECKING:
    from com.sun.star.text import XTextRange
    from ooo.dyn.text.control_character import ControlCharacterEnum
    from com.sun.star.text import XTextContent
    from ooodev.proto.style_obj import StyleT
    from ooodev.utils.type_var import PathOrStr, Table
    from ooodev.units import UnitT
    from ooodev.proto.component_proto import ComponentT

    T = TypeVar("T", bound="ComponentT")

from ooodev.utils.color import Color, CommonColor
from ooodev.office import write as mWrite
from ooodev.adapter.text.text_cursor_comp import TextCursorComp
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils import lo as mLo
from ooodev.adapter.drawing.graphic_object_shape_comp import GraphicObjectShapeComp

from . import write_text_content as mWriteTextContent
from . import write_text_table as mWriteTextTable
from . import write_text_frame as mWriteTextFrame


class WriteTextCursor(Generic[T], TextCursorComp, PropPartial, QiPartial):
    """Represents a writer text cursor."""

    def __init__(self, owner: T, component: XTextRange) -> None:
        """
        Constructor

        Args:
            owner (WriteDoc): Doc that owns this component.
            col_obj (Any): Range object.
        """
        self.__owner = owner
        TextCursorComp.__init__(self, component)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        # self.__doc = doc

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
        return mWrite.Write.add_bookmark(self.component, name)

    # region add_formula()
    @overload
    def add_formula(self, formula: str) -> mWriteTextContent.WriteTextContent:
        """
        Adds a formula

        Args:
            formula (str): formula

        Returns:
            WriteTextContent: Embedded Object.
        """
        ...

    @overload
    def add_formula(self, formula: str, styles: Sequence[StyleT]) -> mWriteTextContent.WriteTextContent:
        """
        Adds a formula

        Args:
            formula (str): formula
            styles (Sequence[StyleT]): One or more styles to apply to frame. Only styles that support ``com.sun.star.text.TextEmbeddedObject`` service are applied.

        Returns:
            WriteTextContent: Embedded Object.
        """
        ...

    def add_formula(self, formula: str, styles: Sequence[StyleT] | None = None) -> mWriteTextContent.WriteTextContent:
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
        if styles:
            return mWriteTextContent.WriteTextContent(self, mWrite.Write.add_formula(self.component, formula, styles))
        else:
            return mWriteTextContent.WriteTextContent(self, mWrite.Write.add_formula(self.component, formula))

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
        return mWrite.Write.add_hyperlink(self.component, label, url_str)

    # region add_image_link
    @overload
    def add_image_link(self, fnm: PathOrStr) -> mWriteTextContent.WriteTextContent:
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
    ) -> mWriteTextContent.WriteTextContent:
        """
        Add Image Link.

        Args:
            fnm (PathOrStr): Image path.
            width (int, UnitT): Width in ``1/100th mm`` or ``UnitT``.
            height (int, UnitT): Height in ``1/100th mm`` or ``UnitT``.

        Returns:
            WriteTextContent: Image Link on success; Otherwise, ``None``.
        """

    @overload
    def add_image_link(
        self,
        fnm: PathOrStr,
        *,
        styles: Sequence[StyleT],
    ) -> mWriteTextContent.WriteTextContent:
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
    ) -> mWriteTextContent.WriteTextContent:
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
    ) -> mWriteTextContent.WriteTextContent:
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
        result = mWrite.Write.add_image_link(
            doc=self.owner.component,
            cursor=self.component,
            fnm=fnm,
            width=width,
            height=height,
            styles=styles,
        )
        return mWriteTextContent.WriteTextContent(self, result)

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
        result = mWrite.Write.add_image_shape(cursor=self.component, fnm=fnm, width=width, height=height)
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
        mWrite.Write.add_line_divider(self.component, line_width)

    def add_table(
        self,
        table_data: Table,
        header_bg_color: Color | None = CommonColor.DARK_BLUE,
        header_fg_color: Color | None = CommonColor.WHITE,
        tbl_bg_color: Color | None = CommonColor.LIGHT_BLUE,
        tbl_fg_color: Color | None = CommonColor.BLACK,
        first_row_header: bool = True,
        styles: Sequence[StyleT] | None = None,
    ) -> mWriteTextTable.WriteTextTable:
        """
        Adds a table.

        Each row becomes a row of the table. The first row is treated as a header.

        Args:
            cursor (XTextCursor): Text Cursor.
            table_data (Table): 2D Table with the the first row containing column names.
            header_bg_color (:py:data:`~.utils.color.Color`, optional): Table header background color.
                Set to None to ignore header color. Defaults to ``CommonColor.DARK_BLUE``.
            header_fg_color (:py:data:`~.utils.color.Color`, optional): Table header foreground color.
                Set to None to ignore header color. Defaults to ``CommonColor.WHITE``.
            tbl_bg_color (:py:data:`~.utils.color.Color`, optional): Table background color.
                Set to None to ignore background color. Defaults to ``CommonColor.LIGHT_BLUE``.
            tbl_fg_color (:py:data:`~.utils.color.Color`, optional): Table background color.
                Set to None to ignore background color. Defaults to ``CommonColor.BLACK``.
            first_row_header (bool, optional): If ``True`` First row is treated as header data. Default ``True``.
            styles (Sequence[StyleT], optional): One or more styles to apply to frame.
                Only styles that support ``com.sun.star.text.TextTable`` service are applied.

        Raises:
            ValueError: If table_data is empty
            CreateInstanceMsfError: If unable to create instance of text.TextTable
            CancelEventError:  If ``WriteNamedEvent.TABLE_ADDING`` event cancelled
            Exception: If unable to add table

        Returns:
            WriteTextTable: Table that is added to document.

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
        result = mWrite.Write.add_table(
            cursor=self.component,
            table_data=table_data,
            header_bg_color=header_bg_color,
            header_fg_color=header_fg_color,
            tbl_bg_color=tbl_bg_color,
            tbl_fg_color=tbl_fg_color,
            first_row_header=first_row_header,
            styles=styles,
        )
        return mWriteTextTable.WriteTextTable(self, result)

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
    ) -> mWriteTextFrame.WriteTextFrame:
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
        result = mWrite.Write.add_text_frame(
            cursor=self.component,
            text=text,
            ypos=ypos,
            width=width,
            height=height,
            page_num=page_num,
            border_color=border_color,
            background_color=background_color,
            styles=styles,
        )
        return mWriteTextFrame.WriteTextFrame(self, result)

    # region append()
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
        mWrite.Write.append(self.component, *args, **kwargs)

    # endregion append()

    def append_date_time(self) -> None:
        """
        Append two DateTime fields, one for the date, one for the time

        Raises:
            MissingInterfaceError: If required interface cannot be obtained.
        """
        mWrite.Write.append_date_time(self.component)

    # region append_line()
    @overload
    def append_line(
        self,
    ) -> None:
        ...

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
            mWrite.Write.append_line(self.component, text, styles)
        else:
            mWrite.Write.append_line(self.component, text)

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
        if styles:
            mWrite.Write.append_para(self.component, text, styles)
        else:
            mWrite.Write.append_para(self.component, text)

    # endregion append_para()

    def column_break(self) -> None:
        """
        Inserts a column break

        """
        mWrite.Write.column_break(self.component)

    def end_line(self) -> None:
        """
        Inserts a line break
        """
        mWrite.Write.end_line(self.component)

    def end_paragraph(self) -> None:
        """
        Inserts a paragraph break
        """
        mWrite.Write.end_paragraph(self.component)

    def get_all_text(self) -> str:
        """
        Gets the text part of the document

        Returns:
            str: text
        """
        self.component.gotoStart(False)
        self.component.gotoEnd(True)
        text = self.component.getString()
        self.component.gotoEnd(False)  # to deselect everything
        return text

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self.__owner

    # endregion Properties
