from __future__ import annotations
import uno
from ooodev.adapter.io.pipe_comp import PipeComp
from ooodev.adapter.io.text_input_stream_comp import TextInputStreamComp
from ooodev.adapter.io.text_output_stream_comp import TextOutputStreamComp

from ooodev.utils.string.str_list import StrList


class TextStream:
    """Class for Text Stream."""

    @staticmethod
    def get_text_input_stream_from_bytes(*byte: int) -> TextInputStreamComp:
        """
        Gets a text input stream from bytes.

        Args:
            byte (int): One or more Bytes to convert.

        Returns:
            TextInputStreamComp: Text input stream.
        """
        pipe = PipeComp.from_lo()
        txt_stream = TextInputStreamComp.from_lo()
        txt_stream.set_input_stream(pipe.component)
        try:
            pipe.write_bytes(*byte)
        finally:
            pipe.close_output()

        return txt_stream

    @staticmethod
    def get_text_input_stream_from_str_list(strings: StrList, encoding: str = "UTF-8") -> TextInputStreamComp:
        """
        Gets a text input stream from bytes.

        Args:
            strings (StrList): One or more Bytes to convert.
            encoding (str, optional): Encoding for the strings. Defaults to ``UTF-8``.

        Returns:
            TextInputStreamComp: Text input stream.
        """
        pipe = PipeComp.from_lo()
        txt_in_stream = TextInputStreamComp.from_lo()
        txt_in_stream.set_input_stream(pipe.component)
        txt_in_stream.set_encoding(encoding)
        txt_out_stream = TextOutputStreamComp.from_lo()
        txt_out_stream.set_output_stream(pipe.component)
        try:
            txt_out_stream.set_encoding(encoding)
            txt_out_stream.write_string(str(strings))
        finally:
            pipe.close_output()

        return txt_in_stream

    @staticmethod
    def get_text_input_stream_from_str(text: str, encoding: str = "UTF-8") -> TextInputStreamComp:
        """
        Gets a text input stream from bytes.

        Args:
            strings (StrList): One or more Bytes to convert.
            encoding (str, optional): Encoding for the strings. Defaults to ``UTF-8``.

        Returns:
            TextInputStreamComp: Text input stream.
        """
        pipe = PipeComp.from_lo()
        txt_in_stream = TextInputStreamComp.from_lo()
        txt_in_stream.set_input_stream(pipe.component)
        txt_in_stream.set_encoding(encoding)
        txt_out_stream = TextOutputStreamComp.from_lo()
        txt_out_stream.set_output_stream(pipe.component)
        try:
            txt_out_stream.set_encoding(encoding)
            txt_out_stream.write_string(text)
        finally:
            pipe.close_output()

        return txt_in_stream

    @staticmethod
    def get_str_list_from_text_input_stream(
        txt_stream: TextInputStreamComp, sep: str, close_stream: bool = True
    ) -> StrList:
        """
        Gets a StrList from a TextInputStreamComp.

        Args:
            txt_stream (TextInputStreamComp): TextInputStreamComp to convert.
            sep (str): Separator.
            close_stream (bool, optional): Close the stream after reading. Defaults to ``True``.

        Returns:
            StrList: StrList.
        """
        str_lst = StrList(sep=sep)
        while not txt_stream.is_eof():
            str_lst.append(txt_stream.read_string(True, sep))

        if close_stream:
            txt_stream.close_input()
        return str_lst
