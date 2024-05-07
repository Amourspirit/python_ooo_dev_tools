from __future__ import annotations
from typing import cast, TYPE_CHECKING
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

if TYPE_CHECKING:
    from com.sun.star.io import XDataOutputStream

    # from com.sun.star.io import XDataInputStream


def test_pipes(loader) -> None:
    from ooodev.adapter.io.data_input_stream_comp import DataInputStreamComp
    from ooodev.adapter.io.data_output_stream_comp import DataOutputStreamComp
    from ooodev.adapter.io.pipe_comp import PipeComp

    input = DataInputStreamComp.from_lo()
    output = DataOutputStreamComp.from_lo()
    pipe = PipeComp.from_lo()
    input.set_input_stream(pipe.component)
    output.set_output_stream(pipe.component)

    # no need to cast pipe.Predecessor or pipe.Successor they are the equal to output and input

    out_stream = DataOutputStreamComp(cast("XDataOutputStream", pipe.Predecessor))
    assert out_stream == output

    # First, write a series of bytes that represents 3.14159
    output.write_bytes(64, 9, 33, -7, -16, 27, -122, 110)
    d = 2.6
    output.write_double(d)

    # Now, read the pipe.
    # in_stream = DataInputStreamComp(cast("XDataInputStream", pipe.Successor))
    d = input.read_double()

    # Now read the double that was written as a series of bytes.
    s = "2.6 = "
    while input.available() > 0:
        i = input.read_byte()
        # In case the byte was negative, convert it to a positive integer.
        i = i & 255

        s = f"{s}{i:02x}"

    assert s == "2.6 = 4004cccccccccccd"

    input.close_input()
    output.close_output()
    pipe.close_input()
    pipe.close_output()
