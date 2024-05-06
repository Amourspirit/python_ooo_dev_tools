from __future__ import annotations
import pytest
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])


def test_simple_file_access(loader, tmp_path_session: Path) -> None:
    from ooodev.adapter.ucb.simple_file_access_comp import SimpleFileAccessComp
    from ooodev.adapter.io.text_output_stream_comp import TextOutputStreamComp
    from ooodev.adapter.io.text_input_stream_comp import TextInputStreamComp

    comp = SimpleFileAccessComp.from_lo()
    strings = ("One", "UTF:Āā", "1@3")

    pth = tmp_path_session / "simple_file_access"
    pth.mkdir(exist_ok=True)
    pth = pth / "test.txt"

    with open(pth, "w") as f:
        f.write("Hello World" + "\n")

    assert pth.exists()

    uri = pth.as_uri()

    if comp.exists(uri):
        comp.kill(uri)

    assert pth.exists() is False

    txt_out_stream = TextOutputStreamComp.from_lo()

    stream = comp.open_file_write(uri)
    # Attach the simple stream to the text stream.
    # The text stream will use the simple stream.
    txt_out_stream.set_output_stream(stream)

    # Write the strings.
    try:
        for s in strings:
            txt_out_stream.write_string(f"{s}\n")
    finally:
        # Close the stream.
        txt_out_stream.close_output()

    assert pth.exists()
    with open(pth, "r") as f:
        assert f.read() == "\n".join(strings) + "\n"

    in_stream = TextInputStreamComp.from_lo()
    stream = comp.open_file_read(uri)
    in_stream.set_input_stream(stream)

    try:
        for s in strings:
            assert in_stream.read_line().rstrip() == s
    finally:
        in_stream.close_input()
