from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])


def test_basic(loader) -> None:
    # get_sheet is overload method.
    # testing each overload.
    from ooodev.calc import CalcDoc
    from ooodev.macro.script.basic import Basic

    doc = CalcDoc.create_doc(loader)
    try:

        def r_trim(input: str):
            # Deletes the String 'SmallString' out of the String 'BigString'
            # in case SmallString's Position in BigString is right at the end.
            # Does not to a real trim operation only removes the exact string for the second input.
            script = Basic.get_basic_script(macro="RTrimStr", module="Strings", library="Tools", embedded=False)
            res = script.invoke((input, " "), (), ())
            return res[0]

        result = r_trim("hello ")
        assert result == "hello"

        # test to see if lrucache is working
        result = r_trim("hello again ")
        assert result == "hello again"
    finally:
        doc.close_doc()
