import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.loader.lo import Lo
from ooodev.write import WriteDoc


def test_search(loader):
    delay = 0
    doc = WriteDoc.create_doc(loader)
    try:
        cursor = doc.get_cursor()

        cursor.append_para("The following points are important:")
        cursor.append_para("Have a good breakfast, This is really important.")

        search_desc = doc.create_search_descriptor()
        search_desc.set_search_string("important")
        search_desc.search_regular_expression = False
        first = doc.find_first(search_desc.component)
        assert first is not None

        first = search_desc.find_first()
        assert first is not None

        nxt = search_desc.find_next(first)
        assert nxt is not None

        things = search_desc.find_all()
        assert things is not None
        assert len(things) == 2
        for thing in things:
            assert thing is not None
            assert thing.get_string() == "important"

        first = doc.find_first(search_desc)
        assert first is not None
        assert first.get_string() == "important"

        Lo.delay(delay)
    finally:
        doc.close_doc()


def test_replace(loader):
    delay = 0
    doc = WriteDoc.create_doc(loader)
    try:
        cursor = doc.get_cursor()

        cursor.append_para("The following points are important:")
        cursor.append_para("Have a good breakfast, This is really important.")

        desc = doc.create_replace_descriptor()
        desc.search_str = "important"
        desc.search_regular_expression = False
        first = doc.find_first(desc.component)
        assert first is not None

        first = desc.find_first()
        assert first is not None

        nxt = desc.find_next(first)
        assert nxt is not None

        things = desc.find_all()
        assert things is not None
        assert len(things) == 2
        for thing in things:
            assert thing is not None
            assert thing.get_string() == "important"

        first = doc.find_first(desc)
        assert first is not None
        assert first.get_string() == "important"

        count = desc.replace_words(["important", "good"], ["great", "funny"])
        assert count == 3
        desc.search_str = "funny"
        s = desc.search_str
        assert s == "funny"
        first = desc.find_first()
        assert first is not None
        assert first.get_string() == "funny"

        desc.set_search_string("great")
        first = desc.find_first()
        assert first is not None
        assert first.get_string() == "great"

        cursor = first.get_cursor()
        # the cursor will be at the end of the word "great"
        # the cursor can access the entire document however the selection of the text range is currently selected.
        assert cursor is not None
        assert cursor.get_string() == "great"

        Lo.delay(delay)
    finally:
        doc.close_doc()


def test_replace_regex(loader):
    delay = 0
    doc = WriteDoc.create_doc(loader)
    try:
        cursor = doc.get_cursor()

        cursor.append_para("abc")
        cursor.append_para("def")
        cursor.append_para("abc\tdef")

        desc = doc.create_replace_descriptor()
        desc.search_regular_expression = True

        desc.search_str = "^abc\\tdef$"
        first = desc.find_first()
        assert first is not None

        Lo.delay(delay)
    finally:
        doc.close_doc()


def test_replace_regex2(loader):
    delay = 0
    doc = WriteDoc.create_doc(loader)
    try:
        # It seems it not possible to replace a newline character.
        # Although a newline Character can be found via the $ character,
        # it cannot be replaced when replace is being called.
        # This text matches lines that are a single digit that ends with \n.
        # Because \n is not contained in the found match, it cannot be replaced directly.
        # The solution here is to create a cursor that moves right by one character and then replace the text.
        # See: https://ask.libreoffice.org/t/find-replace-line-end-n-in-a-macro/102467

        cursor = doc.get_cursor()

        cursor.append_para("1")
        cursor.append_para("a")
        cursor.append_para("2")
        cursor.append_para("b")
        cursor.append_para("3")
        cursor.append_para("c")

        desc = doc.create_replace_descriptor()
        desc.search_regular_expression = True

        # match only lines that are a single digit.
        desc.search_str = "^([0-9])$"
        found = desc.find_first()
        assert found is not None
        # found_cursor = found.get_cursor()
        # found_cursor.go_right(1, True)
        # s = found_cursor.get_string()
        # assert s == "1\n"
        while found is not None:
            found_cursor = found.get_cursor()
            txt = found.get_string()
            found_cursor.go_right(1, True)
            found_cursor.set_string(txt)
            found = desc.find_next(found)

        cursor = doc.get_cursor()
        s = cursor.get_all_text()
        assert s == "1a\n2b\n3c\n"

        Lo.delay(delay)
    finally:
        doc.close_doc()
