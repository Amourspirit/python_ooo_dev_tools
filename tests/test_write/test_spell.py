import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])


def _test_spell_failing(loader):
    """
    This test Fails on:
         alts = speller.spell(word, loc, ())

        Errors with: CannotConvertException: TYPE is not supported! ./stoc/source/typeconv/convert.cxx:362

    Getting sepller from  speller = lingo_mgr.getSpellChecker() does not work.

    Not certain why but:
        speller = Lo.create_instance_mcf(XSpellChecker, "com.sun.star.linguistic2.SpellChecker")
    Getting speller this way works.
    """
    import uno
    from ooodev.loader.lo import Lo
    from com.sun.star.linguistic2 import XLinguServiceManager2
    from com.sun.star.lang import Locale  # struct class

    lingo_mgr = Lo.create_instance_mcf(XLinguServiceManager2, "com.sun.star.linguistic2.LinguServiceManager")
    assert lingo_mgr is not None

    # load spell checker & thesaurus
    speller = lingo_mgr.getSpellChecker()
    assert speller is not None

    word = "horsebacc"

    loc = Locale("en", "US", "")
    alts = speller.spell(word, loc, ())
    assert alts is not None


def test_spell2(loader):
    """
    This test requires Write to be visible.
    If not visible then Write.is_anything_selected() will return false every time.
    """
    import uno
    from ooodev.loader.lo import Lo
    from com.sun.star.lang import Locale  # struct class
    from com.sun.star.linguistic2 import XSpellChecker

    # https://cgit.freedesktop.org/libreoffice/core/commit/?id=af8143bc40cf2cfbc12e77c9bb7de01b655f7b30

    speller = Lo.create_instance_mcf(XSpellChecker, "com.sun.star.linguistic2.SpellChecker")
    assert speller is not None

    word = "horsebacc"

    loc = Locale("en", "US", "")
    alts = speller.spell(word, loc, ())
    assert alts is not None
