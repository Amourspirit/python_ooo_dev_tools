import pytest
import sys

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from com.sun.star.document import MacroExecMode


def test_get_active_window(loader, fix_writer_path) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.gui.gui import GUI
    from ooodev.utils.props import Props

    props = Props.make_props(Hidden=True, MacroExecutionMode=MacroExecMode.ALWAYS_EXECUTE_NO_WARN)

    test_doc = fix_writer_path("scandalStart.odt")
    doc = Lo.open_doc(fnm=test_doc, loader=loader, props=props)
    try:
        win = GUI.get_active_window(doc)
        assert win.endswith("scandalStart.odt")
    finally:
        Lo.close_doc(doc, False)


@pytest.mark.skipif(sys.platform == "win32", reason="does not run on windows")
def test_activate(copy_fix_writer, loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.gui.gui import GUI
    from ooodev.utils.props import Props

    props = Props.make_props(Hidden=True, MacroExecutionMode=MacroExecMode.ALWAYS_EXECUTE_NO_WARN)

    # Not sure why but this stopped working on windows.
    # It works in windows if just this test is run, but, not in a group of tests.

    # for a manual test remove loader arg from test_activate and uncomment the next line.
    # loader = Lo.load_office(Lo.ConnectPipe())
    delay = 100  # 1_000
    #
    # does not assert anything.
    # when run manually you can see window being activated, minimized, reactivated, etc.
    test_doc = copy_fix_writer("scandalStart.odt")
    doc = Lo.open_doc(fnm=test_doc, loader=loader, props=props)
    try:
        GUI.activate(doc)
        Lo.delay(delay)
        GUI.minimize(doc)
        Lo.delay(delay)
        win_info = GUI.get_window_identity(doc)
        GUI.activate(win_info.window_name)
        Lo.delay(delay)
        GUI.minimize(doc)
        Lo.delay(delay)
        GUI.activate(GUI.get_active_window(doc))
        Lo.delay(delay)
        GUI.maximize(doc)
        Lo.delay(delay)
        GUI.minimize(doc)
        Lo.delay(delay)
        GUI.minimize(doc)
        GUI.maximize(doc)
        Lo.delay(delay)

    finally:
        Lo.close_doc(doc, False)
        # pass


@pytest.mark.skipif(sys.platform == "win32", reason="does not run on windows")
def test_activate_new_doc(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.gui.gui import GUI

    # Not sure why but this stopped working on windows.
    # It works in windows if just this test is run, but, not in a group of tests.

    # for a manual test remove loader arg from test_activate and uncomment the next line.
    # loader = Lo.load_office(Lo.ConnectPipe(headless=True))
    delay = 100  # 1_000
    #
    # does not assert anything.
    # when run manually you can see window being activated, minimized, reactivated, etc.
    doc = Lo.create_doc(doc_type=Lo.DocTypeStr.WRITER, loader=loader)
    try:
        GUI.activate(doc)
        Lo.delay(delay)
        GUI.minimize(doc)
        Lo.delay(delay)
        win_info = GUI.get_window_identity(doc)
        GUI.activate(win_info.window_name)
        Lo.delay(delay)
        GUI.minimize(doc)
        Lo.delay(delay)
        GUI.activate(GUI.get_active_window(doc))
        Lo.delay(delay)
        GUI.maximize(doc)
        Lo.delay(delay)
        GUI.minimize(doc)
        Lo.delay(delay)
        GUI.minimize(doc)
        GUI.maximize(doc)
        Lo.delay(delay)

    finally:
        Lo.close_doc(doc, False)
        # pass
