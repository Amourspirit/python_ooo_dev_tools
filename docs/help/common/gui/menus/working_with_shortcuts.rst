.. _help_working_with_shortcuts:

Working with Shortcuts
======================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 1

Working with :ref:`ooodev.gui.menu.Shortcuts`.

The :ref:`class_calc_calc_doc`, :ref:`class_draw_draw_doc`, :ref:`class_draw_impress_doc` and :ref:`class_write_write_doc` 
have a ``shortcuts`` property that can be used to add shortcuts to the menu for the given application.

:ref:`ooodev.gui.menu.MenuApp` can create and store shortcuts automatically via ``ShortCut`` key entry.

Menu Config
-----------

Sample config that sets up a menu with shortcuts.

.. tabs::

    .. code-tab:: python

        new_menu = {
            "Label": "My Menu",
            "CommandURL": menu_name,
            "Submenu": [
                {
                    "Label": "Execute macro...",
                    "CommandURL": "RunMacro",
                    "ShortCut": "Shift+Ctrl+Alt+E",
                },
                {
                    "Label": "My macro",
                    "CommandURL": {"library": "test", "name": "hello"},
                    "ShortCut": {"key": "Shift+Ctrl+Alt+F", "save": False},
                },
            ],
        }

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

ShortCut
--------

The ``ShortCut`` can be a string or a dictionary.

String
^^^^^^

As a string the value is automatically saved into the ``registrymodifications.xcu`` as a global or local entry.

**Example:**

.. tabs::

    .. code-tab:: python

        "ShortCut": "Shift+Ctrl+Alt+E"

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Dictionary
^^^^^^^^^^^

As a dictionary the accepted keys are ``key`` and ``save`` (lower case).

If ``key`` is not present then the shortcut is ignored.
if ``save`` is not present then it defaults to ``True``.

**Example:**

.. tabs::

    .. code-tab:: python

        "ShortCut": {"key": "Shift+Ctrl+Alt+F", "save": False}

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

When the shortcut is set to false it will still work in the menu but will not be persisted into ``registrymodifications.xcu``.
This means it would have to be loaded again after restarting LibreOffice.

Global Shortcut
---------------

Create
^^^^^^

Manually adding a shortcut to global.

.. tabs::

    .. code-tab:: python

        new_menu = {
            "Label": "My Menu",
            "CommandURL": menu_name,
            "Submenu": [
                {
                    "Label": "all Alone",
                    "CommandURL": ".custom:alone.here",
                },
            ],
        }

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Create a persistent global shortcut.

.. tabs::

    .. code-tab:: python

        from ooodev.gui.menu import Shortcuts
        # ...

        sc = Shortcuts()
        sc.set("Shift+Ctrl+Alt+A", ".custom:alone.here", True)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

``registrymodifications.xcu`` entry.

.. tabs::

    .. code-tab:: xml

        <item oor:path="/org.openoffice.Office.Accelerators/PrimaryKeys/Global">
            <node oor:name="A_SHIFT_MOD1_MOD2" oor:op="replace">
                <prop oor:name="Command" oor:op="fuse">
                    <value xml:lang="en-US">alone.here</value>
                </prop>
            </node>
        </item>

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Remove
^^^^^^

By Command
""""""""""

Removes shortcut from running instance but does not save so on next load the shortcut is back.

.. tabs::

    .. code-tab:: python

        from ooodev.gui.menu import Shortcuts
        # ...
        sc = Shortcuts()
        sc.remove_by_command(".custom:alone.here", False)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Removes shortcut from running instance and persist changes.

.. tabs::

    .. code-tab:: python

        from ooodev.gui.menu import Shortcuts
        # ...
        sc = Shortcuts()
        sc.remove_by_command(".custom:alone.here", True)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

By Shortcut
"""""""""""

Removes shortcut from running instance but does not save so on next load the shortcut is back.

.. tabs::

    .. code-tab:: python

        from ooodev.gui.menu import Shortcuts
        # ...
        sc = Shortcuts()
        sc.remove_by_shortcut("Shift+Ctrl+Alt+A", False)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Removes shortcut from running instance and persist changes.

.. tabs::

    .. code-tab:: python

        from ooodev.gui.menu import Shortcuts
        # ...
        sc = Shortcuts()
        sc.remove_by_shortcut("Shift+Ctrl+Alt+A", True)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Local Shortcut
--------------

Create
^^^^^^

Create a persistent local shortcut.

.. tabs::

    .. code-tab:: python

        from ooodev.calc import CalcDoc

        # ...
        doc = CalcDoc.from_current_doc()
        doc.shortcuts.set("Shift+Ctrl+Alt+A", ".custom:alone.here", True)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

``registrymodifications.xcu`` entry.

.. tabs::

    .. code-tab:: xml

        <item oor:path="/org.openoffice.Office.Accelerators/PrimaryKeys/Modules/org.openoffice.Office.Accelerators:Module['com.sun.star.sheet.SpreadsheetDocument']">
            <node oor:name="A_SHIFT_MOD1_MOD2" oor:op="replace">
                <prop oor:name="Command" oor:op="fuse">
                    <value xml:lang="en-US">alone.here</value>
                </prop>
            </node>
        </item>

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Remove
^^^^^^

By Command
""""""""""

Removes shortcut from running instance but does not save so on next load the shortcut is back.

.. tabs::

    .. code-tab:: python

        from ooodev.calc import CalcDoc

        # ...
        doc = CalcDoc.from_current_doc()
        doc.shortcuts.remove_by_command(".custom:alone.here", False)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Removes shortcut from running instance and persist changes.

.. tabs::

    .. code-tab:: python

        # ...
        doc = CalcDoc.from_current_doc()
        doc.shortcuts.remove_by_command(".custom:alone.here", True)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

By Shortcut
"""""""""""

Removes shortcut from running instance but does not save so on next load the shortcut is back.

.. tabs::

    .. code-tab:: python

        # ...
        doc = CalcDoc.from_current_doc()
        doc.shortcuts.remove_by_shortcut("Shift+Ctrl+Alt+A", False)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Removes shortcut from running instance and persist changes.

.. tabs::

    .. code-tab:: python

        # ...
        doc = CalcDoc.from_current_doc()
        doc.shortcuts.remove_by_shortcut("Shift+Ctrl+Alt+A", True)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Reset
-----

Resetting is suppose to cause the component to reset to some default value.

Resetting stores the configuration including any current changes.

.. tabs::

    .. code-tab:: python

        # ...
        doc = CalcDoc.from_current_doc()
        print(doc.shortcuts.set("Shift+Ctrl+Alt+A", ".custom:alone.here", False))
        doc.shortcuts.reset()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Calling ``doc.shortcuts.reset()`` causes the changes to be saved and the shortcut persist.

Related Topics
--------------

- :ref:`help_creating_menu_using_menu_app`
- :ref:`help_working_with_menu_app`
- :ref:`help_working_with_menu_bar`
- :ref:`help_getting_info_on_commands`