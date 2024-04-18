.. _help_creating_menu_using_menu_app:

Creating a menu using MenuApp
=============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 1

.. tabs::

    .. code-tab:: python

        from ooodev.calc import CalcDoc
        from ooodev.utils.kind.menu_lookup_kind import MenuLookupKind

        doc = CalcDoc.create_doc(loader=loader, visible=True)
        menu = doc.menu[MenuLookupKind.TOOLS] # or .menu[".uno:ToolsMenu"]
        itm = menu.items[".uno:AutoComplete"] # or .items[6]

        menu_name = ".custom:my.custom_menu"
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
                    "ShortCut": {"Key": "Shift+Ctrl+Alt+F", "Save": True},
                },
            ],
        }
        if not menu_name in menu:
            # only add the menu if it does not already exist
            menu.insert(new_menu, after=itm.command, save=True)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

In this case ``doc.menu`` is :py:class:`~ooodev.gui.menu.MenuApp` for Calc menus.

Config
------

Menu Keys
^^^^^^^^^

Menu keys are case sensitive

Label
"""""

``Label`` this is the label for the menu and is required.

To insert a separator set the Label value to ``-``

.. tabs::

    .. code-tab:: python

        {
            "Label": "-",
        }

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

CommandURL
"""""""""""

``CommandURL`` Optional for menus separators, Otherwise required.

``CommandURL`` is the command to be executed when the menu item is clicked, such as ``RunMacro``.
This can be a string or a dictionary.

String
~~~~~~

As a string ``CommandURL``  is as follows.

.. tabs::

    .. code-tab:: python

        "CommandURL": ".custom:just.a.command",

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The menu command would be ``just.a.command``.

If string does not start with ``.custom:`` or ``.uno:`` then ``.uno:`` is automatically prepended when command is being dispatched.

In the following code ``CommandURL`` value is set to ``RunMacro``.

.. tabs::

    .. code-tab:: python

        {
            "Label": "Execute macro...",
            "CommandURL": "RunMacro",
            "ShortCut": "Shift+Ctrl+Alt+E",
        },

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

When the menu is build the ``CommandURL`` value is actually ``.uno:RunMacro``.

If you need the actual ``CommandURL`` to be ``RunMacro`` then when in the actual menu then:

.. tabs::

    .. code-tab:: python

        {
            "Label": "Execute macro...",
            "CommandURL": ".custom:RunMacro",
            "ShortCut": "Shift+Ctrl+Alt+E",
        },

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Dictionary
~~~~~~~~~~

A dictionary value indicates that a macro URL should be constructed to run a macro when the menu item is clicked.

.. tabs::

    .. code-tab:: python

        {
            "Label": "My macro",
            "CommandURL": {"library": "test", "name": "hello"},
        },

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

As a Dictionary ``CommandURL`` can accept to following values.

- ``library``  (str, optional): Macro Library. Defaults to ``Standard``.
- ``name`` (str): Macro Name.
- ``language`` (str, optional): Language ``Basic`` or ``Python``. Defaults to ``Basic``.
- ``location`` (str, optional): Location ``user`` or ``application``. Defaults to "user".
- ``module`` (str, optional): Module portion. Only Applies if ``language` is not ``Basic`` or ``Python``. Defaults to ".".

These are the exact same parameters that are accepted by :py:meth:`ooodev.macro.script.MacroScript.get_url_script()` method.

Style
"""""

This is the style of the menu.

.. tabs::

    .. code-tab:: python

        from ooodev.utils.kind.item_style_kind import ItemStyleKind
        from ooodev.utils.kind.menu_lookup_kind import MenuLookupKind

        menu = doc.menu[MenuLookupKind.TOOLS]

        menu_name = ".custom:my.custom_menu"
        new_menu = {
            "Label": "My Menu",
            "CommandURL": menu_name,
            "Submenu": [
                {
                    "Label": "A",
                    "CommandURL": ".uno:WarningCellStyles",
                    "Style": ItemStyleKind.RADIO_CHECK,
                },
                {
                    "Label": "B",
                    "CommandURL": ".uno:FootnoteCellStyles",
                    "Style": ItemStyleKind.RADIO_CHECK,
                },
                {
                    "Label": "C",
                    "CommandURL": ".uno:NoteCellStyles",
                    "Style": ItemStyleKind.RADIO_CHECK,
                },
            ],
        }

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

See `API ItemStyle Constant Group <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1ui_1_1ItemStyle.html>`__.

There is a  :py:class:`~ooodev.utils.kind.item_style_kind.ItemStyleKind` for accessing the constants in an enum.

Related Topics
--------------

- :ref:`help_working_with_menu_app`
- :ref:`help_working_with_menu_bar`
