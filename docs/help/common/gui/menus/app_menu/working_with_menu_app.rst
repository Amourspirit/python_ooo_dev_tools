.. _help_working_with_menu_app:

Working with MenuApp
====================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 1

Working with :ref:`ooodev.gui.menu.MenuApp`.

.. tabs::

    .. code-tab:: python

        from ooodev.calc import CalcDoc
        from ooodev.utils.kind.menu_lookup_kind import MenuLookupKind

        doc = CalcDoc.create_doc(loader=loader, visible=True)
        menu = doc.menu[MenuLookupKind.TOOLS] # or .menu[".uno:ToolsMenu"]
        itm = menu.items[".uno:AutoComplete"] # or .items[6]

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The menu can be accessed via the ``doc.menu`` property.
The ``doc.menu`` is an instance of the :ref:`ooodev.gui.menu.MenuApp` and has access to all the menu items of the current menu bar.

The :py:class`~ooodev.utils.kind.menu_lookup_kind.MenuLookupKind` Enum is for convenience and can be replaced with the command name of the menu.

``doc.menu[MenuLookupKind.TOOLS]`` is the same as ``doc.menu[".uno:ToolsMenu"]``.

The menu can be accessed even if the menu is not visible in the LibreOffice Window.

Menu Items
----------

Menu items in this context are either a :py:class:`~ooodev.gui.menu.item.MenuItem`, :py:class:`~ooodev.gui.menu.item.MenuItemSub` or :py:class:`~ooodev.gui.menu.item.MenuItemSep` class instances.

The :py:class:`~ooodev.gui.menu.item.MenuItemKind`` Enum is used to check the type of menu item.

.. tabs::

    .. code-tab:: python

        from ooodev.gui.menu.item import MenuItem
        from ooodev.gui.menu.item import MenuItemSep
        from ooodev.gui.menu.item import MenuItemSub
        from ooodev.gui.menu.item import MenuItemKind

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

MenuItemSub
-----------

:py:class:`~ooodev.gui.menu.item.MenuItemSub` is a child class of :py:class:`~ooodev.gui.menu.item.MenuItem` so:

Sub Menu Item

.. tabs::

    .. code-tab:: python

        menu = doc.menu[MenuLookupKind.TOOLS]
        itm = menu.items['.uno:LanguageMenu']
        assert itm.item_kind == MenuItemKind.ITEM_SUBMENU # equals 3
        assert itm.item_kind >= MenuItemKind.ITEM # equals 2
        assert itm.item_kind != MenuItemKind.SEP # equals 1
        assert itm.item_kind > MenuItemKind.SEP

        assert isinstance(itm, MenuItem)
        assert isinstance(itm, MenuItemSub)
        assert not isinstance(itm, MenuItemSep)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:class:`~ooodev.gui.menu.item.MenuItemSub` menu items also contain a ``sub_menu`` property that get access to is sub menu as another instance of the :py:class:`~ooodev.gui.menu.Menu` class.

MenuItem
--------

:py:class:`~ooodev.gui.menu.Menu` Item.

.. tabs::

    .. code-tab:: python

        itm = menu.items[".uno:AutoComplete"]
        assert itm.item_kind < MenuItemKind.ITEM_SUBMENU
        assert itm.item_kind == MenuItemKind.ITEM
        assert itm.item_kind != MenuItemKind.SEP
        assert itm.item_kind > MenuItemKind.SEP

        assert isinstance(itm, MenuItem)
        assert not isinstance(itm, MenuItemSub)
        assert not isinstance(itm, MenuItemSep)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

MenuItemSep
-----------

:py:class:`~ooodev.gui.menu.item.MenuItemSep` represent a separator in a menu.


.. tabs::

    .. code-tab:: python

        itm = menu.items[4]
        assert itm.item_kind == MenuItemKind.SEP
        assert itm.item_kind < MenuItemKind.ITEM
        assert itm.item_kind < MenuItemKind.ITEM_SUBMENU

        assert isinstance(itm, MenuItemSep)
        assert not isinstance(itm, MenuItem)
        assert not isinstance(itm, MenuItemSub)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Getting Access to Menus
-----------------------

Accessing a menu is simple when working with a doc.

.. tabs::

    .. code-tab:: python

        from ooodev.utils.kind.menu_lookup_kind import MenuLookupKind
        from ooodev.calc import CalcDoc
        from ooodev.loader import Lo
        # ...

        loader = Lo.load_office(connector=Lo.ConnectPipe())
        doc = CalcDoc.create_doc(loader=loader, visible=True)
        # doc.menu contains all the top level menus
        tool_menu = doc.menu[MenuLookupKind.TOOLS]
        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The :py:class:`~ooodev.utils.kind.menu_lookup_kind.MenuLookupKind` is for convenience and in this case returns ``.uno:ToolsMenu``.

Index Access
------------

The ``doc.menu[]`` index access can take a string or a zero-based index number.
``doc.menu[0]`` would give access to the first menu, most likely the ``File`` menu.

There is no recursive search in the :py:class:`~ooodev.gui.menu.MenuApp` or ``MenuItem*`` classes. There is index access via menu position and menu command name.

Usually using the name is more practical as it will find the menu even if the user has reorder it in a different place.

The :py:class:`~ooodev.utils.kind.menu_lookup_kind.MenuLookupKind` Enum is for convenience and can be replaced with the command name of the menu.

``doc.menu[MenuLookupKind.TOOLS]`` is the same as ``doc.menu[".uno:ToolsMenu"]``.

Getting menu items in a menu is basically the same as finding a menu.

Menus has some limits as not all popup menus are actually sub menus.
For instance the menu ``Insert -> Shapes -> Basic Shapes`` corresponds to the following:

.. tabs::

    .. code-tab:: python

        >>> itm = (
        >>> 	doc.menu[".uno:InsertMenu"]
        >>> 	.items[".uno:ShapesMenu"]
        >>> 	.sub_menu.items[".uno:BasicShapes"]
        >>> )
        >>> repl(itm)
        '<MenuItem(command=".uno:BasicShapes", kind=MenuItemKind.ITEM)>'

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Although there is a  popup menu for ``Insert -> Shapes -> Basic Shapes`` it is reported as a ``MenuItem`` and not a ``MenuItemSub``.
This by design because ``.uno:BasicShapes`` has a popup menu but the popup is not really a submenu.
This is also reflected by the Menu Id of the ``.uno:BasicShapes`` popup items.
The first item in the popup has a menu id of ``1`` and the second item has an id of ``2`` and so on.

Alternatively the ``Basic Shapes`` menu can be access in the following way:

.. tabs::

    .. code-tab:: python

        basic_shapes = doc.menu[".uno:InsertMenu"][".uno:ShapesMenu"][".uno:BasicShapes"]

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

However, like the previous method the sub menu is still not available.

.. tabs::

    .. code-tab:: python

        >>> basic_shapes.items[".uno:BasicShapes.circle"]
        KeyError: "Menu item '.uno:BasicShapes.circle' not found"

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

In the ``menubar.xml`` file you can also see that ``.uno:BasicShapes`` has no popup menu in the configuration.

.. tabs::

    .. code-tab:: xml

        <menu menu:id=".uno:ShapesMenu">
            <menupopup>
                <menu menu:id=".uno:ShapesLineMenu">
                    <menupopup>
                        <menuitem menu:id=".uno:Line" />
                        <menuitem menu:id=".uno:Freeline_Unfilled" />
                        <menuitem menu:id=".uno:Freeline" />
                        <menuitem menu:id=".uno:Bezier_Unfilled" />
                        <menuitem menu:id=".uno:BezierFill" />
                        <menuitem menu:id=".uno:Polygon_Unfilled" />
                        <menuitem menu:id=".uno:Polygon_Diagonal_Unfilled" />
                        <menuitem menu:id=".uno:Polygon_Diagonal" />
                    </menupopup>
                </menu>
                <menuitem menu:id=".uno:BasicShapes" />
                <menuitem menu:id=".uno:ArrowShapes" />
                <menuitem menu:id=".uno:SymbolShapes" />
                <menuitem menu:id=".uno:StarShapes" />
                <menuitem menu:id=".uno:CalloutShapes" />
                <menuitem menu:id=".uno:FlowChartShapes" />
            </menupopup>
        </menu>

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The :py:class:`~ooodev.gui.menu.item.MenuItem`, :py:class:`~ooodev.gui.menu.item.MenuItemSub` and :py:class:`~ooodev.gui.menu.item.MenuItemSep` have a ``item_kind`` property that also can be used to check for the appropriate type before taking action.

.. tabs::

    .. code-tab:: python

        from ooodev.gui.menu.item import MenuItemKind
        # ...

        if itm.item_kind >= MenuItemKind.ITEM:
            # `MenuItem, do work
            MenuItem.execute() # run the menu command

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Add Menu
--------

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
                    "Label": "Python Hello World",
                    "CommandURL": {
                        "library": "HelloWorld",
                        "name": "HelloWorldPython",
                        "language": "Python",
                        "location": "share",
                    },
                },
            ],
        }

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Adding a menu is done with the ``insert()`` method.

Only add the menu if it does not exist.
If the menu did exist then this could cause some issues at getting a menu my name or index may return the incorrect instance if the menu was added twice with the same name.
The ``save=True`` option means the changes will be persisted.

If you only wanted the menu to be available for the current instance then ``save=False`` could be used and the menu would not be persisted.

.. tabs::

    .. code-tab:: python

        if not menu_name in menu:
            # only add the menu if it does not already exist
            menu.insert(new_menu, after=itm.command, save=True)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Remove Menu
-----------

The ``remove()`` method is used to remove a submenu from a menu.
The ``save=True`` option means the changes will be persisted.

.. tabs::

    .. code-tab:: python

        menu_name = ".custom:my.custom_menu" # or can just be "my.custom_menu"
        if menu_name in menu:
            menu.remove(menu_name, save=True)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Execute Menu Item
-----------------

Menu commands are mostly dispatch calls or a URL to run a macro. :py:class:`~ooodev.gui.menu.item.MenuItem` and :py:class:`~ooodev.gui.menu.item.MenuItemSub` have an execute method that will call call the dispatch or run the macro.

.. tabs::

    .. code-tab:: python

        from ooodev.gui.menu.item import MenuItemKind
        # ...
        menu = doc.menu[MenuLookupKind.TOOLS]
        itm = menu.items[".uno:AutoComplete"]
        if itm.item_kind >= MenuItemKind.ITEM:
            MenuItem.execute() # run the menu command

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

- :ref:`help_creating_menu_using_menu_app`
- :ref:`help_working_with_menu_bar`
- :ref:`help_working_with_shortcuts`