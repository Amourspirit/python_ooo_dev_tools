.. _help_menus_introduction:

Menu Introduction
=================

Introduction
------------

This Library support several types of menus.
Each menu type supports loading and saving menu data to and from JSON.
Some menus can persist data into LibreOffice and some are only temporary.

Depending on you needs working with the menus can be very simple or rather complex.

App Menu
---------

App menus are generally used when there is a need to persist the menu into LibreOffice.
Although there are times that using App Menu with Popup menu will make sense.

Quick Start
^^^^^^^^^^^

See :ref:`help_app_menu_about_example`.

Popup Menu
-----------

Popup Menus are not persisted after a restart.
Popup menus can be used to add additional menus to the menu bar or to create a new popup menu base upon some other action.

Quick Start
^^^^^^^^^^^

See :ref:`help_popup_menu_about_example`.

Menu Bar Menu
--------------

Menu Bar menus are menus that are added to the main menu bar at the top of the application.
Menus are added by creating a new Popup Menu and then adding it to the menu bar.
For this reason working with Menu Bar is much like working with Popup Menus.
See :ref:`help_working_with_menu_bar` for more information.

Quick Start
^^^^^^^^^^^

See :ref:`help_menubar_menu_about_example`.

Context Action Menu
-------------------

Context Action Menu are menus are are inserted into existing context menus by incepting the menu and modifying it.
See :ref:`help_menu_context_incept` for more information.

Quick Start
^^^^^^^^^^^

See :ref:`help_context_menu_about_example`.