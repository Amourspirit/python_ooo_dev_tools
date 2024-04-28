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

App Menus are used to add, remove, and modify menus in the LibreOffice on both a Global and Document level.

:ref:`More Information <help_common_gui_menus_app_menu>`.

Quick Start
^^^^^^^^^^^

See :ref:`help_app_menu_about_example`.

Popup Menu
-----------

Popup Menus are not persisted after a restart.
Popup menus can be used to add additional menus to the menu bar or to create a new popup menu base upon some other action.

:ref:`More Information <help_common_gui_menus_popup_menu>`.

Quick Start
^^^^^^^^^^^

See :ref:`help_popup_menu_about_example`.

Menu Bar Menu
--------------

Menu Bar menus are menus that are added to the main menu bar at the top of the application.
Menus are added by creating a new Popup Menu and then adding it to the menu bar.
For this reason working with Menu Bar is much like working with Popup Menus.
:ref:`More Information <help_common_gui_menus_menu_bar>`.

Quick Start
^^^^^^^^^^^

See :ref:`help_menubar_menu_about_example`.

Context Action Menu
-------------------

Context Action Menu are menus are are inserted into existing context menus by incepting the menu and modifying it.
:ref:`More Information <help_common_gui_menus_context>`.

Quick Start
^^^^^^^^^^^

See :ref:`help_context_menu_about_example`.