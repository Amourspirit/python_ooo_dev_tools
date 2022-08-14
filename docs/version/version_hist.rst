###############
Version History
###############

**************
Version 0.4.12
**************

Add defaults for cfg in case config.json is not available.

**************
Version 0.4.11
**************

Fix bug in ``Lo.print_names()``

Remove internal events from some print functions that should not have had them.

Fix bug that did copy config.json during setup.

**************
Version 0.4.10
**************

Add new event_source property to internal event classes.

*************
Version 0.4.9
*************

| Added a Bridge Connector :py:attr:`.Lo.bridge`
| See also: :ref:`ch04_bridge_stop`
| See example: `Office Window Monitor <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_monitor>`_

Added Session class for registering and importing.
See example: `Shared Library Access <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_share_lib>`_

*************
Version 0.4.8
*************

New listeners in ooodev.listeners namespace

Fix For Lo.XSCRIPTCONTEXT

*************
Version 0.4.7
*************

Added ``minimize()``, ``maximize()`` and ``activate()`` methods to :py:class:`~.gui.GUI` class.

*************
Version 0.4.6
*************

Updates and fixes for :py:class:`~.utils.info.Info` class.


*************
Version 0.4.5
*************

Added :py:class:`~.break_context.BreakContext` class.

*************
Version 0.4.4
*************

Bug fix reading document properties.

*************
Version 0.4.2
*************

Fix bug in windows connections

*************
Version 0.4.1
*************

Fix bug in :py:class:`~.utils.info.Info`.
Some methods were expecting string but got Path object.

*************
Version 0.4.0
*************

New more flexable and robust way of connecting to office.

This update change :py:meth:`.Lo.load_office` method

Paths used internally now automatically resolve to absolute paths.

*************
Version 0.3.0
*************

Write module released

*************
Version 0.2.0
*************

Initial release with full support for calc.