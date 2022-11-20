***************
Version History
***************

Version 0.6.6
=============

Add overload to ``Calc.convert_to_floats``

Version 0.6.5
=============

Added overload to ``FileIo.make_directory`` that handles creating directory from file path.

Fix for ``FileIo.url_to_path`` on windows sometimes not converting correctly.

Other ``FileIo`` Minor updates.

Fix bug in ``Chart2.set_template`` when ``diagram_name`` was passed as string.

Fix bug in ``Draw.warns_position`` when no Slide size is available.

Renamed ``Calc.get_range_str`` args from ``start_col``, ``start_row``, ``end_col``, ``end_row`` to ``col_start``, ``row_start``, ``col_end``, ``row_end`` respectively.
Change is backwards compatible.

Renamed ``Calc.get_cell_range`` args from ``start_col``, ``start_row``, ``end_col``, ``end_row`` to ``col_start``, ``row_start``, ``col_end``, ``row_end`` respectively.
Change is backwards compatible.

Version 0.6.4
=============

Fix for ``Draw.report_pos_size``. Now handles when a shape does not have a ``Name`` property an other errors.

Version 0.6.3
=============

Overloads for ``GUI.get_window_handle()``

Removed unused ``*titles`` arg from ``Draw.add_dispatch_shape()`` method.

Removed unused ``*titles`` arg from ``Draw.create_dispatch_shape()`` method.

``GUI.get_title_bar()`` method now returns empty string when not able to get title bar text.

Version 0.6.2
=============

Rename private enum ``_LayoutKind`` to public ``LayoutKind`` to make available for public use.

Added new Fast Lookup methods to ``Props`` class.

New Exceptions ``PropertyGeneralError``

Version 0.6.1
=============

Added ``Draw.add_dispatch_shape()`` method.

Added ``Draw.create_dispatch_shape()`` method.

Added Dispatch Lookup ``ShapeDispatchKind`` Enum.

Added None to ``GraphicArrowStyleKind`` Enum.

Added classes ``WindowTitle`` and ``DialogTitle`` for working with GUI packages.

Version 0.6.0
=============

Breaking changes.

``Write.ControlCharacter`` was an alias of ``ooo.dyn.text.control_character.ControlCharacterEnum``.
Now ``ControlCharacterEnum`` must be used instead of ``Write.ControlCharacter``.
``ControlCharacterEnum`` can be imported from ``Write``.
:abbreviation:`e.g.` ``from ooodev.office.write import Write, ControlCharacterEnum``

``Write.DictionaryType`` was an alias of ``ooo.dyn.linguistic2.dictionary_type.DictionaryType``.
Now ``DictionaryType`` must be used instead of ``Write.DictionaryType``.
``DictionaryType`` can be imported from ``Write``.
:abbreviation:`e.g.` ``from ooodev.office.write import Write, DictionaryType``

``Calc.CellFlags`` was an alias of ``ooo.dyn.sheet.cell_flags.CellFlagsEnum``.
Now ``CellFlagsEnum`` must be used instead of ``Calc.CellFlags``.
``CellFlagsEnum`` can be imported from ``Calc``.
:abbreviation:`e.g.` ``from ooodev.office.calc import Calc, CellFlagsEnum``

``Calc.GeneralFunction`` was an alias of ``ooo.dyn.sheet.general_function.GeneralFunction``.
Now ``GeneralFunction`` must be used instead of ``Calc.GeneralFunction``.
``GeneralFunction`` can be imported from ``Calc``.
:abbreviation:`e.g.` ``from ooodev.office.calc import Calc, GeneralFunction``

``Calc.SolverConstraintOperator`` was an alias of ``ooo.dyn.sheet.solver_constraint_operator.SolverConstraintOperator``.
Now ``SolverConstraintOperator`` must be used instead of ``Calc.SolverConstraintOperator``.
``SolverConstraintOperator`` can be imported from ``Calc``.
:abbreviation:`e.g.` ``from ooodev.office.calc import Calc, SolverConstraintOperator``


``Calc.FillDateMode`` was an alias of ``ooo.dyn.sheet.fill_date_mode.FillDateMode``.
Now ``FillDateMode`` must be used instead of ``Calc.FillDateMode``.
``FillDateMode`` can be imported from ``Calc``.
:abbreviation:`e.g.` ``from ooodev.office.calc import Calc, FillDateMode``

Version 0.5.3
=============

``Lo.dispatch_cmd`` Now returns the result of the dispatch command if any.
Formerly a ``bool`` was returned.

``Lo.dispatch_cmd`` Now raises ``DispatchError`` if an error occurs.

Version 0.5.2
=============

Chart Samples and tests

Misc code tweaks.

Version 0.5.1
=============

Chart 2 Samples and tests.

Version 0.5.0
=============

New modules
 - Draw
 - Chart
 - Chart2

Added ``utils.dispatch`` which as several new classes for looking up dispatch values.

Misc bug fixes.

Version 0.4.19
==============

Fix bug in setup.py

Version 0.4.17
==============

Update to Write:

- new method ``split_paragraph_into_sentences``
- new overloads for ``print_meaning``
- new overloads for ``print_services_info``
- new overloads for ``proof_sentence``
- new overloads for ``spell_sentence``
- new overloads for ``spell_word``
- ``load_spell_checker`` now load spell checker from ``com.sun.star.linguistic2.SpellChecker``


Version 0.4.16
==============

Fixes for Write spell checking


Version 0.4.15
==============

Update Graphic methods to move away from ``GraphicURL``

Other minor bug fixes.

Version 0.4.14
==============

Minor fix in ``Write.set_page_numbers``

Version 0.4.13
==============

Fix for  ``Write.add_text_frame()`` events.

Version 0.4.12
==============

Add defaults for cfg in case config.json is not available.

Version 0.4.11
==============

Fix bug in ``Lo.print_names()``

Remove internal events from some print functions that should not have had them.

Fix bug that did copy config.json during setup.

Version 0.4.10
==============

Add new event_source property to internal event classes.

Version 0.4.9
=============

| Added a Bridge Connector :py:attr:`.Lo.bridge`
| See also: :ref:`ch04_bridge_stop`
| See example: `Office Window Monitor <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_monitor>`_

Added Session class for registering and importing.
See example: `Shared Library Access <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_share_lib>`_

Version 0.4.8
=============

New listeners in ooodev.listeners namespace

Fix For Lo.XSCRIPTCONTEXT

Version 0.4.7
=============

Added ``minimize()``, ``maximize()`` and ``activate()`` methods to :py:class:`~.gui.GUI` class.

Version 0.4.6
=============

Updates and fixes for :py:class:`~.utils.info.Info` class.


Version 0.4.5
=============

Added :py:class:`~.break_context.BreakContext` class.

Version 0.4.4
=============

Bug fix reading document properties.

Version 0.4.2
=============

Fix bug in windows connections

Version 0.4.1
=============

Fix bug in :py:class:`~.utils.info.Info`.
Some methods were expecting string but got Path object.

Version 0.4.0
=============

New more flexable and robust way of connecting to office.

This update change :py:meth:`.Lo.load_office` method

Paths used internally now automatically resolve to absolute paths.

Version 0.3.0
=============

Write module released

Version 0.2.0
=============

Initial release with full support for calc.