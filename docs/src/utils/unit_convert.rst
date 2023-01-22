.. _module_unit_convert:

Module unit_convert
===================

Convert from and to various units of ``Length`` enumeration.

.. code-block:: python

    >>> from ooodev.utils.unit_convert import UnitConvert, Length
    # convert 100 inches into milli-meters
    >>> UnitConvert.convert(100, frm=Length.IN, to=Length.MM)
    2540.0
    # convert to twips from 1/100 mm
    >>> UnitConvert.to_twips(10_000, Length.MM100)
    5669.291338582677


.. automodule:: ooodev.utils.unit_convert
    :members:
    :undoc-members: