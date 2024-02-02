.. code-block:: python

    from ooodev.loader.lo import Lo
    from ooodev.office.calc import Calc

    doc = Calc.get_ss_doc(Lo.ThisComponent)
    # do work ...

