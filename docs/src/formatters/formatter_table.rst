.. _formatters_formatter_table:

Class FormatterTable
====================

Info
----

.. versionadded:: 0.6.6

Allow for complex formatting of Table Data.

.. tabs::

    .. code-tab:: python

        # Without Formatting
        >>> data = Calc.get_array(sheet=sheet, range_name="A1:E10")
        >>> Calc.print_array(data)

        Row x Column size: 10 x 5
        Stud. No.  Proj/20  Mid/35  Fin/45  Total%
        22001.0  16.4583333333333  30.9166666666667  37.0125  0.843875
        22028.0  11.875  23.0416666666667  25.4625  0.603791666666667
        22048.0  13.9583333333333  19.25  25.9875  0.591958333333333
        23715.0  12.0833333333333  18.6666666666667  20.475  0.51225
        23723.0  17.2916666666667  27.7083333333333  36.225  0.81225
        24277.0  0.0  16.0416666666667  19.6875  0.357291666666667
          11.9444444444444  22.6041666666667  27.475  0.620236111111111
          0.597222222222222  0.645833333333334  0.610555555555556
          Proj/20  Mid/35  Fin/45  Total%

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. tabs::

    .. code-tab:: python

        # With Formatting
        >>> data = Calc.get_array(sheet=sheet, range_name="A1:E10")
        >>> fl = FormatterTable(format=(".2f", ">9"), idxs=(0, 9))
        >>> fl.custom_formats_row.append(FormatTableItem(format=">9", idxs_inc=(0, 9)))
        >>> fl.col_formats.append(FormatTableItem(format=(".0f", "^9"), idxs_inc=(0,), row_idxs_exc=(0, 9)))
        >>> fl.col_formats.append(FormatTableItem(format=(".0%", ">9"), idxs_inc=(4,), row_idxs_exc=(0, 9)))
        >>> Calc.print_array(data, fl)

        Row x Column size: 10 x 5
        Stud. No.   Proj/20    Mid/35    Fin/45    Total%
          22001       16.46     30.92     37.01       84%
          22028       11.88     23.04     25.46       60%
          22048       13.96     19.25     25.99       59%
          23715       12.08     18.67     20.48       51%
          23723       17.29     27.71     36.23       81%
          24277        0.00     16.04     19.69       36%
                      11.94     22.60     27.48       62%
                       0.60      0.65      0.61
                    Proj/20    Mid/35    Fin/45    Total%

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. seealso::

    :py:meth:`.Calc.print_array`

Class
-----

.. autoclass:: ooodev.formatters.formatter_table.FormatterTable
    :members:
    :undoc-members: