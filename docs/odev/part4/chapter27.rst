.. _ch27:

***************************************
Chapter 27. Functions and Data Analysis
***************************************

.. topic:: Overview

    Calling Calc Functions from Code; Pivot Tables; Goal Seek; Linear and Nonlinear Solving (using ``SCO``, ``DEPS``)

    Examples: |fun_ex|_

This chapter looks at how to utilize Calc's spreadsheet functions directly from Java, and
then examines four of Calc's data analysis features: pivot tables, goal seeking, and linear and nonlinear solving.
There are two nonlinear examples, one using the ``SCO`` solver, the using employing ``DEPS``.

.. _ch27_calling_func_from_code:

27.1 Calling Calc Functions from Code
=====================================

Calc comes with an extensive set of functions, which are described in Appendix B of the Calc User Guide, available from `<https://libreoffice.org/get-help/documentation>`__.
The information is also online at `<https://help.libreoffice.org/Calc/Functions_by_Category>`__, organized into eleven categories:

1. | Database: for extracting information from Calc tables, where the data is organized into rows.
   | The "Database" name is a little misleading, but the documentation makes the point that
   | Calc database functions have nothing to do with Base databases. Chapter 13 of the
   | Calc User Guide ("Calc as a Simple Database") explains the distinction in detail.
2. Date and Time; :abbreviation:`i.e.` see the ``EASTERSUNDAY`` function below
3. Financial: for business calculations;
4. Information: many of these return boolean information about cells, such as whether a cell contains text or a formula;
5. Logical: functions for boolean logic;
6. Mathematical: trigonometric, hyperbolic, logarithmic, and summation functions; e.g. see ROUND, SIN, and RADIANS below;
7. Array: many of these operations treat cell ranges like 2D arrays; :abbreviation:`i.e.` see TRANSPOSE below;
8. Statistical: for statistical and probability calculations; :abbreviation:`i.e.`, see AVERAGE and SLOPE below;
9. Spreadsheet: for finding values in tables, cell ranges, and cells;
10. Text: string manipulation functions;
11. | Add-ins: a catch-all category that includes a lot of functions – extra data and time operations, conversion functions between number bases, more statistics, and complex numbers.
    | See IMSUM and ROMAN below for examples.
    | The "Add-ins" documentation starts at |calc_add_in|_, and continues in
    | `Add-in Functions, List of Analysis Functions Part One <https://help.libreoffice.org/latest/en-US/text/scalc/01/04060115.html>`__ and
    | `Add-in Functions, List of Analysis Functions Part Two <https://help.libreoffice.org/latest/en-US/text/scalc/01/04060116.html>`__.

A different organization for the functions documentation is used at the OpenOffice site (`Calc Functions listed by category <https://wiki.openoffice.org/wiki/Documentation/How_Tos/Calc:_Functions_listed_by_category>`__),
and is probably easy to use when browsing/searching for a suitable function.

If you know the name of the function, then a reasonably effective way of finding its documentation is to search for ``libreoffice calc function`` + the function name.

The standard way of using these functions is, of course, inside cell formulae.
But it's also possible to call them from code via the XFunctionAccess_ interface.
XFunctionAccess_ only contains a single function, ``callFunction()``, but it can be a bit hard to use due to data typing issues.

:py:meth:`.Calc.call_fun` creates an XFunctionAccess_ instance, and executes ``callFunction()``:

.. tabs::

    .. code-tab:: python

        # in Calc class
        @staticmethod
        def call_fun(func_name: str, *args: any) -> object:
            args_len = len(args)
            if args_len == 0:
                arg = ()
            else:
                arg = args
            try:
                fa = Lo.create_instance_mcf(
                    XFunctionAccess, "com.sun.star.sheet.FunctionAccess", raise_err=True
                )
                return fa.callFunction(func_name.upper(), arg)
            except Exception as e:
                Lo.print(f"Could not invoke function '{func_name.upper()}'")
                Lo.print(f"    {e}")
            return None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Calc.call_fun` is passed the Calc function name and and optionally an sequenence of arguments.
The function's result is returned as an Object instance.

Several examples of how to use :py:meth:`.Calc.call_fun` can be found in the |fun_ex|_ example:

.. tabs::

    .. code-tab:: python

        # in calc_functions.py
        def main(self) -> None:
            with Lo.Loader(Lo.ConnectPipe()) as loader:
                doc = Calc.create_doc(loader)
                sheet = Calc.get_sheet(doc=doc, index=0)
                # round
                print("ROUND result for 1.999 is: ", end="")
                print(Calc.call_fun("ROUND", 1.999))
                # more explained below.

                Lo.close(closeable=doc, deliver_ownership=False)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The printed result is:

::

    ROUND result for 1.999 is: 2.0

Function calls can be nested, as in:

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 3

        # in calc_functions.py
        print("SIN result for 30 degrees is:", end="")
        print(f'{Calc.call_fun("SIN", Calc.call_fun("RADIANS", 30)):.3f}')

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The call to ``RADIANS`` converts ``30`` degrees to radians.
The returned Object is accepted by the ``SIN`` function as input.
The output is: ``SIN`` result for ``30`` degrees is: ``0.500`` Many functions require more than one argument.

For instance:

.. tabs::

    .. code-tab:: python

        # in calc_functions.py
        avg = float(Calc.call_fun("AVERAGE", 1, 2, 3, 4, 5))
        print(f"Average of the numbers is: {avg:.1f}")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

This reports the average to be ``3.0``.

When the Calc function documentation talks about an "array" or "matrix" argument, then the data needs to be packaged as a 2D sequence such as list or tuple.
However for methods that were tested that required a matrix is showed that a list or tuple was not accepted.
What does work howerver, is writing the 2D data into a sheet and reading it back as XCellRange_ values.

For example, the ``SLOPE`` function takes two arrays of x and y coordinates as input, and calculates the slope of the line through them.
So first the 2D array is written to the sheet using :py:meth:`.Calc.set_array`.
Next the value are read from the sheet as XCellRange_ values into ``xrng`` and ``yrng``.
Now ``xrng`` and ``yrng`` can be passed to ``SLOPE``.

.. tabs::

    .. code-tab:: python

        # in calc_functions.py
        # the slope function only seems to work if passed XCellRange
        arr = [[1.0, 2.0, 3.0], [3.0, 6.0, 9.0]]
        Calc.set_array(values=arr, sheet=sheet, name="A1")
        Lo.delay(500)
        xrng = Calc.get_cell_range(sheet=sheet, range_name="A1:C1")
        yrng = Calc.get_cell_range(sheet=sheet, range_name="A2:C2")
        slope = float(Calc.call_fun("SLOPE", yrng, xrng))
        print(f"SLOPE of the line: {slope}")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The slope result is ``3.0``, as expected.


The functions in the "Array" category almost all use 2D arrays as arguments. For example, the ``TRANSPOSE`` function is called like so:

.. tabs::

    .. code-tab:: python

        # in calc_functions.py
        arr = [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]]
        Calc.set_array(values=arr, sheet=sheet, name="A1")
        Lo.delay(500)
        rng = Calc.get_cell_range(sheet=sheet, range_name="A1:C3")
        trans_mat = Calc.call_fun("TRANSPOSE", rng)
        # add a little extra formatting
        fl = FormatterTable(format=(".1f", ">5"))
        Calc.print_array(trans_mat, fl)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


The input array is in row-order, so the ``arr`` created above has two rows and three columns.
Extra fromatting is use by passing :py:meth:`.Calc.print_array` a :ref:`formatters_formatter_table` instance.
The printed transpose is:

::

    Row x Column size: 3 x 3
      1.0  4.0
      2.0  5.0
      3.0  6.0

Note that the result of this call to :py:meth:`.Calc.call_fun` is a 2D tuple.

There are several functions for manipulating imaginary numbers, which must be written in the form of strings.
For example, ``IMSUM`` sums a series of complex numbers like so:

.. tabs::

    .. code-tab:: python

        # in calc_functions.py
        # sum two imaginary numbers: "13+4j" + "5+3j" returns 18+7j.
        sum = Calc.call_fun("IMSUM", "13+4j", "5+3j")
        print(f"13+4j + 5+3j: {sum}")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The summed complex number is returned as the string ``"18+7j"``. This means that the :py:meth:`.Calc.call_fun` result is cast to String in this case.

.. _ch27_func_help:

Functions Help
--------------

If you can't access the Calc documentation on functions, then :py:class:`calc.Calc` contains two help functions: :py:meth:`.Calc.get_function_names` and :py:meth:`.Calc.print_function_info`.
The former prints a very long array of function names:

.. cssclass:: rst-collapse

    .. collapse:: List of 508 Functions

        ::

            Function Names
            No. of names: 508
              -------------------------|--------------------------|--------------------------|--------------------------
              ABS                      | ACCRINT                  | ACCRINTM                 | ACOS
              ACOSH                    | ACOT                     | ACOTH                    | ADDRESS
              AGGREGATE                | AMORDEGRC                | AMORLINC                 | AND
              ARABIC                   | AREAS                    | ASC                      | ASIN
              ASINH                    | ATAN                     | ATAN2                    | ATANH
              AVEDEV                   | AVERAGE                  | AVERAGEA                 | AVERAGEIF
              AVERAGEIFS               | B                        | BAHTTEXT                 | BASE
              BESSELI                  | BESSELJ                  | BESSELK                  | BESSELY
              BETA.DIST                | BETA.INV                 | BETADIST                 | BETAINV
              BIN2DEC                  | BIN2HEX                  | BIN2OCT                  | BINOM.DIST
              BINOM.INV                | BINOMDIST                | BITAND                   | BITLSHIFT
              BITOR                    | BITRSHIFT                | BITXOR                   | CEILING
              CEILING.MATH             | CEILING.PRECISE          | CEILING.XCL              | CELL
              CHAR                     | CHIDIST                  | CHIINV                   | CHISQ.DIST
              CHISQ.DIST.RT            | CHISQ.INV                | CHISQ.INV.RT             | CHISQ.TEST
              CHISQDIST                | CHISQINV                 | CHITEST                  | CHOOSE
              CLEAN                    | CODE                     | COLOR                    | COLUMN
              COLUMNS                  | COMBIN                   | COMBINA                  | COMPLEX
              CONCAT                   | CONCATENATE              | CONFIDENCE               | CONFIDENCE.NORM
              CONFIDENCE.T             | CONVERT                  | CONVERT_OOO              | CORREL
              COS                      | COSH                     | COT                      | COTH
              COUNT                    | COUNTA                   | COUNTBLANK               | COUNTIF
              COUNTIFS                 | COUPDAYBS                | COUPDAYS                 | COUPDAYSNC
              COUPNCD                  | COUPNUM                  | COUPPCD                  | COVAR
              COVARIANCE.P             | COVARIANCE.S             | CRITBINOM                | CSC
              CSCH                     | CUMIPMT                  | CUMIPMT_ADD              | CUMPRINC
              CUMPRINC_ADD             | CURRENT                  | DATE                     | DATEDIF
              DATEVALUE                | DAVERAGE                 | DAY                      | DAYS
              DAYS360                  | DAYSINMONTH              | DAYSINYEAR               | DB
              DCOUNT                   | DCOUNTA                  | DDB                      | DDE
              DEC2BIN                  | DEC2HEX                  | DEC2OCT                  | DECIMAL
              DEGREES                  | DELTA                    | DEVSQ                    | DGET
              DISC                     | DMAX                     | DMIN                     | DOLLAR
              DOLLARDE                 | DOLLARFR                 | DPRODUCT                 | DSTDEV
              DSTDEVP                  | DSUM                     | DURATION                 | DVAR
              DVARP                    | EASTERSUNDAY             | EDATE                    | EFFECT
              EFFECT_ADD               | ENCODEURL                | EOMONTH                  | ERF
              ERF.PRECISE              | ERFC                     | ERFC.PRECISE             | ERROR.TYPE
              ERRORTYPE                | EUROCONVERT              | EVEN                     | EXACT
              EXP                      | EXPON.DIST               | EXPONDIST                | F.DIST
              F.DIST.RT                | F.INV                    | F.INV.RT                 | F.TEST
              FACT                     | FACTDOUBLE               | FALSE                    | FDIST
              FILTERXML                | FIND                     | FINDB                    | FINV
              FISHER                   | FISHERINV                | FIXED                    | FLOOR
              FLOOR.MATH               | FLOOR.PRECISE            | FLOOR.XCL                | FORECAST
              FORECAST.ETS.ADD         | FORECAST.ETS.MULT        | FORECAST.ETS.PI.ADD      | FORECAST.ETS.PI.MULT
              FORECAST.ETS.SEASONALITY | FORECAST.ETS.STAT.ADD    | FORECAST.ETS.STAT.MULT   | FORECAST.LINEAR
              FORMULA                  | FOURIER                  | FREQUENCY                | FTEST
              FV                       | FVSCHEDULE               | GAMMA                    | GAMMA.DIST
              GAMMA.INV                | GAMMADIST                | GAMMAINV                 | GAMMALN
              GAMMALN.PRECISE          | GAUSS                    | GCD                      | GCD_EXCEL2003
              GEOMEAN                  | GESTEP                   | GETPIVOTDATA             | GROWTH
              HARMEAN                  | HEX2BIN                  | HEX2DEC                  | HEX2OCT
              HLOOKUP                  | HOUR                     | HYPERLINK                | HYPGEOM.DIST
              HYPGEOMDIST              | IF                       | IFERROR                  | IFNA
              IFS                      | IMABS                    | IMAGINARY                | IMARGUMENT
              IMCONJUGATE              | IMCOS                    | IMCOSH                   | IMCOT
              IMCSC                    | IMCSCH                   | IMDIV                    | IMEXP
              IMLN                     | IMLOG10                  | IMLOG2                   | IMPOWER
              IMPRODUCT                | IMREAL                   | IMSEC                    | IMSECH
              IMSIN                    | IMSINH                   | IMSQRT                   | IMSUB
              IMSUM                    | IMTAN                    | INDEX                    | INDIRECT
              INFO                     | INT                      | INTERCEPT                | INTRATE
              IPMT                     | IRR                      | ISBLANK                  | ISERR
              ISERROR                  | ISEVEN                   | ISEVEN_ADD               | ISFORMULA
              ISLEAPYEAR               | ISLOGICAL                | ISNA                     | ISNONTEXT
              ISNUMBER                 | ISO.CEILING              | ISODD                    | ISODD_ADD
              ISOWEEKNUM               | ISPMT                    | ISREF                    | ISTEXT
              JIS                      | KURT                     | LARGE                    | LCM
              LCM_EXCEL2003            | LEFT                     | LEFTB                    | LEN
              LENB                     | LINEST                   | LN                       | LOG
              LOG10                    | LOGEST                   | LOGINV                   | LOGNORM.DIST
              LOGNORM.INV              | LOGNORMDIST              | LOOKUP                   | LOWER
              MATCH                    | MAX                      | MAXA                     | MAXIFS
              MDETERM                  | MDURATION                | MEDIAN                   | MID
              MIDB                     | MIN                      | MINA                     | MINIFS
              MINUTE                   | MINVERSE                 | MIRR                     | MMULT
              MOD                      | MODE                     | MODE.MULT                | MODE.SNGL
              MONTH                    | MONTHS                   | MROUND                   | MULTINOMIAL
              MUNIT                    | N                        | NA                       | NEGBINOM.DIST
              NEGBINOMDIST             | NETWORKDAYS              | NETWORKDAYS.INTL         | NETWORKDAYS_EXCEL2003
              NOMINAL                  | NOMINAL_ADD              | NORM.DIST                | NORM.INV
              NORM.S.DIST              | NORM.S.INV               | NORMDIST                 | NORMINV
              NORMSDIST                | NORMSINV                 | NOT                      | NOW
              NPER                     | NPV                      | NUMBERVALUE              | OCT2BIN
              OCT2DEC                  | OCT2HEX                  | ODD                      | ODDFPRICE
              ODDFYIELD                | ODDLPRICE                | ODDLYIELD                | OFFSET
              OPT_BARRIER              | OPT_PROB_HIT             | OPT_PROB_INMONEY         | OPT_TOUCH
              OR                       | PDURATION                | PEARSON                  | PERCENTILE
              PERCENTILE.EXC           | PERCENTILE.INC           | PERCENTRANK              | PERCENTRANK.EXC
              PERCENTRANK.INC          | PERMUT                   | PERMUTATIONA             | PHI
              PI                       | PMT                      | POISSON                  | POISSON.DIST
              POWER                    | PPMT                     | PRICE                    | PRICEDISC
              PRICEMAT                 | PROB                     | PRODUCT                  | PROPER
              PV                       | QUARTILE                 | QUARTILE.EXC             | QUARTILE.INC
              QUOTIENT                 | RADIANS                  | RAND                     | RAND.NV
              RANDBETWEEN              | RANDBETWEEN.NV           | RANK                     | RANK.AVG
              RANK.EQ                  | RATE                     | RAWSUBTRACT              | RECEIVED
              REGEX                    | REPLACE                  | REPLACEB                 | REPT
              RIGHT                    | RIGHTB                   | ROMAN                    | ROT13
              ROUND                    | ROUNDDOWN                | ROUNDSIG                 | ROUNDUP
              ROW                      | ROWS                     | RRI                      | RSQ
              SEARCH                   | SEARCHB                  | SEC                      | SECH
              SECOND                   | SERIESSUM                | SHEET                    | SHEETS
              SIGN                     | SIN                      | SINH                     | SKEW
              SKEWP                    | SLN                      | SLOPE                    | SMALL
              SQRT                     | SQRTPI                   | STANDARDIZE              | STDEV
              STDEV.P                  | STDEV.S                  | STDEVA                   | STDEVP
              STDEVPA                  | STEYX                    | STYLE                    | SUBSTITUTE
              SUBTOTAL                 | SUM                      | SUMIF                    | SUMIFS
              SUMPRODUCT               | SUMSQ                    | SUMX2MY2                 | SUMX2PY2
              SUMXMY2                  | SWITCH                   | SYD                      | T
              T.DIST                   | T.DIST.2T                | T.DIST.RT                | T.INV
              T.INV.2T                 | T.TEST                   | TAN                      | TANH
              TBILLEQ                  | TBILLPRICE               | TBILLYIELD               | TDIST
              TEXT                     | TEXTJOIN                 | TIME                     | TIMEVALUE
              TINV                     | TODAY                    | TRANSPOSE                | TREND
              TRIM                     | TRIMMEAN                 | TRUE                     | TRUNC
              TTEST                    | TYPE                     | UNICHAR                  | UNICODE
              UPPER                    | VALUE                    | VAR                      | VAR.P
              VAR.S                    | VARA                     | VARP                     | VARPA
              VDB                      | VLOOKUP                  | WEBSERVICE               | WEEKDAY
              WEEKNUM                  | WEEKNUM_EXCEL2003        | WEEKNUM_OOO              | WEEKS
              WEEKSINYEAR              | WEIBULL                  | WEIBULL.DIST             | WORKDAY
              WORKDAY.INTL             | XIRR                     | XNPV                     | XOR
              YEAR                     | YEARFRAC                 | YEARS                    | YIELD
              YIELDDISC                | YIELDMAT                 | Z.TEST                   | ZTEST

If you know a function name, then :py:meth:`.Calc.print_function_info` will print details about it.

For instance, information about the ``ROMAN`` function is obtained like so:

.. tabs::

    .. code-tab:: python

        # in calc_functions.py
        Calc.print_function_info("ROMAN")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The output is:

::

    Properties for "ROMAN"":
      Id: 383
      Category: 10
      Name: ROMAN
      Description: Converts a number to a Roman numeral.
      Arguments: [Number, Mode (optional)]

    No. of arguments: 2
    1. Argument name: Number
      Description: 'The number to be converted to a Roman numeral must be in the 0 - 3999 range.'
      Is optional?: False

    2. Argument name: Mode
      Description: 'The more this value increases, the more the Roman numeral is simplified. The value must be in the 0 - 4 range.'
      Is optional?: True

This output states that ``ROMAN`` can be called with one or two arguments, the first being a decimal,
and the second an optional argument for the amount of 'simplification' carried out on the Roman numeral.
For example, here are two ways to convert ``999`` into Roman form:

.. tabs::

    .. code-tab:: python

        # in calc_functions.py
        # Roman numbers
        roman = Calc.call_fun("ROMAN", 999)
        # use max simplification
        roman4 = Calc.call_fun("ROMAN", 999, 4)
        print(f"999 in Roman numerals: {roman} or {roman4}")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The output is:

:: 

    999 in Roman numerals: CMXCIX or IM

:py:meth:`.Calc.get_function_names` and :py:meth:`.Calc.print_function_info` utilize the XFunctionDescriptions_ interface for retrieving an indexed container of function descriptions.
Each function description is an array of PropertyValue_ objects, which contain a ``Name`` property.
:py:meth:`.Calc.find_function` uses this organization to return a tuple of PropertyValue_ for a given function name:

.. tabs::

    .. code-tab:: python

        # in Calc class (simplified, overlaods)
        @staticmethod
        def find_function(func_nm: str) -> Tuple[PropertyValue] | None:
            if not func_nm:
                raise ValueError("Invalid arg, please supply a function name to find.")
            try:
                func_desc = Lo.create_instance_mcf(
                    XFunctionDescriptions, "com.sun.star.sheet.FunctionDescriptions", raise_err=True
                )
            except Exception as e:
                raise Exception("No function descriptions were found") from e

            for i in range(func_desc.getCount()):
                try:
                    props = cast(Sequence[PropertyValue], func_desc.getByIndex(i))
                    for p in props:
                        if p.Name == "Name" and str(p.Value) == func_nm:
                            return tuple(props)
                except Exception:
                    continue
            Lo.print(f"Function '{func_nm}' not found")
            return None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. only:: html

    .. seealso::

        .. cssclass:: src-link

            :odev_src_calc_meth:`find_function`

The tuple of PropertyValue_ contains five properties: ``Name``, ``Description``, ``Id``, ``Category``, and ``Arguments``.
The ``Arguments`` property stores an array of FunctionArgument_ objects which contain information about each argument's name, description, and whether it is optional.
This information is printed by :py:meth:`.Calc.print_fun_arguments`:

.. tabs::

    .. code-tab:: python

        # in Calc class
        @classmethod
        def print_fun_arguments(cls, prop_vals: Sequence[PropertyValue]) -> None:
            fargs = cast(
                "Sequence[FunctionArgument]", mProps.Props.get_value(name="Arguments", props=prop_vals)
            )
            if fargs is None:
                print("No arguments found")
                return

            print(f"No. of arguments: {len(fargs)}")
            for i, fa in enumerate(fargs):
                print(f"{i+1}. Argument name: {fa.Name}")
                print(f"  Description: '{fa.Description}'")
                print(f"  Is optional?: {fa.IsOptional}")
                print()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Calc.print_function_info` calls :py:meth:`.Calc.find_function` to report on a complete function:

.. tabs::

    .. code-tab:: python

        # in Calc class (simplified)
        @classmethod
        def print_function_info(cls, func_name: str) -> None:
            prop_vals = cls.find_function(func_name)
            if prop_vals is None:
                return
            Props.show_props(func_name, prop_vals)
            cls.print_fun_arguments(prop_vals)
            print()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch27_pivot_tables:

27.2 Pivot Tables
=================

Work in progress ...

.. |calc_add_in| replace:: Calc Add-in Functions
.. _calc_add_in: https://help.libreoffice.org/latest/en-US/text/scalc/01/04060111.html

.. |fun_ex| replace:: Calc Funcitons
.. _fun_ex: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_functions

.. |fun_ex_py| replace:: calc_functions.py
.. _fun_ex_py: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_functions/calc_functions.py

.. _XFunctionAccess: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XFunctionAccess.html
.. _XCellRange: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XCellRange.html
.. _XFunctionDescriptions: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XFunctionDescriptions.html
.. _PropertyValue: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1beans_1_1PropertyValue.html
.. _FunctionArgument: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1sheet_1_1FunctionArgument.html
