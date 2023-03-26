.. _ch10:

*******************************
Chapter 10. The Linguistics API
*******************************

.. topic:: Overview

    Linguistic Tools; Using the Spell Checker; Using the Thesaurus; Grammar Checking; Guessing the Language used in a String; Spell Checking and Grammar Checking a Document

The linguistics API has four main components:

1. spell checker
2. hyphenator
3. thesaurus
4. grammar checker (which Office calls a proof reader).

We'll look at how to program using the spell checker, thesaurus, and two grammar checkers, but skip the hyphenator which is easier to use interactively through Office's GUI.

However, if you have an urge to hyphenate, then |lingustic_ex|_ in the Developer's Guide examples contains some code;
it can be downloaded from |lingustic_ex|_.


We'll describe two examples, Lingo_ and |lingo_file|_.
The first lists information about the linguistic services, then uses the spell checker, thesaurus, and grammar checker 'standalone' without having to load an Office document first.
|lingo_file|_ automatically spell checks and grammar checks a complete file, reporting errors without altering the document.

One topic I'll be ignoring is how to create and edit the data files used by the linguistic services.
For that task, you should have a look at PTG (Proofing Tool GUI) developed by Marco Pinto at https://www.proofingtoolgui.org/.
It's an open source tool for editing Office's dictionary, thesaurus, hyphenation, and autocorrect files.

Another area that is skipped here is the use of events and listeners.
Please refer to the "Linguistics" sub-section of chapter 6 of the Developer's Guide for details (``loguide Linguistics``).
Listener code can be found in |lingustic_ex|_ mentioned above.
Also see :ref:`ch04`.

The linguistic features accessible through Office's GUI are explained in chapter 3 of the "Writer Guide", available at https://libreoffice.org/get-help/documentation,
starting from the section called "Checking spelling and grammar".

An older information source is the ``Lingucomponent Project`` page at the OpenOffice website,
https://openoffice.org/lingucomponent, which links to some useful tools, such as alternative grammar checkers.
An interesting set of slides by Daniel Naber explaining the state of the project in 2005 can be found at http://danielnaber.de/publications/, along with more recent material.

.. _ch10_linguistic_tools:

10.1 The Linguistic Tools
=========================

Lingo_ example prints a variety of information about the linguistics services:

.. tabs::

    .. code-tab:: python

        def main() -> int:

            with Lo.Loader(Lo.ConnectSocket(headless=True)) as loader:

                # print linguistics info
                Write.dicts_info()

                lingu_props = Write.get_lingu_properties()
                Props.show_props("Linguistic Manager", lingu_props)

                Info.list_extensions()  # these include linguistic extensions

                lingo_mgr = Lo.create_instance_mcf(XLinguServiceManager2, "com.sun.star.linguistic2.LinguServiceManager")
                if lingo_mgr is None:
                    print("No linguistics manager found")
                    return 0

                Write.print_services_info(lingo_mgr)

                # : code for using the services; see later

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch10_dict_info:

10.1.1 Dictionary Information
-----------------------------

:py:meth:`.Write.dicts_info` prints brief details about Office's dictionaries:

.. code-block:: text

    No. of dictionaries: 5
      standard.dic (1); active; ""; positive
      en-GB.dic (42); active; "GB"; positive
      en-US.dic (42); active; "US"; positive
      technical.dic (258); active; ""; positive
      IgnoreAllList (0); active; ""; positive

    No. of conversion dictionaries: 0

Each line includes the name of a dictionary, its number of entries, whether it's active (i.e. being used), its locale, and whether it's a positive, negative, or mixed dictionary.

A positive dictionary holds correctly spelled words only, while a negative one lists incorrectly spelled words.
A mixed dictionary contains both correctly and incorrectly spelled entries.

If a dictionary has a locale, such as "GB" for ``en-GB.dic``, then it's only utilized during spell checking if its locale matches Office's.
The Office locale can be set via the Tools, Options, Language Settings, "Languages" dialog shown in :numref:`ch10fig_lang_dial_ss`.

..
    figure 1

.. cssclass:: screen_shot invert

    .. _ch10fig_lang_dial_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/186284804-cc04946a-ac3f-4581-b295-2b71491763af.png
        :alt: Screen shot of The Languages Dialog
        :figclass: align-center

        :The Languages Dialog.

 :numref:`ch10fig_lang_dial_ss` shows that my version of Office is using the American English locale, and so ``en-GB.dic`` won't be consulted when text is spell checked.

:py:meth:`.Write.dicts_info` is defined as:

.. tabs::

    .. code-tab:: python

        @classmethod
        def dicts_info(cls) -> None:
            dict_lst = Lo.create_instance_mcf(XSearchableDictionaryList, "com.sun.star.linguistic2.DictionaryList")
            if not dict_lst:
                print("No list of dictionaries found")
                return
            cls.print_dicts_info(dict_lst)

            cd_list = mLo.Lo.create_instance_mcf(
                XConversionDictionaryList, "com.sun.star.linguistic2.ConversionDictionaryList"
            )
            if cd_list is None:
                print("No list of conversion dictionaries found")
                return
            cls.print_con_dicts_info(cd_list)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

It retrieves a conventional dictionary list first (called ``dict_lst``), and iterates through its dictionaries using :py:meth:`~.Write.print_con_dicts_info`.
Then it obtains the conversion dictionary list (called ``cd_list``), and iterates over that with :py:meth:`~.Write.print_con_dicts_info`.

:numref:`ch09fig_dicts_services` shows the main services and interfaces used by ordinary dictionaries.

..
    figure 2

.. cssclass:: diagram invert

    .. _ch09fig_dicts_services:
    .. figure:: https://user-images.githubusercontent.com/4193389/186043401-8c5b5ac4-0620-4fd0-b0b5-a328521ec64c.png
        :alt: Diagram of Dictionary List and Dictionary Services.
        :figclass: align-center

        :The DictionaryList_ and Dictionary_ Services.

Each dictionary in the list has an XDictionary_ interface which contains methods for accessing and changing its entries.
:py:meth:`~.Write.print_dicts_info` retrieves an XDictionary_ sequence from the list, and prints out a summary of each dictionary:

.. tabs::

    .. code-tab:: python

        @classmethod
        def print_dicts_info(cls, dict_list: XSearchableDictionaryList) -> None:
            if dict_list is None:
                print("Dictionary list is null")
                return
            print(f"No. of dictionaries: {dict_list.getCount()}")
            dicts = dict_list.getDictionaries()
            for d in dicts:ch10fig_convert_dicts_services
                print(
                    f"  {d.getName()} ({d.getCount()}); ({'active' if d.isActive() else 'na'}); '{d.getLocale().Country}'; {cls.get_dict_type(d.getDictionaryType())}"
                )
            print()

        @staticmethod
        def get_dict_type(dt: Write.DictionaryType) -> str:
              if dt == Write.DictionaryType.POSITIVE:
                return "positive"
            if dt == Write.DictionaryType.NEGATIVE:
                return "negative"
            if dt == Write.DictionaryType.MIXED:
                return "mixed"
            return "??"

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Conversion dictionaries map words in one language/dialect to corresponding words in another language/dialect.
:numref:`ch10fig_convert_dicts_services` shows that conversion dictionaries are organized in a similar way to ordinary ones.
The interfaces for manipulating a conversion dictionary are XConversionDictionary_ and XConversionPropertyType_.

..
    figure 3

.. cssclass:: diagram invert

    .. _ch10fig_convert_dicts_services:
    .. figure:: https://user-images.githubusercontent.com/4193389/186044163-06e65425-a158-4a1e-a28c-17faad1b8e84.png
        :alt: Diagram of the The Conversion Dictionary List and Conversion Dictionary Services.
        :figclass: align-center

        :The ConversionDictionaryList_ and ConversionDictionary_ Services.

:py:meth:`.Write.dicts_info` calls :py:meth:`~.Write.print_con_dicts_info` to print the names of the conversion dictionaries – by extracting
an XNameContainer_ from the dictionary list, and then pulling a list of the names from the container:

.. tabs::

    .. code-tab:: python

        @staticmethod
        def print_con_dicts_info(cd_lst: XConversionDictionaryList) -> None:
            if cd_lst is None:
                print("Conversion Dictionary list is null")
                return

            dc_con = cd_lst.getDictionaryContainer()
            dc_names = dc_con.getElementNames()
            print(f"No. of conversion dictionaries: {len(dc_names)}")
            for name in dc_names:
                print(f"  {name}")
            print()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Output similar to :py:meth:`.Write.dicts_info` can be viewed via Office's Tools, Options, Language Settings, "Writing Aids" dialog, shown in :numref:`ch10fig_writing_aids_ss`.

..
    figure 4

.. cssclass:: screen_shot invert

    .. _ch10fig_writing_aids_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/186285125-c619c721-2491-4c67-82fe-5c5af400f173.png
        :alt: Screen shot of The Writing Aids Dialog
        :figclass: align-center

        :The Writing Aids Dialog.

The dictionaries are listed in the second pane of the dialog.
Also, at the bottom of the window is a "Get more dictionaries online" hyperlink which takes the user to Office's extension website, and displays the "Dictionary" category (see :numref:`ch10fig_ext_dict_ss`).

..
    figure 5

.. cssclass:: screen_shot invert

    .. _ch10fig_ext_dict_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/186045815-bf013c3c-aff4-429e-b71f-92675b68b884.png
        :alt: Screen shot of The Dictionary Extensions at the LibreOffice Website
        :figclass: align-center

        :The Dictionary Extensions at the LibreOffice Website.

The URL of the page in :numref:`ch10fig_ext_dict_ss` is: https://extensions.libreoffice.org/?Tags%5B%5D=50.
If you can't find what you're looking for, don't forget the extensions for OpenOffice, at: https://extensions.openoffice.org
If you're unclear about how to install extensions, the process is explained online at https://wiki.documentfoundation.org/Documentation/HowTo/install_extension,
or in the "Installing Extensions" guide available at https://libreoffice.org/get-help/documentation.

.. _ch10_linguistic_props:

10.1.2 Linguistic Properties
----------------------------

Back in the Lingo_ example, :py:meth:`.Write.get_lingu_properties` returns an instance of XLinguProperties_,
and its properties are printed by calling :py:meth:`.Props.show_props`:

.. tabs::

    .. code-tab:: python

        # code fragment from lingo example
        lingu_props = Write.get_lingu_properties()
        Props.show_props("Linguistic Manager", lingu_props)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The output:

.. code-block:: text

    Linguistic Manager Properties
      DefaultLanguage: 0
      DefaultLocale: (com.sun.star.lang.Locale){ Language = (string)"", Country = (string)"", Variant = (string)"" }
      DefaultLocale_CJK: (com.sun.star.lang.Locale){ Language = (string)"", Country = (string)"", Variant = (string)"" }
      DefaultLocale_CTL: (com.sun.star.lang.Locale){ Language = (string)"", Country = (string)"", Variant = (string)"" }
      HyphMinLeading: 2
      HyphMinTrailing: 2
      HyphMinWordLength: 5
      IsGermanPreReform: None
      IsHyphAuto: False
      IsHyphSpecial: True
      IsIgnoreControlCharacters: True
      IsSpellAuto: True
      IsSpellCapitalization: True
      IsSpellHide: None
      IsSpellInAllLanguages: None
      IsSpellSpecial: True
      IsSpellUpperCase: True
      IsSpellWithDigits: False
      IsUseDictionaryList: True
      IsWrapReverse: False

These properties are explained in the online documentation for the XLinguProperties_ interface (``lodoc XLinguProperties``), and also in the Developer's Guide.

The properties are spread across several dialog in Office's GUI, starting from the Tools, Options, "Language Settings" menu item.
However, most of them are in the "Options" pane of the "Writing Aids" Dialog in :numref:`ch10fig_writing_aids_ss`.

.. _ch10_installed_ext:

10.1.3 Installed Extensions
---------------------------

Additional dictionaries, and other language tools such as grammar checkers, are loaded into Office as extensions, so calling :py:meth:`.Info.list_extensions` can be informative.

The output on one of my test machine is:

.. code-block:: text

    Extensions:
    1. ID: apso.python.script.organizer
       Version: 1.3.0
       Loc: file:///C:/Users/bigby/AppData/Roaming/LibreOffice/4/user/uno_packages/cache/uno_packages/lu1271241oyk.tmp_/apso.oxt

    2. ID: org.openoffice.en.hunspell.dictionaries
       Version: 2021.11.01
       Loc: file:///C:/Program%20Files/LibreOffice/program/../share/extensions/dict-en

    3. ID: French.linguistic.resources.from.Dicollecte.by.OlivierR
       Version: 7.0
       Loc: file:///C:/Program%20Files/LibreOffice/program/../share/extensions/dict-fr

    4. ID: org.openoffice.languagetool.oxt
       Version: 5.8
       Loc: file:///C:/Users/bigby/AppData/Roaming/LibreOffice/4/user/uno_packages/cache/uno_packages/lu107803j3h0.tmp_/LanguageTool-stable.oxt

    5. ID: com.sun.star.comp.Calc.NLPSolver
       Version: 0.9
       Loc: file:///C:/Program%20Files/LibreOffice/program/../share/extensions/nlpsolver

    6. ID: spanish.es.dicts.from.rla-es
       Version: __VERSION__
       Loc: file:///C:/Program%20Files/LibreOffice/program/../share/extensions/dict-es

    7. ID: com.sun.wiki-publisher
       Version: 1.2.0
       Loc: file:///C:/Program%20Files/LibreOffice/program/../share/extensions/wiki-publisher

The ``Loc`` entries are the directories or OXT files containing the extensions. Most extensions are placed in the share extensions folder on Windows.

Office can display similar information via its Tools, "Extension Manager" dialog, as in :numref:`ch10fig_ext_dial_ss`.

..
    figure 6

.. cssclass:: screen_shot invert

    .. _ch10fig_ext_dial_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/186285373-d9375dc6-c544-476c-bdb1-72754810546f.png
        :alt: Screen shot of The Extension Manager Dialog.
        :figclass: align-center

        :The Extension Manager Dialog.

The code for :py:meth:`.Info.list_extensions`:

.. tabs::

    .. code-tab:: python

        @classmethod
        def list_extensions(cls) -> None:
            try:
                pip = cls.get_pip()
            except MissingInterfaceError:
                print("No package info provider found")
                return
            exts_tbl = pip.getExtensionList()
            print("\nExtensions:")
            for i in range(len(exts_tbl)):
                print(f"{i+1}. ID: {exts_tbl[i][0]}")
                print(f"   Version: {exts_tbl[i][1]}")
                print(f"   Loc: {pip.getPackageLocation(exts_tbl[i][0])}")
                print()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Extensions are accessed via the XPackageInformationProvider_ interface.

.. _ch10_exam_lingu:

10.1.4 Examining the Lingu Services
-----------------------------------

The LinguServiceManager_ provides access to three of the four main linguistic services: the spell checker, the hyphenator, and thesaurus.
The proof reader (:abbreviation:`ex:` the grammar checker) is managed by a separate Proofreader_ service, which is explained later.

:numref:`ch10fig_longu_serv_interface` shows the interfaces accessible from the LinguServiceManager service.

..
    figure 7

.. cssclass:: diagram invert

    .. _ch10fig_longu_serv_interface:
    .. figure:: https://user-images.githubusercontent.com/4193389/186255983-5ed8f694-3bcc-4fce-874b-a860b1deef9d.png
        :alt: Diagram of The Lingu Service Manager Service and Interfaces.
        :figclass: align-center

        :The LinguServiceManager_ Service and Interfaces.

In Lingo_ example, the LinguServiceManager_ is instantiated and then :py:meth:`.Write.print_services_info` reports details about its services:

.. tabs::

    .. code-tab:: python

        # in lingo example

        # get lingo manager
        lingo_mgr = Lo.create_instance_mcf(XLinguServiceManager2, "com.sun.star.linguistic2.LinguServiceManager")
        if lingo_mgr is None:
            print("No linguistics manager found")
            return 0

        Write.print_services_info(lingo_mgr)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Typical output from :py:meth:`.Write.print_services_info`:

.. code-block:: text

    Available Services:
    SpellChecker (1):
      org.openoffice.lingu.MySpellSpellChecker
    Thesaurus (1):
      org.openoffice.lingu.new.Thesaurus
    Hyphenator (1):
      org.openoffice.lingu.LibHnjHyphenator
    Proofreader (2):
      org.languagetool.openoffice.Main
      org.libreoffice.comp.pyuno.Lightproof.en

    Configured Services:
    SpellChecker (1):
      org.openoffice.lingu.MySpellSpellChecker
    Thesaurus (1):
      org.openoffice.lingu.new.Thesaurus
    Hyphenator (1):
      org.openoffice.lingu.LibHnjHyphenator
    Proofreader (1):
      org.libreoffice.comp.pyuno.Lightproof.en

    Locales for SpellChecker (46)
      AR  AU  BE  BO  BS  BZ  CA  CA  CH  CL
      CO  CR  CU  DO  EC  ES  FR  GB  GH  GQ
      GT  HN  IE  IN  JM  LU  MC  MW  MX  NA
      NI  NZ  PA  PE  PH  PH  PR  PY  SV  TT
      US  US  UY  VE  ZA  ZW

    Locales for Thesaurus (46)
      AR  AU  BE  BO  BS  BZ  CA  CA  CH  CL
      CO  CR  CU  DO  EC  ES  FR  GB  GH  GQ
      GT  HN  IE  IN  JM  LU  MC  MW  MX  NA
      NI  NZ  PA  PE  PH  PH  PR  PY  SV  TT
      US  US  UY  VE  ZA  ZW

    Locales for Hyphenator (46)
      AR  AU  BE  BO  BS  BZ  CA  CA  CH  CL
      CO  CR  CU  DO  EC  ES  FR  GB  GH  GQ
      GT  HN  IE  IN  JM  LU  MC  MW  MX  NA
      NI  NZ  PA  PE  PH  PH  PR  PY  SV  TT
      US  US  UY  VE  ZA  ZW

    Locales for Proofreader (111)
      AE  AF  AO  AR  AT  AU  BE  BE  BE  BH
      BO  BR  BS  BY  BZ  CA  CA  CD  CH  CH
      CH  CI  CL  CM  CN  CR  CU  CV  DE  DE
      DK  DO  DZ  EC  EG  ES  ES  ES  ES  ES
      FI  FR  FR  GB  GH  GR  GT  GW  HN  HT
      IE  IE  IN  IN  IQ  IR  IT  JM  JO  JP
      KH  KW  LB  LI  LU  LU  LY  MA  MA  MC
      ML  MO  MX  MZ  NA  NI  NL  NZ  OM  PA
      PE  PH  PH  PL  PR  PT  PY  QA  RE  RO
      RU  SA  SD  SE  SI  SK  SN  ST  SV  SY
      TL  TN  TT  UA  US  US  UY  VE  YE  ZA
      ZW

The print-out contains three lists: a list of available services, a list of configured services (i.e. ones that are activated inside Office),
and a list of the locales available to each service.

:numref:`ch10fig_longu_serv_interface` shows that LinguServiceManager_ only manages the spell checker, hyphenator, and thesaurus, and yet :py:meth:`.Write.print_services_info`
includes information about the proof reader. Somewhat confusingly, although LinguServiceManager_ cannot instantiate a proof reader it can print information about it.

The output shows that two ``proofreader`` services are available (``org.languagetool.openoffice.Main`` and ``org.libreoffice.comp.pyuno.Lightproof.en``), but only one is configured (i.e. active).
This setup is explained  when we talk about the proof reader later.

The three lists are generated by :py:meth:`.Write.print_services_info` calling :py:meth:`.Write.print_avail_service_info`, :py:meth:`.Write.print_config_service_info`, and :py:meth:`.Write.print_locales`:

.. tabs::

    .. code-tab:: python

        @classmethod
        def print_services_info(cls, lingo_mgr: XLinguServiceManager2, loc: Locale | None = None) -> None:
            if loc is None:
                loc = Locale("en", "US", "")
            print("Available Services:")
            cls.print_avail_service_info(lingo_mgr, "SpellChecker", loc)
            cls.print_avail_service_info(lingo_mgr, "Thesaurus", loc)
            cls.print_avail_service_info(lingo_mgr, "Hyphenator", loc)
            cls.print_avail_service_info(lingo_mgr, "Proofreader", loc)
            print()

            print("Configured Services:")
            cls.print_config_service_info(lingo_mgr, "SpellChecker", loc)
            cls.print_config_service_info(lingo_mgr, "Thesaurus", loc)
            cls.print_config_service_info(lingo_mgr, "Hyphenator", loc)
            cls.print_config_service_info(lingo_mgr, "Proofreader", loc)
            print()

            cls.print_locales("SpellChecker", lingo_mgr.getAvailableLocales("com.sun.star.linguistic2.SpellChecker"))
            cls.print_locales("Thesaurus", lingo_mgr.getAvailableLocales("com.sun.star.linguistic2.Thesaurus"))
            cls.print_locales("Hyphenator", lingo_mgr.getAvailableLocales("com.sun.star.linguistic2.Hyphenator"))
            cls.print_locales("Proofreader", lingo_mgr.getAvailableLocales("com.sun.star.linguistic2.Proofreader"))
            print()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The choice of services depends on the current locale by default, so :py:meth:`.Write.print_services_info` begins by creating an American English locale, which matches my version of Office.
:py:meth:`.Write.print_services_info` can also take a Locale as an option.

:py:meth:`.Write.print_avail_service_info` utilizes ``XLinguServiceManager.getAvailableServices()`` to retrieve a list of the available services.
In a similar way, :py:meth:`.Write.print_config_service_info` calls ``XLinguServiceManager.getConfiguredServices()``,
and :py:meth:`.Write.print_locales` gets a sequence of Locale objects from ``XLinguServiceManager.getAvailableLocales()``.


.. _ch10_use_spell_check:

10.2 Using the Spell Checker
============================

There's a few examples in Lingo_ example of applying the spell checker to individual words:

.. tabs::

    .. code-tab:: python

        # in lingo example
        # use spell checker
        Write.spell_word("horseback", speller)
        Write.spell_word("ceurse", speller)
        Write.spell_word("magisian", speller)
        Write.spell_word("ellucidate", speller)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

``XLinguServiceManager.getSpellChecker()`` returns a reference to the spell checker, and :py:meth:`.Write.spell_word` checks the supplied word.
For the code above, the following is printed:

.. code-block:: text

    * "ceurse" is unknown. Try:
    No. of names: 8
      "curse"  "course"  "secateurs"  "cerise"
      "surcease"  "secure"  "cease"  "Ceausescu"

    * "magisian" is unknown. Try:
    No. of names: 7
      "magician"  "magnesia"  "Malaysian"  "mismanage"
      "imagining"  "mastication"  "fumigation"

    * "ellucidate" is unknown. Try:
    No. of names: 7
      "elucidate"  "elucidation"  "hallucinate"  "pellucid"
      "fluoridate"  "elasticated"  "illustrated"

Nothing is reported for ``horseback`` because that's correctly spelled, and :py:meth:`~.Write.spell_word` returns the boolean true.

The SpellChecker_ service and its important interfaces are shown in :numref:`ch10fig_spellcheck_serv_interface`.

..
    figure 8

.. cssclass:: diagram invert

    .. _ch10fig_spellcheck_serv_interface:
    .. figure:: https://user-images.githubusercontent.com/4193389/186258999-3a05d7ff-87fc-49d4-a662-8a5d43fe6f66.png
        :alt: Diagram of The Spell Checker Service and Interfaces.
        :figclass: align-center

        :The SpellChecker_ Service and Interfaces.

:py:meth:`.Write.spell_word` utilizes ``XSpellChecker.spell()`` to find a spelling mistake, then prints the alternative spellings:

.. tabs::

    .. code-tab:: python

        # in the Write class
        @staticmethod
        def spell_word(word: str, speller: XSpellChecker, loc: Locale | None = None) -> bool:
            if loc is None:
                loc = Locale("en", "US", "")
            alts = speller.spell(word, loc, ())
            if alts is not None:
                print(f"* '{word}' is unknown. Try:")
                alt_words = alts.getAlternatives()
                mLo.Lo.print_names(alt_words)
                return False
            return True

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

``XSpellChecker.spell()`` requires a tuple and an array of properties, which is left empty.
The properties are those associated with XLinguProperties_, which were listed above using :py:meth:`.Write.get_lingu_properties`.
Its output shows that ``IsSpellCapitalization`` is presently ``True``, which means that words in all-caps will be checked.
The property can be changed to false inside the ``PropertyValue`` tuple passed to ``XSpellChecker.spell()``. For example:

.. tabs::

    .. code-tab:: python

        props = Props.make_props(IsSpellCapitalization=False)
        alts = speller.spell(word, loc, props);

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Now an incorrectly spelled word in all-caps, such as ``CEURSE`` will be skipped over.
This means that ``Write.spellWord("CEURSE", speller)`` should return ``True``.

Unfortunately, ``XSpellChecker.spell()`` seems to ignore the property array, and still reports ``CEURSE`` as incorrectly spelled.

Even a property change performed through the XLinguProperties_ interface, such as:

.. tabs::

    .. code-tab:: python

        lingu_props = Write.get_lingu_properties()
        Props.set_property(lingu_props, "IsSpellCapitalization", False)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

fails to change ``XSpellChecker.spell()``'s behavior.
The only way to make a change to the linguistic properties that is acted upon is through the "Options" pane in the "Writing Aids" dialog, as in :numref:`ch10fig_change_cap_ss`.

..
    figure 9

.. cssclass:: screen_shot invert

    .. _ch10fig_change_cap_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/186261366-3e73934b-f9f2-48bd-a827-67a39a299864.png
        :alt: Screen shot of Changing the Capitalization Property
        :figclass: align-center

        :Changing the Capitalization Property.

Office's default spell checker is Hunspell (from https://hunspell.github.io/), and has been part of OpenOffice since v.2, when it replaced
``MySpell``, adding several features such as support for Unicode. The ``MySpell`` name lives on in a few places, such as in the spelling service (``org.openoffice.lingu.MySpellSpellChecker``).

Hunspell offers extra properties in addition to those in the "Options" pane of the "Writing Aids" dialog.
They can be accessed through the Tools, Options, Language Settings, "English sentence checking" dialog shown in :numref:`ch10fig_eng_sentence_dialog_ss`.

..
    figure 10

.. cssclass:: screen_shot invert

    .. _ch10fig_eng_sentence_dialog_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/186285751-c995b0ed-6a96-4fe0-9f96-471f4f7325ae.png
        :alt: Screen shot of The English Sentence Checking Dialog
        :figclass: align-center

        :The English Sentence Checking Dialog.

The same dialog can also be reached through the Extension Manager window shown back in :numref:`ch10fig_eng_opt_btn_ss`.
Click on the "English Spelling dictionaries" extension, and then press the "Options" button which appears as in Figure 11.

..
    figure 11

.. cssclass:: screen_shot

    .. _ch10fig_eng_opt_btn_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/186286211-37b8fa64-d7dc-477c-add4-2a07a9e7758b.png
        :alt: Screen shot of The English Spelling Options Button
        :figclass: align-center

        :The English Spelling Options Button.

Unfortunately, there appears to be no API for accessing these Hunspell options.
The best that can be done is to use a dispatch message to open the "English Sentence Checking" dialog in :numref:`ch10fig_eng_sentence_dialog_ss`.
This done by calling :py:meth:`.Write.open_sent_check_options`:

.. tabs::

    .. code-tab:: python

        GUI.set_visible(True, doc) # Office must be visible...
        Lo.delay(2000)
        Write.open_sent_check_options() # for the dialog to appear

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Write.open_sent_check_options` uses an ``.uno:OptionsTreeDialog`` dispatch along with an URL argument for the dialog's XML definition file:

.. tabs::

    .. code-tab:: python

        @staticmethod
        def open_sent_check_options() -> None:
            pip = Info.get_pip()
            lang_ext = pip.getPackageLocation("org.openoffice.en.hunspell.dictionaries")
            Lo.print(f"Lang Ext: {lang_ext}")
            url = f"{lang_ext}/dialog/en.xdl"
            props = Props.make_props(OptionsPageURL=url)
            Lo.dispatch_cmd(cmd="OptionsTreeDialog", props=props)
            Lo.wait(2000)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The XML file's location is obtained in two steps.
First the ID of the Hunspell service (``org.openoffice.en.hunspell.dictionaries``) is passed to ``XPackageInformationProvider.getPackageLocation()``
to obtain the spell checker's installation folder.
:numref:`ch10fig_hunspell_instal_dir_ss` shows a hunspell install directory.

..
    figure 12

.. cssclass:: screen_shot invert

    .. _ch10fig_hunspell_instal_dir_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/186286838-8c6deeb3-dfb2-4314-9ab8-74b584d5770a.png
        :alt: Screen shot of The English Spelling Options Button
        :figclass: align-center

        :The Hunspell Installation Folder.

The directory contains a dialog sub-directory, which holds an ``XXX.xdl`` file that defines the dialog's XML structure and data.
The ``XXX`` name will be Office's locale language, which in this case is "en".

The URL required by the ``OptionsTreeDialog`` dispatch is constructed by appending ``/dialog/en.xdl`` to the installation folder string.

.. _ch10_use_thesaurus:

10.3 Using the Thesaurus
========================

Lingo_ contains two examples of how to use the thesaurus:

.. tabs::

    .. code-tab:: python

        # in lingo exmaple
        lingo_mgr = Lo.create_instance_mcf(
            XLinguServiceManager2,
            "com.sun.star.linguistic2.LinguServiceManager"
            )
        if lingo_mgr is None:
            print("No linguistics manager found")
            return 0

        # load thesaurus
        thesaurus = lingo_mgr.getThesaurus()
        Write.spell_word("magisian", speller)
        Write.spell_word("ellucidate", speller)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The output from the first call to :py:meth:`.Write.print_meaning` is:

.. code-block:: text

    "magician" found in thesaurus; number of meanings: 2
    1.  Meaning: (noun) prestidigitator

      No. of synonyms: 6
        prestidigitator
        conjurer
        conjuror
        illusionist
        performer (generic term)
        performing artist (generic term)

    2.  Meaning: (noun) sorcerer

      No. of synonyms: 6
        sorcerer
        wizard
        necromancer
        thaumaturge
        thaumaturgist
        occultist (generic term)

``XLinguServiceManager2.getThesaurus()`` returns an instance of XThesaurus_ whose service and main interfaces are shown in :numref:`ch10fig_thesaurus_serv_interface`.

..
    figure 13

.. cssclass:: diagram invert

    .. _ch10fig_thesaurus_serv_interface:
    .. figure:: https://user-images.githubusercontent.com/4193389/186267659-aca316ae-f069-4a4a-8d52-b94b2f805027.png
        :alt: Diagram of The Thesaurus Service and Interfaces.
        :figclass: align-center

        :The Thesaurus_ Service and Interfaces.

:py:meth:`.Write.print_meaning` calls ``XThesaurus.queryMeanings()``, and prints the array of results:

.. tabs::

    .. code-tab:: python

        @staticmethod
        def print_meaning(word: str, thesaurus: XThesaurus, loc: Locale | None = None) -> int:
            if loc is None:
                loc = Locale("en", "US", "")
            meanings = thesaurus.queryMeanings(word, loc, tuple())
            if meanings is None:
                print(f"'{word}' NOT found int thesaurus")
                print()
                return 0
            m_len = len(meanings)
            print(f"'{word}' found in thesaurus; number of meanings: {m_len}")

            for i, meaning in enumerate(meanings):
                print(f"{i+1}. Meaning: {meaning.getMeaning()}")
                synonyms = meaning.querySynonyms()
                print(f" No. of  synonyms: {len(synonyms)}")
                for synonym in synonyms:
                    print(f"    {synonym}")
                print()
            return m_len

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

In a similar way to ``XSpellChecker.spell()``, ``XThesaurus.queryMeanings()`` requires a locale and an optional tuple of properties.
:py:meth:`~.Write.print_meaning` utilizes a default of  **American English**, and no properties.

If you need a non-English thesaurus which isn't part of Office, then look through the dictionary extensions at https://extensions.libreoffice.org/?Tags%5B%5D=50;
many include a thesaurus with the dictionary.

The files are built from WordNet data (https://wordnet.princeton.edu/), but use a text-based format explained very briefly in
Daniel Naber's slides about the ``Lingucomponent`` Project (at http://danielnaber.de/publications/ooocon2005-lingucomponent.pdf).
Also, the ``Lingucomponent`` website has some C++ code for reading ``.idx`` and ``.dat`` files (in https://openoffice.org/lingucomponent/MyThes-1.zip).

However, if you want to write code using a thesaurus independently of Office,
then consider programming with one of the many APIs for WordNet; listed at https://wordnet.princeton.edu/related-projects#Python.

.. _ch10_grammar_check:

10.4 Grammar Checking
=====================

Office's default grammar checker (or proof reader) is **Lightproof**, a Python application developed by :spelling:word:`László` :spelling:word:`Németh`.
``Lightproof.py``, and its support files, are installed in the same folder as the spell checker and thesaurus; on my machine that's ``\share\extensions\dict-en``.

Older versions of **Lightproof** are available from OpenOffice's extensions website at https://extensions.services.openoffice.org/project/lightproof.
One reason for downloading the old version is that it contains documentation on adding new grammar rules missing from the version installed in Office.

Another way to modify **Lightproof's** grammar rules is with its editor, which can be obtained from https://extensions.libreoffice.org/extension-center/lightproof-editor.

There are a number of alternative grammar checkers for Office, such as LanguageTool_ which are easily added to Office as extensions via the "Language Tools"

When these examples were first coded, by default the default Lightproof for grammar checking, but it doesn't have a very extensive set of built-in
grammar rules (it seems best at catching punctuation mistakes).
A switch to LanguageTool_ was made because of its larger set of rules, and its support for many languages.
It also can be used as a standalone Java library, separate from Office, and that its site includes lots of documentation.
Perhaps its biggest drawback is that it requires Java 8 or later.

Another issue is that LanguageTool and Lightproof cannot happily coexist inside Office.
**Lightproof** must be disabled and **LanguageTool** enabled via the Options, Language Settings, Writing aids, "Available language modules" pane at the top of :numref:`ch10fig_language_tool_on_ss`.

..
    figure 14

.. cssclass:: screen_shot invert

    .. _ch10fig_language_tool_on_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/186289065-dcf825b2-caac-4b90-a1e1-954e116c6a9d.png
        :alt: Screen shot of Goodbye Lightproof, hello LanguageTool
        :figclass: align-center

        :Goodbye Lightproof, hello LanguageTool

:py:meth:`.Write.print_services_info` was used earlier to list the available and configured services.

.. code-block:: text

    Available Services:
        :
    Proofreader (2):
      org.languagetool.openoffice.Main
      org.libreoffice.comp.pyuno.Lightproof.en

    Configured Services:
        :
    Proofreader (1):
      org.languagetool.openoffice.Main

``org.languagetool.openoffice.Main`` refers to the LanguageTool extension, while ``org.libreoffice.comp.pyuno.Lightproof.en`` is the English version of Lightproof.

This information can be used to set the proof reader. LanguageTool is made the default by calling :py:meth:`.Write.set_configured_services` like so:

.. tabs::

    .. code-tab:: python

        from ooodev.utils.lo import Lo
        from com.sun.star.linguistic2 import XLinguServiceManager2

        with Lo.Loader(Lo.ConnectSocket(headless=True)) as loader:
            lingo_mgr = Lo.create_instance_mcf(XLinguServiceManager2, "com.sun.star.linguistic2.LinguServiceManager")
            Write.set_configured_services(lingo_mgr, "Proofreader", "org.languagetool.openoffice.Main")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Alternatively, Lightproof can be enabled with:

.. tabs::

    .. code-tab:: python

        from ooodev.utils.lo import Lo
        from com.sun.star.linguistic2 import XLinguServiceManager2

        with Lo.Loader(Lo.ConnectSocket(headless=True)) as loader:
            lingo_mgr = Lo.create_instance_mcf(XLinguServiceManager2, "com.sun.star.linguistic2.LinguServiceManager")
            Write.set_configured_services(lingo_mgr, "Proofreader", "org.libreoffice.comp.pyuno.Lightproof.en")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The code for :py:meth:`.Write.set_configured_services` is:

.. tabs::

    .. code-tab:: python

        @staticmethod
        def set_configured_services(
            lingo_mgr: XLinguServiceManager2, service: str, impl_name: str, loc: Locale | None = None
        ) -> bool:
            cargs = CancelEventArgs(Write.set_configured_services.__qualname__)
            cargs.event_data = {
                "lingo_mgr": lingo_mgr,
                "service": service,
                "impl_name": impl_name,
            }
            _Events().trigger(WriteNamedEvent.CONFIGURED_SERVICES_SETTING, cargs)
            if cargs.cancel:
                return False
            service = cargs.event_data["service"]
            impl_name = cargs.event_data["impl_name"]
            if loc is None:
                loc = Locale("en", "US", "")
            impl_names = (impl_name,)
            lingo_mgr.setConfiguredServices(f"com.sun.star.linguistic2.{service}", loc, impl_names)
            _Events().trigger(WriteNamedEvent.CONFIGURED_SERVICES_SET, EventArgs.from_args(cargs))
            return True

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The method utilizes ``XLinguServiceManager.setConfiguredServices()`` to attach a particular implementation service
(:abbreviation:`eg:` LanguageTool) to a specified linguistic service (:abbreviation:`eg:` the Proofreader).

.. _ch10_err_rpt:

10.4.1 Error Reporting Options
------------------------------

The kinds of errors reported by the proof reader can be adjusted through Office's GUI.

One configuration pane, used by both Lightproof and LanguageTool, is in the "English Sentence Checking" dialog shown back in :numref:`ch10fig_eng_sentence_dialog_ss`.
If you look closely, the first group of check boxes are titled "Grammar checking".

If you install LanguageTool, Office's Tools menu will be modified to contain a new "LanguageTool" sub-menu shown in :numref:`ch10fig_language_tool_sub_menu_ss`.

..
    figure 15

.. cssclass:: screen_shot invert

    .. _ch10fig_language_tool_sub_menu_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/186451641-3559589a-5433-4639-8934-f2588a954df5.png
        :alt: Screen shot of The LanguageTool Submenu.
        :figclass: align-center

        :The LanguageTool Sub-menu.

The "Options" menu item in the ``LanguageTool`` sub-menu brings up an extensive set of options, reflecting the greater number of grammar rules in the checker (see :numref:`ch10fig_language_opt_dialog_ss`).

..
    figure 16

.. cssclass:: screen_shot invert

    .. _ch10fig_language_opt_dialog_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/186452371-ebd994b8-2f3b-4eca-9c0d-a254bd7efef6.png
        :alt: Screen shot of The Language Tool Options Dialog.
        :figclass: align-center

        :The LanguageTool Options Dialog.

Unfortunately, there seems to be no way to modify these options through Office's Proofreader API.

.. _ch10_proof_reader:

10.4.2 Using the Proof Reader
-----------------------------

In Lingo_ the proof reader is loaded and called like so:

.. tabs::

    .. code-tab:: python

        # load & use proof reader (Lightproof or LanguageTool)
        proofreader = Write.load_proofreader()
        print("Proofing...")
        num_errs = Write.proof_sentence("i dont have one one dogs.", proofreader)
        print(f"No. of proofing errors: {num_errs}")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The output is:

.. code-block:: text

    Proofing...
    G* This sentence does not start with an uppercase letter. in: 'i'
      Suggested change: 'I'

    G* Spelling mistake in: 'dont'
      Suggested change: 'don't'

    G* Word repetition in: 'one one'
      Suggested change: 'one'

    No. of proofing errors: 3

The proof reader isn't accessed through the linguistics manager; instead a Proofreader_ service is created, and its interfaces employed.
A simplified view of the services and interfaces are shown in :numref:`ch10fig_proofreader_serv_interface`.

..
    figure 17

.. cssclass:: diagram invert

    .. _ch10fig_proofreader_serv_interface:
    .. figure:: https://user-images.githubusercontent.com/4193389/186455013-38f47842-e1b0-448a-b5ba-6b46c7abb883.png
        :alt: Diagram of The Proof reader Service and Interfaces..
        :figclass: align-center

        :The Proofreader_ Service and Interfaces.

:py:meth:`.Write.load_proofreader` creates the service:

.. tabs::

    .. code-tab:: python

        @staticmethod
        def load_proofreader() -> XProofreader:
            proof = mLo.Lo.create_instance_mcf(XProofreader, "com.sun.star.linguistic2.Proofreader", raise_err=True)
            return proof

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Write.proof_sentence` passes a sentence to ``XProofreader.doProofreading()``, prints the errors inside the :py:meth:`~.Write.print_proof_error` and returns number of errors:

.. tabs::

    .. code-tab:: python

        @classmethod
        def proof_sentence(cls, sent: str, proofreader: XProofreader, loc: Locale | None = None) -> int:
            if loc is None:
                loc = Locale("en", "US", "")
            pr_res = proofreader.doProofreading("1", sent, loc, 0, len(sent), ())
            num_errs = 0
            if pr_res is not None:
                errs = pr_res.aErrors
                if len(errs) > 0:
                    for err in errs:
                        cls.print_proof_error(sent, err)
                        num_errs += 1
            return num_errs

        @staticmethod
        def print_proof_error(string: str, err: SingleProofreadingError) -> None:
            e_end = err.nErrorStart + err.nErrorLength
            err_txt = string[err.nErrorStart : e_end]
            print(f"G* {err.aShortComment} in: '{err_txt}'")
            if len(err.aSuggestions) > 0:
                print(f"  Suggested change: '{err.aSuggestions[0]}'")
            print()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

``XProofreader.doProofreading()`` requires a locale and properties in the same way as the earlier spell checking and thesaurus methods.
It also needs two indices for the start and end of the sentence being checked, and a document ID which is used in the error results.

The results are returned as an array of SingleProofreadingError_ objects, one for each detected error.
It's worth having a look at the documentation for the SingleProofreadingError_ class (use ``lodoc SingleProofreadingError``),
since the object contains more information than is used in :py:meth:`.Write.print_proof_error`;
for example, the ID of the grammar rule that was 'broken', a full comment string, and multiple suggestions in a String array.

Grammar rule IDs are one area where the proof reader API could be improved.
The XProofreader_ interface includes methods for switching on and off rules based on their IDs,
but there's no way to find out what these IDs are except by looking at SingleProofreadingError_ objects.

.. _ch10_string_guess:

10.5 Guessing the Language used in a String
===========================================

An oft overlooked linguistics API feature is the ability to guess the language used in a string,
which is implemented by one service, LanguageGuessing_, and a single interface, XLanguageGuessing_.
The documentation for XLanguageGuessing_ includes a long list of supported languages, including Irish Gaelic, Scots Gaelic, and Manx Gaelic.

There are two examples of language guessing in Lingo_:

.. tabs::

    .. code-tab:: python

        # from lingo example
        # guess the language
        loc = Write.guess_locale("The rain in Spain stays mainly on the plain.")
        Write.print_locale(loc)

        if loc is not None:
            print("Guessed language: " + loc.Language)

        loc = Write.guess_locale("A vaincre sans p�ril, on triomphe sans gloire.")

        if loc is not None:
            print("Guessed language: " + loc.Language)


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The output is:

.. code-block:: text

    Locale lang: 'en'; country: ''; variant: ''
    Guessed language: en
    Guessed language: fr

:py:meth:`.Write.guess_locale` creates the service, its interface, and calls ``XLanguageGuessing.guessPrimaryLanguage()``:

.. tabs::

    .. code-tab:: python

        # in the Writer class
        @staticmethod
        def guess_locale(test_str: str) -> Locale | None:
            guesser = Lo.create_instance_mcf(XLanguageGuessing, "com.sun.star.linguistic2.LanguageGuessing")
            if guesser is None:
                Lo.print("No language guesser found")
                return None
            return guesser.guessPrimaryLanguage(test_str, 0, len(test_str))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

XLanguageGuessing_ actually guesses a Locale_ rather than a language, and it includes information about the language, country and a variant BCP 47 language label.

.. _ch10_spell_chk_grammar_chk:

10.6 Spell Checking and Grammar Checking a Document
===================================================

Lingo_ only spell checks individual words, and grammar checks a single sentence.

The |lingo_file|_ example shows how these features can be applied to an entire document.

One way to scan every sentence in a document is to combine XParagraphCursor_ and XSentenceCursor_,
as in the |speak_text|_ example from :ref:`ch05_txt_cursors`. An outer loop iterates over every paragraph using XParagraphCursor_,
and an inner loop splits each paragraph into sentences with the help of XSentenceCursor_.
Initially, |lingo_file|_ was coded in this way, but it was found that XSentenceCursor_ occasionally didn't divide a paragraph into the correct number of sentences;
sometimes two sentences were treated as one.

So there was a switch to a combined Office/python approach – the outer loop in |lingo_file|_ still utilizes XParagraphCursor_ to scan the paragraphs,
but the sentences in a paragraph are extracted using :py:meth:`.Write.split_paragraph_into_sentences` that splits sentences into a list of strings.

The ``check_sentences()`` function of |lingo_file|_:

.. tabs::

    .. code-tab:: python

        def check_sentences(doc: XTextDocument) -> None:
            # load spell checker, proof reader
            speller = Write.load_spell_checker()
            proofreader = Write.load_proofreader()

            para_cursor = Write.get_paragraph_cursor(doc)
            para_cursor.gotoStart(False)  # go to start test; no selection

            while 1:
                para_cursor.gotoEndOfParagraph(True)  # select all of paragraph
                curr_para_str = para_cursor.getString()

                if len(curr_para_str) > 0:
                    print(f"\n>> {curr_para_str}")

                    sentences = Write.split_paragraph_into_sentences(curr_para_str)
                    for sentence in sentences:
                        # print(f'S <{sentence}>')
                        Write.proof_sentence(sentence, proofreader)
                        Write.spell_sentence(sentence, speller)

                if para_cursor.gotoNextParagraph(False) is False:
                    break

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`Write.load_spell_checker` does not use LinguServiceManager_ manager to create SpellChecker_.
For a yet unknown reason when speller comes from ``lingo_mgr.getSpellChecker()`` it errors when passed to methods such as :py:meth:`.Write.spell_word`.
For this reason ``com.sun.star.linguistic2.SpellChecker`` is used to get a instance of XSpellChecker_,

.. tabs::

    .. code-tab:: python

        # in the Write class
        @staticmethod
        def load_spell_checker() -> XSpellChecker:
            # lingo_mgr = Lo.create_instance_mcf(
            #     XLinguServiceManager, "com.sun.star.linguistic2.LinguServiceManager", raise_err=True
            # )
            # return lingo_mgr.getSpellChecker()
            speller = Lo.create_instance_mcf(
                XSpellChecker,
                "com.sun.star.linguistic2.SpellChecker",
                raise_err=True
                )
            return speller

        @classmethod
        def spell_sentence(cls, sent: str, speller: XSpellChecker, loc: Locale | None = None) -> int:
            words = re.split("\W+", sent)
            count = 0
            for word in words:
                is_correct = cls.spell_word(word=word, speller=speller, loc=loc)
                count = count + (0 if is_correct else 1)
            return count

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The poorly written ``badGrammar.odt`` is shown in :numref:`ch10fig_poor_writing_ss`.

..
    figure 18

.. cssclass:: screen_shot invert

    .. _ch10fig_poor_writing_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/186493075-c061f4f7-4599-45ca-8d16-83ff3a171f0d.png
        :alt: Screen shot of poor writing
        :figclass: align-center

        :Poor writing.

The output from |lingo_file|_ when given ``badGrammar.odt``:

.. code-block:: text

    >> I have a dogs. I have one dogs.

    G* Possible agreement error in: "a dogs"
       Suggested change: "a dog"


    >> I allow of of go home.  i dogs. John don’t like dogs. So recieve
    no cats also.

    G* Word repetition in: "of of"
       Suggested change: "of"

    G* This sentence does not start with an uppercase letter in: "i"
       Suggested change: "I"

    G* Grammatical problem in: "dogs"
       Suggested change: "dog"

    G* 'Also' at the end of sentence in: "also"
       Suggested change: "as well"

    * "recieve" is unknown. Try:
    No. of names: 8
      "receive"  "relieve"  "retrieve"  "reprieve"
      "reverie"  "recitative"  "Recife"  "reserve"

    The grammar errors (those starting with "G*") are produced  by the LanguageTool
    proof checker. If the default Lightproof checker is utilized instead, then less errors are
    found:

    >> I have a dogs. I have one dogs.


    >> I allow of of go home.  i dogs. John don’t like dogs. So recieve
    no cats also.

    G* Word duplication? in: "of of"
       Suggested change: "of"

    G* Missing capitalization? in: "i"
       Suggested change: "I"

    * "recieve" is unknown. Try:
    No. of names: 8
      "receive"  "relieve"  "retrieve"  "reprieve"
      "reverie"  "recitative"  "Recife"  "reserve"

On larger documents, it's a good idea to redirect the voluminous output to a temporary file so it can be examined easily.

The output can be considerably reduced if LanguageTool's unpaired rule is disabled, via the Options dialog in :numref:`ch10fig_language_opt_dialog_ss`.
:numref:`ch10fig_lang_tool_inparied_desel_ss` shows the dialog with the "Unpaired" checkbox deselected in the Punctuation section.

..
    figure 19

.. cssclass:: screen_shot invert

    .. _ch10fig_lang_tool_inparied_desel_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/186496075-81cbf885-8c78-46a1-94f9-7b7313ca2589.png
        :alt: Screen shot ofThe Language Tool Options Dialog with the Unpaired Rule Deselected.
        :figclass: align-center

        :The LanguageTool Options Dialog with the Unpaired Rule Deselected.

The majority of the remaining errors are words unknown to the spell checker, such as names and places, and British English spellings.

Most of the grammar errors relate to how direct speech is written.
The grammar checker mistakenly reports an error if the direct speech ends with a question mark or exclamation mark without a comma after the quoted text.

.. |lingo_file| replace:: Lingo File
.. _lingo_file: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_lingo_file

.. |lingustic_ex| replace:: LinguisticExample.java
.. _lingustic_ex: https://api.libreoffice.org/examples/DevelopersGuide/OfficeDev/Linguistic/LinguisticExamples.java

.. _Lingo: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_lingo
.. _LanguageTool: https://extensions.libreoffice.org/en/extensions/show/languagetool

.. |speak_text| replace:: Speak Text
.. _speak_text: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_speak

.. _ConversionDictionary: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1linguistic2_1_1ConversionDictionary.html
.. _ConversionDictionaryList: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1linguistic2_1_1ConversionDictionaryList.html
.. _Dictionary: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1linguistic2_1_1Dictionary.html
.. _DictionaryList: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1linguistic2_1_1DictionaryList.html
.. _LanguageGuessing: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1linguistic2_1_1LanguageGuessing.html
.. _LinguServiceManager: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1linguistic2_1_1LinguServiceManager.html
.. _Locale: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1lang_1_1Locale.html
.. _Proofreader: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1linguistic2_1_1Proofreader.html
.. _Proofreader: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1linguistic2_1_1Proofreader.html
.. _SingleProofreadingError: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1linguistic2_1_1SingleProofreadingError.html
.. _SpellChecker: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1linguistic2_1_1SpellChecker.html
.. _Thesaurus: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1linguistic2_1_1Thesaurus.html
.. _XConversionDictionary: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1linguistic2_1_1XConversionDictionary.html
.. _XConversionPropertyType: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1linguistic2_1_1XConversionPropertyType.html
.. _XDictionary: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1linguistic2_1_1XDictionary.html
.. _XLanguageGuessing: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1linguistic2_1_1XLanguageGuessing.html
.. _XLinguProperties: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1linguistic2_1_1XLinguProperties.html
.. _XNameContainer: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XNameContainer.html
.. _XPackageInformationProvider: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1deployment_1_1XPackageInformationProvider.html
.. _XParagraphCursor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XParagraphCursor.html
.. _XProofreader: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1linguistic2_1_1XProofreader.html
.. _XSentenceCursor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XSentenceCursor.html
.. _XSpellChecker: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1linguistic2_1_1XSpellChecker.html
.. _XThesaurus: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1linguistic2_1_1XThesaurus.html
