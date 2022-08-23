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

.. cssclass:: screen_shot invert

    .. _ch10fig_lang_dial_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/186042384-13763979-562a-497b-9d86-54e1e7200b89.png
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

It retrieves a conventional dictionary list first (called ``dict_lst``), and iterates through its dictionaries using :py:meth:`~.Write.print_con_dicts_info`.
Then it obtains the conversion dictionary list (called ``cd_list``), and iterates over that with :py:meth:`~.Write.print_con_dicts_info`.

:numref:`ch09fig_dicts_services` shows the main services and interfaces used by ordinary dictionaries.

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
            for d in dicts:
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

Conversion dictionaries map words in one language/dialect to corresponding words in another language/dialect.
:numref:`ch09fig_convert_dicts_services` shows that conversion dictionaries are organized in a similar way to ordinary ones.
The interfaces for manipulating a conversion dictionary are XConversionDictionary_ and XConversionPropertyType_.

.. cssclass:: diagram invert

    .. _ch09fig_convert_dicts_services:
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

Output similar to :py:meth:`.Write.dicts_info` can be viewed via Office's Tools, Options, Language Settings, "Writing Aids" dialog, shown in :numref:`ch10fig_writing_aids_ss`.

.. cssclass:: screen_shot invert

    .. _ch10fig_writing_aids_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/186045267-11db8569-cbb2-4b6d-bb3f-abe69c7f8073.png
        :alt: Screen shot of The Writing Aids Dialog
        :figclass: align-center

        :The Writing Aids Dialog.

The dictionaries are listed in the second pane of the dialog.
Also, at the bottom of the window is a "Get more dictionaries online" hyperlink which takes the user to Office's extension website, and displays the "Dictionary" category (see :numref:`ch10fig_ext_dict_ss`).

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

10.1.2 Linguistic Properties
----------------------------

Back in the Lingo_ example, :py:meth:`.Write.get_lingu_properties` returns an instance of XLinguProperties_,
and its properties are printed by calling :py:meth:`.Props.show_props`:

.. tabs::

    .. code-tab:: python

        # code fragment from lingo example
        lingu_props = Write.get_lingu_properties()
        Props.show_props("Linguistic Manager", lingu_props)

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
      IsSpellAuto: False
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

10.1.3 Installed Extensions
---------------------------

Additional dictionaries, and other language tools such as grammar checkers, are loaded into Office as extensions, so calling :py:meth:`.Info.list_extensions` can be informative.

The output on one of my test machine is:

.. code-block:: text

    Extensions:
    1. ID: apso.python.script.organizer
       Version: 1.2.8
       Loc: file:///home/user/.config/libreoffice/4/user/uno_packages/cache/uno_packages/lu59147xjqms4.tmp_/apso-v2.oxt

    2. ID: French.linguistic.resources.from.Dicollecte.by.OlivierR
       Version: 5.7
       Loc: file:///home/user/.config/libreoffice/4/user/uno_packages/cache/uno_packages/lu287421qavj.tmp_/lo-oo-ressources-linguistiques-fr-v5-7.oxt

    3. ID: org.openoffice.languagetool.oxt
       Version: 5.8
       Loc: file:///home/user/.config/libreoffice/4/user/uno_packages/cache/uno_packages/lu14553844wbl51.tmp_/LanguageTool-stable.oxt

    4. ID: mytools.mri
       Version: 1.3.3
       Loc: file:///home/user/.config/libreoffice/4/user/uno_packages/cache/uno_packages/lu1050215332vj9.tmp_/MRI-1.3.3.oxt

    5. ID: spanish.es.dicts.from.rla-es
       Version: 2.6
       Loc: file:///home/user/.config/libreoffice/4/user/uno_packages/cache/uno_packages/lu305561ujwy.tmp_/es.oxt

    6. ID: org.openoffice.legacy.nlpsolver
       Version:
       Loc: file:///usr/lib/libreoffice/share/extensions/nlpsolver

    7. ID: org.openoffice.legacy.wiki-publisher
       Version:
       Loc: file:///usr/lib/libreoffice/share/extensions/wiki-publisher

The ``Loc`` entries are the directories or OXT files containing the extensions. Most extensions are placed in the share extensions folder on Windows.

Office can display similar information via its Tools, "Extension Manager" dialog, as in :numref:`ch10fig_ext_dial_ss`.

.. cssclass:: screen_shot invert

    .. _ch10fig_ext_dial_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/186047591-8a85daa8-e637-410b-b37f-a5151908961c.png
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

Extensions are accessed via the XPackageInformationProvider_ interface.

Work in progress ...

.. |lingo_file| replace:: Lingo File
.. _lingo_file: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_lingo_file

.. |lingustic_ex| replace:: LinguisticExample.java
.. _lingustic_ex: https://api.libreoffice.org/examples/DevelopersGuide/OfficeDev/Linguistic/LinguisticExamples.java

.. _Lingo: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_lingo

.. _ConversionDictionary: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1linguistic2_1_1ConversionDictionary.html
.. _ConversionDictionaryList: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1linguistic2_1_1ConversionDictionaryList.html
.. _Dictionary: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1linguistic2_1_1Dictionary.html
.. _DictionaryList: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1linguistic2_1_1DictionaryList.html
.. _XConversionDictionary: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1linguistic2_1_1XConversionDictionary.html
.. _XConversionPropertyType: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1linguistic2_1_1XConversionPropertyType.html
.. _XDictionary: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1linguistic2_1_1XDictionary.html
.. _XLinguProperties: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1linguistic2_1_1XLinguProperties.html
.. _XNameContainer: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XNameContainer.html
.. _XPackageInformationProvider: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1deployment_1_1XPackageInformationProvider.html
