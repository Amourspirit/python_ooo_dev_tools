# coding: utf-8
# region Imports
from __future__ import annotations
from typing import TYPE_CHECKING, Iterable, List, cast, overload

import uno


from com.sun.star.awt import XControl
from com.sun.star.awt import XControlModel
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XChild
from com.sun.star.container import XIndexContainer
from com.sun.star.container import XNameAccess
from com.sun.star.container import XNameAccess
from com.sun.star.container import XNameContainer
from com.sun.star.container import XNamed
from com.sun.star.drawing import XControlShape
from com.sun.star.drawing import XDrawPage
from com.sun.star.drawing import XDrawPagesSupplier
from com.sun.star.drawing import XDrawPageSupplier
from com.sun.star.drawing import XShapes
from com.sun.star.form import XForm
from com.sun.star.form import XFormsSupplier
from com.sun.star.form import XGridColumnFactory
from com.sun.star.lang import XComponent
from com.sun.star.lang import XServiceInfo
from com.sun.star.script import XEventAttacherManager

from ooo.dyn.awt.point import Point
from ooo.dyn.awt.size import Size as UnoSize
from ooo.dyn.form.form_component_type import FormComponentType
from ooo.dyn.form.list_source_type import ListSourceType
from ooo.dyn.script.script_event_descriptor import ScriptEventDescriptor
from ooo.dyn.sdb.command_type import CommandType
from ooo.dyn.text.text_content_anchor_type import TextContentAnchorType


from ooodev.utils import lo as mLo
from ooodev.utils import info as mInfo
from ooodev.utils import props as mProps
from ooodev.utils import gui as mGui
from ooodev.utils.kind.form_component_kind import FormComponentKind
from ooodev.proto.style_obj import StyleT
from .controls.form_ctl_button import FormCtlButton
from .controls.form_ctl_check_box import FormCtlCheckBox

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from ooodev.units import UnitT
# endregion Imports


class Forms:
    # region    access forms in document
    # region        get_forms()
    @overload
    @classmethod
    def get_forms(cls, obj: XComponent) -> XNameContainer:
        ...

    @overload
    @classmethod
    def get_forms(cls, obj: XDrawPage) -> XNameContainer:
        ...

    @classmethod
    def get_forms(cls, obj: XComponent | XDrawPage) -> XNameContainer:
        """
        Gets Forms

        Args:
            obj (XComponent | XDrawPage): component or draw page

        Returns:
            XNameContainer: name container
        """
        if mLo.Lo.is_uno_interfaces(obj, XDrawPage):
            draw_page = obj
        else:
            draw_page = cls.get_draw_page(cast(XComponent, obj))

        forms_supp = mLo.Lo.qi(XFormsSupplier, draw_page, True)

        return forms_supp.getForms()

    # endregion     get_forms()

    @staticmethod
    def get_draw_page(doc: XComponent) -> XDrawPage:
        """
        Gets draw page

        Args:
            doc (XComponent): Component

        Raises:
            Exception: If unable to get draw page

        Returns:
            XDrawPage: Draw Page
        """
        # sourcery skip: raise-specific-error
        try:
            supp_page = mLo.Lo.qi(XDrawPageSupplier, doc)
            if supp_page is not None:
                return supp_page.getDrawPage()

            # doc supports multiple DrawPages
            supp_pages = mLo.Lo.qi(XDrawPagesSupplier, doc, True)

            pages = supp_pages.getDrawPages()
            return mLo.Lo.qi(XDrawPage, pages.getByIndex(0), True)
        except Exception as e:
            raise Exception(f"Unable to get draw page: {e}") from e

    # region        get_form()
    @overload
    @classmethod
    def get_form(cls, obj: XComponent) -> XNameContainer:
        ...

    @overload
    @classmethod
    def get_form(cls, obj: XComponent, form_name: str) -> XForm:
        ...

    @overload
    @classmethod
    def get_form(cls, obj: XDrawPage) -> XNameContainer:
        ...

    @classmethod
    def get_form(cls, obj: XComponent | XDrawPage, form_name: str = "") -> XNameContainer | XForm:
        """
        Gets form as name container

        Args:
            obj (XComponent | XDrawPage): Component or draw page
            form_name (str, optional): the name of form to get.

        Raises:
            Exception: If unable to get form

        Returns:
            XNameContainer: Name container
        """
        # sourcery skip: raise-specific-error
        if form_name:
            # get_form(cls, obj: XComponent, for_name: str)
            try:
                named_forms = cls.get_forms(obj)
                con = cls.get_form_by_name(form_name, named_forms)
                return mLo.Lo.qi(XForm, con, True)
            except Exception as e:
                raise Exception(f"Unable to get form: {e}") from e
        try:
            if mLo.Lo.is_uno_interfaces(obj, XDrawPage):
                draw_page = obj
            else:
                draw_page = cls.get_draw_page(cast(XComponent, obj))

            idx_forms = cls.get_indexed_forms(cast(XDrawPage, draw_page))
        except Exception as e:
            raise Exception(f"Unable to get form: {e}") from e

        try:
            return mLo.Lo.qi(XNameContainer, idx_forms.getByIndex(0), True)
        except Exception as e:
            raise Exception(f"Could not find default form: {e}") from e

    # endregion     get_form()

    @staticmethod
    def get_form_by_name(form_name: str, named_forms: XNameContainer) -> XNameContainer:
        """
        Get a form by name

        Args:
            form_name (str): form name
            named_forms (XNameContainer): name container

        Raises:
            Exception: If not able to find form

        Returns:
            XNameContainer: Name Container
        """
        # sourcery skip: raise-specific-error
        try:
            return mLo.Lo.qi(XNameContainer, named_forms.getByName(form_name), True)
        except Exception as e:
            raise Exception(f'Could not find the form "{form_name}"') from e

    @staticmethod
    def get_indexed_forms(draw_page: XDrawPage) -> XIndexContainer:
        """
        Get index forms

        Args:
            draw_page (XDrawPage): Draw page

        Returns:
            XIndexContainer: Index container
        """
        form_supp = mLo.Lo.qi(XFormsSupplier, draw_page, True)
        return mLo.Lo.qi(XIndexContainer, form_supp.getForms(), True)

    # region        insert_form()
    @classmethod
    def _insert_form_name_comp(cls, doc: XComponent) -> XNameContainer:
        doc_forms = cls.get_forms(doc)
        return cls._insert_form_name_container("GridForm", doc_forms)

    @classmethod
    def _insert_form_name_container(cls, form_name: str, named_forms: XNameContainer) -> XNameContainer:
        # sourcery skip: raise-specific-error
        if named_forms.hasByName(form_name):
            mLo.Lo.print(f'"{form_name}" already exists')
            return cls.get_form_by_name(form_name=form_name, named_forms=named_forms)

        try:
            xnamed_forms = mLo.Lo.create_instance_msf(
                XNameContainer, "com.sun.star.form.component.DataForm", raise_err=True
            )
            named_forms.insertByName(form_name, xnamed_forms)
            return xnamed_forms
        except Exception as e:
            raise Exception(f'Could not insert the form "{form_name}": {e}') from e

    @overload
    @classmethod
    def insert_form(cls, doc: XComponent) -> XNameContainer:
        """
        Insert form

        Args:
            form_name (str): Form name
            doc (XComponent): Component

        Returns:
            XNameContainer: Name Container
        """
        ...

    @overload
    @classmethod
    def insert_form(cls, form_name: str, named_forms: XNameContainer) -> XNameContainer:
        """
        Insert form

        Args:
            form_name (str): Form name
            named_forms (XNameContainer): Name Container

        Returns:
            XNameContainer: Name Container
        """
        ...

    @classmethod
    def insert_form(cls, *args, **kwargs) -> XNameContainer:
        """
        Insert form

        Args:
            form_name (str): Form name
            doc (XComponent): Component
            named_forms (XNameContainer): Name Container

        Returns:
            XNameContainer: Name Container
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("form_name", "doc", "named_forms")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("insert_form() got an unexpected keyword argument")
            keys = ("doc", "form_name")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            ka[2] = kwargs.get("named_forms", None)
            return ka

        if count not in (1, 2):
            raise TypeError("insert_form() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            return cls._insert_form_name_comp(kargs[1])

        return cls._insert_form_name_container(form_name=kargs[1], named_forms=kargs[2])
        # endregion     insert_form()

    @classmethod
    def has_form(cls, doc: XComponent, form_name: str) -> bool:
        """
        Gets if component has form by name

        Args:
            doc (XComponent): Component
            form_name (str): Form name

        Returns:
            bool: ``True`` if has form, Otherwise ``False``
        """
        try:
            draw_page = cls.get_draw_page(doc)
        except Exception:
            return False
        try:
            named_forms_container = cls.get_forms(draw_page)
            if named_forms_container is None:
                mLo.Lo.print("No forms found on page")
                return False
            xnamed_forms = mLo.Lo.qi(XNameAccess, named_forms_container, True)
            return xnamed_forms.hasByName(form_name)
        except Exception:
            return False

    @classmethod
    def show_form_names(cls, doc: XComponent) -> None:
        """
        Prints form names to console.

        Args:
            doc (XComponent): Component

        Returns:
            None:
        """
        form_names_con = cls.get_forms(doc)
        form_names = form_names_con.getElementNames()
        print(f"No. of forms found: {len(form_names)}")
        for name in form_names:
            print(f"  {name}")
        print()

    @classmethod
    def list_forms(cls, obj: XComponent | XNameAccess, tab_str: str = "  ") -> None:
        """
        Prints forms information to console.

        Args:
            obj (XComponent | XNameAccess): Component or Name Access
            tab_str (str, optional): tab string

        Returns:
            None:
        """
        if mLo.Lo.is_uno_interfaces(obj, XComponent):
            container = cls.get_forms(cast(XComponent, obj))
        else:
            container = cast(XNameContainer, obj)
        nms = container.getElementNames()
        for name in nms:
            try:
                service_info = mLo.Lo.qi(XServiceInfo, container.getByName(name), True)
                if service_info.supportsService("com.sun.star.form.FormComponents"):
                    # this means that the form has been found
                    if mInfo.Info.support_service(service_info, "com.sun.star.form.component.DataForm"):
                        print(f'{tab_str}Data From "{name}"')
                    else:
                        print(f'{tab_str}Form "{name}"')
                        # mInfo.Info.show_services("Form", service_info)
                        # mInfo.Info.show_interfaces("Form", service_info)
                    child_con = mLo.Lo.qi(XNameAccess, service_info, True)
                    # recursively list form components
                    cls.list_forms(child_con, tab_str=f"{tab_str}  ")
                elif service_info.supportsService("com.sun.star.form.FormComponent"):
                    model = mLo.Lo.qi(XControlModel, service_info, True)
                    print(f'{tab_str}"{name}":{cls.get_type_str(model)}')
                    #  mProps.Props.show_obj_props("Model", model)
                else:
                    print(f'{tab_str}unknown: "{name}"')
            except Exception:
                print(f'{tab_str}Could not access "{name}"')

    # endregion access forms in document

    # region    get form models
    @overload
    @classmethod
    def get_models(cls, obj: XComponent) -> List[XControlModel]:
        ...

    @overload
    @classmethod
    def get_models(cls, obj: XNameAccess) -> List[XControlModel]:
        ...

    @classmethod
    def get_models(cls, obj: XComponent | XNameAccess) -> List[XControlModel]:
        """
        Gets models from obj

        Args:
            obj (XComponent | XNameAccess): Component or Name Access

        Returns:
            List[XControlModel]: List of found models

        See Also:
            :py:meth:`~.forms.Forms.get_models2`
        """
        if mLo.Lo.is_uno_interfaces(obj, XComponent):
            container = cls.get_forms(cast(XComponent, obj))
        else:
            container = cast(XNameContainer, obj)
        models: List[XControlModel] = []
        nms = container.getElementNames()
        for name in nms:
            try:
                service_info = mLo.Lo.qi(XServiceInfo, container.getByName(name), True)
                if service_info.supportsService("com.sun.star.form.FormComponents"):
                    # this means that a form has been found
                    child_con = mLo.Lo.qi(XNameAccess, service_info, True)
                    # recursively search
                    models.extend(cls.get_models(child_con))
                elif service_info.supportsService("com.sun.star.form.FormComponent"):
                    model = mLo.Lo.qi(XControlModel, service_info, True)
                    models.append(model)

            except Exception as e:
                mLo.Lo.print(f'Could not access "{name}"')
        return models

    @classmethod
    def get_models2(cls, doc: XComponent, form_name: str) -> List[XControlModel]:
        """
        Gets models from doc

        Args:
            doc (XComponent): Component
            form_name (str): form name.

        Returns:
            List[XControlModel]: List of found models

        See Also:
            :py:meth:`~.forms.Forms.get_models`
        """
        # another way to obtain models, via the control shapes in the DrawPage
        models: List[XControlModel] = []
        try:
            xdraw_page = cls.get_draw_page(doc)
            if xdraw_page is None:
                return models
        except Exception:
            mLo.Lo.print("No draw page found")
            return models
        try:
            for i in range(xdraw_page.getCount()):
                shape = mLo.Lo.qi(XControlShape, xdraw_page.getByIndex(i), True)
                model = shape.getControl()
                if cls.belongs_to_form(model, form_name):
                    models.append(model)
        except Exception as e:
            mLo.Lo.print(f"Could not collect control model: {e}")
        mLo.Lo.print(f"No. of control models found: {len(models)}")
        return models

    @classmethod
    def get_event_source_name(cls, event: EventObject) -> str:
        """
        Gets event source name

        Args:
            event (EventObject): event object

        Returns:
            str: event source name
        """
        control = mLo.Lo.qi(XControlModel, event.Source, True)
        return cls.get_name(control)

    @staticmethod
    def get_event_control_model(event: EventObject) -> XControlModel:
        """
        Gets event control model

        Args:
            event (EventObject): event object

        Returns:
            XControlModel: Event control model
        """
        control = mLo.Lo.qi(XControl, event.Source, True)
        return control.getModel()

    @staticmethod
    def get_form_name(ctl_model: XControlModel) -> str:
        """
        Gets form name

        Args:
            ctl_model (XControlModel): control model

        Returns:
            str: form name
        """
        child = mLo.Lo.qi(XChild, ctl_model, True)
        named = mLo.Lo.qi(XNamed, child.getParent(), True)
        return named.getName()

    @classmethod
    def belongs_to_form(cls, ctl_model: XControlModel, form_name: str) -> bool:
        """
        Get if a control belongs to a form

        Args:
            ctl_model (XControlModel): Control Model
            form_name (str): Form name

        Returns:
            bool: ``True`` if belongs to form; Otherwise, ``False``
        """
        return cls.get_form_name(ctl_model) == form_name

    @staticmethod
    def get_name(ctl_model: XControlModel) -> str:
        """
        Gets name of a given form component

        Args:
            ctl_model (XControlModel): Control Model

        Returns:
            str: Name of component
        """
        return str(mProps.Props.get(ctl_model, "Name"))

    @staticmethod
    def get_label(ctl_model: XControlModel) -> str:
        """
        Gets label of a given form component

        Args:
            ctl_model (XControlModel): Control Model

        Returns:
            str: Label of component
        """
        return str(mProps.Props.get(ctl_model, "Label"))

    @classmethod
    def get_type_str(cls, ctl_model: XControlModel) -> str | None:
        """
        Gets type as string

        Args:
            ctl_model (XControlModel): Control Model

        Returns:
            str | None: Type as string if found; Otherwise, ``None``
        """
        # sourcery skip: low-code-quality
        control_id = cls.get_id(ctl_model)
        if control_id == -1:
            return None

        service_info = mLo.Lo.qi(XServiceInfo, ctl_model)
        if control_id == FormComponentType.COMMANDBUTTON:
            return "Command button"
        elif control_id == FormComponentType.RADIOBUTTON:
            return "Radio button"
        elif control_id == FormComponentType.IMAGEBUTTON:
            return "Image button"
        elif control_id == FormComponentType.CHECKBOX:
            return "Check Box"
        elif control_id == FormComponentType.LISTBOX:
            return "List Box"
        elif control_id == FormComponentType.COMBOBOX:
            return "Combo Box"
        elif control_id == FormComponentType.GROUPBOX:
            return "Group Box"
        elif control_id == FormComponentType.FIXEDTEXT:
            return "Fixed Text"
        elif control_id == FormComponentType.GRIDCONTROL:
            return "Grid Control"
        elif control_id == FormComponentType.FILECONTROL:
            return "File Control"
        elif control_id == FormComponentType.HIDDENCONTROL:
            return "Hidden Control"
        elif control_id == FormComponentType.IMAGECONTROL:
            return "Image Control"
        elif control_id == FormComponentType.DATEFIELD:
            return "Date Field"
        elif control_id == FormComponentType.TIMEFIELD:
            return "Time Field"
        elif control_id == FormComponentType.NUMERICFIELD:
            return "Numeric Field"
        elif control_id == FormComponentType.CURRENCYFIELD:
            return "Currency Field"
        elif control_id == FormComponentType.PATTERNFIELD:
            return "Pattern Field"
        elif control_id == FormComponentType.TEXTFIELD:
            # two services with this class id: text field and formatted field
            if service_info is not None and service_info.supportsService("com.sun.star.form.component.FormattedField"):
                return "Formatted Field"
            else:
                return "Text Field"
        else:
            mLo.Lo.print(f"Unknown class ID: {control_id}")
            return None

    @staticmethod
    def get_id(ctl_model: XControlModel) -> int:
        """
        Gets class id for a form component

        Args:
            ctl_model (XControlModel): Control Model

        Returns:
            int: Class Id if found, Otherwise ``-1``
        """
        class_id = mProps.Props.get(ctl_model, "ClassId")
        if class_id is None:
            mLo.Lo.print("No class ID found for form component")
            return -1
        return int(class_id)

    @classmethod
    def is_button(cls, ctl_model: XControlModel) -> bool:
        """
        Gets if component is a command button or a image button

        Args:
            ctl_model (XControlModel): Control Model

        Returns:
            bool: ``True`` if is button; Otherwise, ``False``
        """
        button_id = cls.get_id(ctl_model)
        if button_id == -1:
            return False

        return button_id in (
            FormComponentType.COMMANDBUTTON,
            FormComponentType.IMAGEBUTTON,
        )

    @classmethod
    def is_text_field(cls, ctl_model: XControlModel) -> bool:
        """
        Gets if component is a text field

        Args:
            ctl_model (XControlModel): Control Model

        Returns:
            bool: ``True`` if is text field; Otherwise, ``False``
        """
        text_id = cls.get_id(ctl_model)
        if text_id == -1:
            return False

        return text_id in (
            FormComponentType.DATEFIELD,
            FormComponentType.TIMEFIELD,
            FormComponentType.NUMERICFIELD,
            FormComponentType.CURRENCYFIELD,
            FormComponentType.PATTERNFIELD,
            FormComponentType.TEXTFIELD,
        )

    @classmethod
    def is_box(cls, ctl_model: XControlModel) -> bool:
        """
        Gets if component is a box

        Args:
            ctl_model (XControlModel): Control Model

        Returns:
            bool: ``True`` if is box; Otherwise, ``False``
        """
        box_id = cls.get_id(ctl_model)
        if box_id == -1:
            return False

        return box_id in (FormComponentType.RADIOBUTTON, FormComponentType.CHECKBOX)

    @classmethod
    def is_list(cls, ctl_model: XControlModel) -> bool:
        """
        Gets if component is a list

        Args:
            ctl_model (XControlModel): Control Model

        Returns:
            bool: ``True`` if is list; Otherwise, ``False``
        """
        control_id = cls.get_id(ctl_model)
        if control_id == -1:
            return False

        return control_id in (FormComponentType.LISTBOX, FormComponentType.COMBOBOX)

    # Other control types
    # FormComponentType.GROUPBOX
    # FormComponentType.FIXEDTEXT
    # FormComponentType.GRIDCONTROL
    # FormComponentType.FILECONTROL
    # FormComponentType.HIDDENCONTROL
    # FormComponentType.IMAGECONTROL
    # FormComponentType.SCROLLBAR
    # FormComponentType.SPINBUTTON
    # FormComponentType.NAVIGATIONBAR

    # endregion get form models

    # region    get control for a model
    @staticmethod
    def get_control(doc: XComponent, ctl_model: XControlModel) -> XControl:
        """
        Gets the control from the specified control model.

        Args:
            doc (XComponent): Component
            ctl_model (XControlModel): Control Model

        Raises:
            Exception: If unable to get control

        Returns:
            XControl: Control
        """
        # sourcery skip: raise-specific-error
        try:
            control_access = mGui.GUI.get_control_access(doc)
            if control_access is None:
                raise Exception("Could not obtain controls access in document")
            return control_access.getControl(ctl_model)
        except Exception as e:
            raise Exception(f"Could not access control: {e}") from e

    @classmethod
    def get_named_control(cls, doc: XComponent, ctl_name: str) -> XControl | None:
        """
        Gets a named control.

        Args:
            doc (XComponent): Component
            ctl_name (str): Name of control

        Returns:
            XControl | None: Control if found; Otherwise, None
        """
        models = cls.get_models(doc)
        ctl = None
        for model in models:
            if cls.get_name(model) == ctl_name:
                mLo.Lo.print(f"Found: {ctl_name}")
                try:
                    ctl = cls.get_control(doc, model)
                except Exception as e:
                    mLo.Lo.print("Error getting control.")
                    mLo.Lo.print(f"  {e}")
                    ctl = None
                finally:
                    break
        return ctl

    @classmethod
    def get_control_model(cls, doc: XComponent, ctl_name) -> XControlModel | None:
        """
        Gets Control Model by Name

        Args:
            doc (XComponent): Component
            ctl_name (str): Name of control

        Returns:
            XControlModel | None: Control Model if found; Otherwise, None
        """
        control = cls.get_named_control(doc, ctl_name)
        return None if control is None else control.getModel()

    # endregion get control for a model

    # region create controls

    @staticmethod
    def create_name(elem_container: XNameAccess, name: str) -> str:
        """
        Creates a name.

        Make a unique string by appending a number to the supplied name

        Args:
            elem_container (XNameAccess): container
            name (str): current name

        Returns:
            str: a name not in container.
        """
        used_name = True
        i = 1
        nm = f"{name}{i}"
        while used_name:
            used_name = elem_container.hasByName(nm)
            if used_name:
                i += 1
                nm = f"{name}{i}"
        return nm

    @staticmethod
    def _get_unit100_value(value: int | UnitT) -> int:
        try:
            return value.get_value_mm100()  # type: ignore
        except AttributeError:
            return int(value) * 100  # type: ignore

    # region    add_control
    @classmethod
    def add_control(
        cls,
        doc: XComponent,
        *,
        label: str | None,
        comp_kind: FormComponentKind | str,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        name: str = "",
        anchor_type: TextContentAnchorType | None = None,
        parent_form: XNameContainer | None = None,
        styles: Iterable[StyleT] | None = None,
    ) -> XPropertySet:
        """
        Add a control

        Args:
            doc (XComponent): Component
            name (str): Control Name
            label (str | None): Label to assign to control
            comp_kind (FormComponentKind | str): Kind of control such as ``CheckBox``.
            x (int, UnitT): Control X position
            y (int, UnitT): Control Y Position
            width (int, UnitT): Control width#
            height (int, UnitT): control height
            anchor_type (TextContentAnchorType | None): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            parent_form (XNameContainer | None): Parent form in which to add control.
            styles (Iterable[StyleT], optional): One or more styles to apply.

        Returns:
            XPropertySet: Control Property Set

        See Also:
            For ``comp_kind`` `API component Module Namespace <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1form_1_1component.html>`_

        .. versionchanged:: 0.9.2
            Added ``styles`` argument.
        """
        try:
            # create a shape to represent the control's view
            #     xDocAsFactory = mLo.Lo.qi(XMultiServiceFactory, doc, True)
            #     shape =
            #     XMultiServiceFactory xDocAsFactory = (XMultiServiceFactory)UnoRuntime.queryInterface(
            #   XMultiServiceFactory.class, s_aDocument);

            #     XControlShape xShape = (XControlShape)UnoRuntime.queryInterface(XControlShape.class,
            #   xDocAsFactory.createInstance("com.sun.star.drawing.ControlShape"));

            shape = mLo.Lo.create_instance_msf(XControlShape, "com.sun.star.drawing.ControlShape", raise_err=True)

            width_value = cls._get_unit100_value(width)
            height_value = cls._get_unit100_value(height)
            x_value = cls._get_unit100_value(x)
            y_value = cls._get_unit100_value(y)

            # position and size of the shape
            shape.setSize(UnoSize(width_value, height_value))
            shape.setPosition(Point(x_value, y_value))

            # adjust the anchor so that the control is tied to the page
            shape_props = mLo.Lo.qi(XPropertySet, shape, True)

            if anchor_type is None:
                shape_props.setPropertyValue("AnchorType", TextContentAnchorType.AT_PARAGRAPH)
            else:
                shape_props.setPropertyValue("AnchorType", TextContentAnchorType(anchor_type))

            # create the control's model, this is a service
            # see: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1form_1_1FormControlModel.html
            model = mLo.Lo.create_instance_mcf(
                XControlModel, f"com.sun.star.form.component.{comp_kind}", raise_err=True
            )

            # insert the model into the form (or default to "Form")
            if parent_form is not None:
                parent_form.insertByName(name, model)
                if not name:
                    name = cls.create_name(parent_form, "Control")
            else:
                if not name:
                    raise ValueError("name must be specified if parent_form is None")

            # link model to the shape
            shape.setControl(model)

            # add the shape to the shapes on the doc's draw page
            draw_page = cls.get_draw_page(doc)
            form_shapes = mLo.Lo.qi(XShapes, draw_page, True)
            form_shapes.add(shape)

            # styles need to not be added until after form_shapes.add(shape) or may not work
            if styles:
                for style in styles:
                    style.apply(shape_props)

            # set Name and Label properties for the model
            model_props = mLo.Lo.qi(XPropertySet, model, True)
            model_props.setPropertyValue("Name", name)
            if label is not None:
                model_props.setPropertyValue("Label", label)

            return model_props
        except Exception:
            raise
        # endregion add_control

    # region    add_labelled_control

    @overload
    @classmethod
    def add_labelled_control(
        cls, doc: XComponent, *, label: str, comp_kind: FormComponentKind | str, y: int
    ) -> XPropertySet:
        ...

    @overload
    @classmethod
    def add_labelled_control(
        cls,
        doc: XComponent,
        *,
        label: str,
        comp_kind: FormComponentKind | str,
        y: int | UnitT,
        lbl_styles: Iterable[StyleT] = ...,
        ctl_styles: Iterable[StyleT] = ...,
    ) -> XPropertySet:
        ...

    @overload
    @classmethod
    def add_labelled_control(
        cls,
        doc: XComponent,
        *,
        label: str,
        comp_kind: FormComponentKind | str,
        x: int | UnitT,
        y: int | UnitT,
        height: int | UnitT,
    ) -> XPropertySet:
        ...

    @overload
    @classmethod
    def add_labelled_control(
        cls,
        doc: XComponent,
        *,
        label: str,
        comp_kind: FormComponentKind | str,
        x: int | UnitT,
        y: int | UnitT,
        height: int | UnitT,
        lbl_styles: Iterable[StyleT] = ...,
        ctl_styles: Iterable[StyleT] = ...,
    ) -> XPropertySet:
        ...

    @classmethod
    def add_labelled_control(cls, *args, **kwargs) -> XPropertySet:
        """
        Create a label and data field control, with the label preceding the control

        Args:
            doc (XComponent): Component
            label (str): Label to assign to control
            comp_kind (FormComponentKind | str): Kind of control such as ``CheckBox``.
            x (int): Control X position
            y (int): Control Y Position
            height (int): control height
            lbl_styles (Iterable[StyleT], optional): One or more styles to apply on the label portion of control.
            ctl_styles (Iterable[StyleT], optional): One or more styles to apply on the Textbox portion of control.

        Returns:
            XPropertySet: DataField Control Property Set

        .. versionchanged:: 0.9.2
            Added ``lbl_styles`` and ``ctl_styles`` arguments.
        """
        lbl_styles = kwargs.pop("lbl_styles", None)
        ctl_styles = kwargs.pop("ctl_styles", None)
        ordered_keys = (1, 2, 3, 4, 5, 6)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("doc", "label", "comp_kind", "x", "y", "height")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("addLabelledControl() got an unexpected keyword argument")
            ka[1] = kwargs.get("doc", None)
            ka[2] = kwargs.get("label", None)
            ka[3] = kwargs.get("comp_kind", None)
            keys = ("x", "y")
            for key in keys:
                if key in kwargs:
                    ka[4] = kwargs[key]
                    break
            if count == 4:
                return ka
            ka[5] = kwargs.get("y", None)
            ka[6] = kwargs.get("height", None)
            return ka

        if count not in (4, 6):
            raise TypeError("addLabelledControl() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 4:
            result = cls._add_labelled_control(
                doc=kargs[1],
                label=kargs[2],
                comp_kind=kargs[3],
                x=2,
                y=kargs[4],
                height=6,
                lbl_styles=lbl_styles,
                ctl_styles=ctl_styles,
            )
        else:
            return cls._add_labelled_control(
                doc=kargs[1],
                label=kargs[2],
                comp_kind=kargs[3],
                x=kargs[4],
                y=kargs[5],
                height=kargs[6],
                lbl_styles=lbl_styles,
                ctl_styles=ctl_styles,
            )

        return result

    @classmethod
    def _add_labelled_control(
        cls,
        doc: XComponent,
        *,
        label: str,
        comp_kind: FormComponentKind | str,
        x: int | UnitT,
        y: int | UnitT,
        height: int | UnitT,
        width: int = 40,
        lbl_styles: Iterable[StyleT] | None,
        ctl_styles: Iterable[StyleT] | None,
    ) -> XPropertySet:
        try:
            name = f"{label}_label"
            # create label (fixed text) control
            label_props = cls.add_control(
                doc=doc,
                name=name,
                label=label,
                comp_kind=FormComponentKind.FIXED_TEXT,
                x=x,
                y=y,
                width=width,
                height=height,
                styles=lbl_styles,
            )

            # create data field control
            x_value = cls._get_unit100_value(x) + 26
            y_value = cls._get_unit100_value(y) + 26
            ctl_props = cls.add_control(
                doc=doc,
                name=label,
                label=None,
                comp_kind=comp_kind,
                x=x_value,
                y=y_value,
                width=width,
                height=height,
                styles=ctl_styles,
            )
            ctl_props.setPropertyValue("DataField", label)

            # add label props to the control
            ctl_props.setPropertyValue("LabelControl", label_props)
            return ctl_props
        except Exception:
            raise

    # endregion add_labelled_control

    # region    add_button
    @classmethod
    def add_button(
        cls,
        doc: XComponent,
        *,
        name: str,
        label: str | None,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        anchor_type: TextContentAnchorType | None = None,
        parent_form: XNameContainer | None = None,
        styles: Iterable[StyleT] | None = None,
    ) -> XPropertySet:
        """
        Adds a button control.

        By Default the button has no tab stop and does not focus on click.

        Args:
            doc (XComponent): Component
            name (str): Button name
            label (str | None): Button Label
            x (int): Button X position
            y (int): Button Y position
            height (int): Button Height
            width (int, optional): Button Height. Defaults to 6.
            anchor_type (TextContentAnchorType | None): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            styles (Iterable[StyleT], optional): One or more styles to apply.

        Returns:
            XPropertySet: Button Property Set

        .. versionchanged:: 0.9.2
            Added ``styles`` argument.
        """
        try:
            btn_props = cls.add_control(
                doc=doc,
                name=name,
                label=label,
                comp_kind=FormComponentKind.COMMAND_BUTTON,
                x=x,
                y=y,
                width=width,
                height=height,
                anchor_type=anchor_type,
                parent_form=parent_form,
                styles=styles,
            )
            # don't want button to be accessible by the "tab" key
            btn_props.setPropertyValue("Tabstop", False)

            # the button should not steal focus when clicked
            btn_props.setPropertyValue("FocusOnClick", False)

            return btn_props
        except Exception:
            raise

    # endregion add_button

    @classmethod
    def add_list(
        cls,
        doc: XComponent,
        name: str,
        entries: Iterable[str],
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        *,
        styles: Iterable[StyleT] | None = None,
    ) -> XPropertySet:
        """
        Adds a list

        Args:
            doc (XComponent): Component
            name (str): List Name
            entries (Iterable[str]): List Entries
            x (int): List X position
            y (int): List Y Position
            width (int): List Width
            height (int): List Height
            styles (Iterable[StyleT], optional): One or more styles to apply.

        Returns:
            XPropertySet: List property set

        .. versionchanged:: 0.9.2
            Added ``styles`` argument.
        """
        try:
            lst_props = cls.add_control(
                doc=doc,
                name=name,
                label=None,
                comp_kind=FormComponentKind.LIST_BOX,
                x=x,
                y=y,
                width=width,
                height=height,
                styles=styles,
            )
            items = mProps.Props.any(*list(entries))
            # lst_props.setPropertyValue("DefaultSelection", 0)
            uno.invoke(lst_props, "setPropertyValue", ("ListSource", items))  # type: ignore
            uno.invoke(lst_props, "setPropertyValue", ("DefaultSelection", mProps.Props.any(0)))  # type: ignore
            lst_props.setPropertyValue("Dropdown", True)
            lst_props.setPropertyValue("MultiSelection", False)
            uno.invoke(lst_props, "setPropertyValue", ("StringItemList", items))  # type: ignore
            uno.invoke(lst_props, "setPropertyValue", ("SelectedItems", mProps.Props.any(0)))  # type: ignore
            return lst_props
        except Exception:
            raise

    @classmethod
    def add_database_list(
        cls,
        doc: XComponent,
        *,
        name: str,
        sql_cmd: str,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        styles: Iterable[StyleT] | None = None,
    ) -> XPropertySet:
        """
        Add a list with a SQL command as it data source

        Args:
            doc (XComponent): Component
            name (str): List Name
            sql_cmd (str): SQL Command
            x (int): List X position
            y (int): List Y Position
            width (int): List Width
            height (int): List Height
            styles (Iterable[StyleT], optional): One or more styles to apply.

        Returns:
            XPropertySet: List property set

        .. versionchanged:: 0.9.2
            Added ``styles`` argument.
        """
        try:
            lst_props = cls.add_control(
                doc=doc,
                name=name,
                label=None,
                comp_kind=FormComponentKind.DATABASE_LIST_BOX,
                x=x,
                y=y,
                width=width,
                height=height,
                styles=styles,
            )
            lst_props.setPropertyValue("Dropdown", True)
            lst_props.setPropertyValue("MultiSelection", False)
            lst_props.setPropertyValue("BoundColumn", 0)

            # data-aware properties
            lst_props.setPropertyValue("ListSourceType", ListSourceType.SQL)
            uno.invoke(lst_props, "setPropertyValue", ("ListSource", mProps.Props.any(sql_cmd)))  # type: ignore
            return lst_props
        except Exception:
            raise

    @staticmethod
    def create_grid_column(grid_model: XControlModel, data_field: str, col_kind: str, width: int) -> None:
        """
        Adds a column to a gird

        Args:
            grid_model (XControlModel): Grid control Model
            data_field (str): the database field to which the column should be bound
            col_kind (str):  the column type such as "NumericField"
            width (int): the column width (in mm). If 0, no width is set.

        Returns:
            None:
        """
        # column container and factory
        col_container = mLo.Lo.qi(XIndexContainer, grid_model, True)
        col_factory = mLo.Lo.qi(XGridColumnFactory, grid_model, True)

        # create the column
        col_props = col_factory.createColumn(col_kind)
        col_props.setPropertyValue("DataField", data_field)
        col_props.setPropertyValue("Label", data_field)
        col_props.setPropertyValue("Name", data_field)
        if width > 0:
            col_props.setPropertyValue("Width", width * 10)

        # add properties column to container
        col_container.insertByIndex(col_container.getCount(), col_props)

    # endregion create controls

    # region  bind form to database
    @staticmethod
    def bind_form_to_table(xform: XForm, src_name: str, tbl_name: str) -> None:
        """
        Bind the form to the database in the src_name URL

        Args:
            xform (XForm): Form
            src_name (str): Source Name URL
            tbl_name (str): Table Name

        Returns:
            None:
        """
        mProps.Props.set(xform, DataSourceName=src_name, Command=tbl_name, CommandType=CommandType.TABLE)

    @staticmethod
    def bind_form_to_sql(xform: XForm, src_name: str, cmd: str) -> None:
        """
        Bind the form to the database in the ``src_name`` URL, and send a SQL cmd

        Args:
            xform (XForm): Form
            src_name (str): Source Name URL
            cmd (str): Command

        Returns:
            None:
        """
        mProps.Props.set(xform, DataSourceName=src_name, Command=cmd, CommandType=CommandType.COMMAND)
        # cannot use CommandType.TABLE for the SELECT cmd

    # endregion  bind form to database

    # region  bind a macro to a form control
    @staticmethod
    def assign_script(
        ctl_props: XPropertySet, interface_name: str, method_name: str, script_name: str, loc: str
    ) -> None:
        """
        Binds a macro to a form control

        Args:
            ctl_props (XPropertySet): Properties of control
            interface_name (str): Interface Name
            method_name (str): Method Name
            script_name (str): Script Name
            loc (str): can be user, share, document, and extensions

        Returns:
            None:

        See Also:
            `Scripting Framework URI Specification <https://wiki.openoffice.org/wiki/Documentation/DevGuide/Scripting/Scripting_Framework_URI_Specification>`_
        """
        # https://wiki.openoffice.org/wiki/Documentation/DevGuide/WritingUNO/XInterface
        # In C++, two objects are the same if their XInterface are the same. The queryInterface() for XInterface will have to
        # be called on both. In Java, check for the identity by calling the runtime function
        # com.sun.star.uni.UnoRuntime.areSame().
        try:
            props_child = mLo.Lo.qi(XChild, ctl_props, True)
            parent_form = mLo.Lo.qi(XIndexContainer, props_child.getParent(), True)

            pos = -1
            for i in range(parent_form.getCount()):
                child = mLo.Lo.qi(XPropertySet, parent_form.getByIndex(i))
                if mInfo.Info.is_same(child, ctl_props):
                    pos = i
                    break

            if pos == -1:
                mLo.Lo.print("Could not find control's position in form")
            else:
                mgr = mLo.Lo.qi(XEventAttacherManager, parent_form, True)
                ed = ScriptEventDescriptor(
                    interface_name,
                    method_name,
                    "",
                    "Script",
                    f"vnd.sun.star.script:{script_name}?language=Java&location={loc}",
                )

                mgr.registerScriptEvent(pos, ed)
        except Exception:
            raise

    # endregion  bind a macro to a form control

    # region Insert Controls
    @classmethod
    def insert_control_button(
        cls,
        doc: XComponent,
        *,
        label: str | None,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        parent_form: XNameContainer,
        name: str = "",
        height: int = 6,
        anchor_type: TextContentAnchorType | None = None,
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlButton:
        """
        Inserts a button control.

        By Default the button has tab stop and does focus on click.

        Args:
            doc (XComponent): Component
            name (str): Button name
            label (str | None): Button Label
            x (int): Button X position
            y (int): Button Y position
            height (int): Button Height
            width (int, optional): Button Height. Defaults to 6.
            anchor_type (TextContentAnchorType | None): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlButton: Button Control

        .. versionadded:: 0.13.8
        """
        if not name:
            name = cls.create_name(parent_form, "Button")
        if styles is None:
            # keeps type checker happy
            styles = ()
        try:
            btn_props = cls.add_button(
                doc=doc,
                name=name,
                label=label,
                x=x,
                y=y,
                width=width,
                height=height,
                anchor_type=anchor_type,
                parent_form=parent_form,
                styles=styles,
            )
            model = mLo.Lo.qi(XControlModel, btn_props, True)
            ctl = cls.get_control(doc, model)
            result = FormCtlButton(ctl)
            result.tab_stop = True
            return result
        except Exception:
            raise

    @classmethod
    def insert_control_check_box(
        cls,
        doc: XComponent,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        parent_form: XNameContainer,
        name: str = "",
        label: str = "",
        height: int | UnitT = 6,
        anchor_type: TextContentAnchorType | None = None,
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlCheckBox:
        """
        Inserts a check box control.

        Args:
            doc (XComponent): Component
            name (str): Button name
            label (str | None): Button Label
            x (int): Button X position
            y (int): Button Y position
            height (int): Button Height
            width (int, optional): Button Height. Defaults to 6.
            anchor_type (TextContentAnchorType | None): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlCheckBox: CheckBox Control

        .. versionadded:: 0.13.8
        """
        if styles is None:
            # keeps type checker happy
            styles = ()

        if not name:
            name = cls.create_name(parent_form, "CheckBox")
        try:
            btn_props = cls.add_control(
                doc=doc,
                name=name,
                label=label,
                comp_kind=FormComponentKind.CHECK_BOX,
                x=x,
                y=y,
                width=width,
                height=height,
                anchor_type=anchor_type,
                styles=styles,
            )
            model = mLo.Lo.qi(XControlModel, btn_props, True)
            ctl = cls.get_control(doc, model)
            return FormCtlCheckBox(ctl)
        except Exception:
            raise

    # endregion Insert Controls
