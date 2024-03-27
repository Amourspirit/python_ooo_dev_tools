from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple

from com.sun.star.form.runtime import XFilterController

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import XControl
    from com.sun.star.form.runtime import XFilterControllerListener
    from ooodev.utils.type_var import UnoInterface


class FilterControllerPartial:
    """
    Partial Class for XFilterController.

    This interface does not really provide an own functionality, it is only for easier runtime identification of form components.
    """

    def __init__(self, component: XFilterController, interface: UnoInterface | None = XFilterController) -> None:
        """
        Constructor

        Args:
            component (XFilterController): UNO Component that implements ``com.sun.star.form.runtime.XFilterController``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``None``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XFilterController
    def add_filter_controller_listener(self, listener: XFilterControllerListener) -> None:
        """
        Registers a listener to be notified of certain changes in the form based filter.

        Registering the same listener multiple times results in multiple notifications of the same event, and also requires multiple revocations of the listener.
        """
        self.__component.addFilterControllerListener(listener)

    def append_empty_disjunctive_term(self) -> None:
        """
        Appends an empty disjunctive term to the list of terms.
        """
        self.__component.appendEmptyDisjunctiveTerm()

    def get_filter_component(self, component: int) -> XControl:
        """
        Retrieves the filter component with the given index.

        The filter control has the same control model as the control which it stands in for. Consequently, you can use this method to obtain the database column which the filter control works on, by examining the control model's BoundField property.

        Raises:
            com.sun.star.lang.IndexOutOfBoundsException: ``IndexOutOfBoundsException``
        """
        return self.__component.getFilterComponent(component)

    def get_predicate_expressions(self) -> Tuple[Tuple[str, ...], ...]:
        """
        retrieves the entirety of the predicate expressions represented by the filter controller.

        Each element of the returned sequence is a disjunctive term, having exactly FilterComponents elements, which denote the single predicate expressions of this term.
        """
        return self.__component.getPredicateExpressions()

    def remove_disjunctive_term(self, term: int) -> None:
        """
        removes a given disjunctive term

        Raises:
            com.sun.star.lang.IndexOutOfBoundsException: ``IndexOutOfBoundsException``
        """
        self.__component.removeDisjunctiveTerm(term)

    def remove_filter_controller_listener(self, listener: XFilterControllerListener) -> None:
        """
        Revokes a listener which was previously registered to be notified of certain changes in the form based filter.
        """
        self.__component.removeFilterControllerListener(listener)

    def set_predicate_expression(self, component: int, term: int, predicate_expression: str) -> None:
        """
        sets a given predicate expression

        Raises:
            com.sun.star.lang.IndexOutOfBoundsException: ``IndexOutOfBoundsException``
        """
        self.__component.setPredicateExpression(component, term, predicate_expression)

    @property
    def active_term(self) -> int:
        """
        Gets/Sets - Denotes the active term of the filter controller.
        """
        return self.__component.ActiveTerm

    @active_term.setter
    def active_term(self, value: int) -> None:
        self.__component.ActiveTerm = value

    @property
    def disjunctive_terms(self) -> int:
        """
        Gets the number of disjunctive terms of the filter expression represented by the form based filter.
        """
        return self.__component.DisjunctiveTerms

    @property
    def filter_components(self) -> int:
        """
        Gets the number of filter components, or filter controls, which the filter controller is responsible for.

        This number is constant during one session of the form based filter.
        """
        return self.__component.FilterComponents

    # endregion XFilterController
