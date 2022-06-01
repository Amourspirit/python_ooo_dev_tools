# coding: utf-8
from typing import Any

class MissingInterfaceError(Exception):
    """Error when a interface is not found for a uno object"""
    def __init__(self, interface: Any, message:Any = None) -> None:
        """
        MissingInterfaceError constructor

        Args:
            interface (Any): Missing Interface that caused error
            message (Any, optional): Message of error
        """
        if message is None:
            try:
                message = f"Missing interface {interface.__pyunointerface__}"
            except AttributeError:
                message = "Missing Uno Interface Error"
        super().__init__(interface, message)
    
    def __str__(self) -> str:
        return repr(self.args[1])


class CellError(Exception):
    """Cell error"""
    pass

class GoalDivergenceError(Exception):
    """Error when goal seek result divergence is too high"""
    def __init__(self, divergence: float, message:Any = None) -> None:
        """
        GoalDivergenceError Constructor

        Args:
            divergence (float): divergence amount
            message (Any, optional): Message of error
        """
        if message is None:
                message = f"Divergence error: {divergence:.4f}"
        super().__init__(divergence, message)
    
    def __str__(self) -> str:
        return repr(self.args[1])