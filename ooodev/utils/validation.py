def check(exp: bool, err_msg: str, msg2: str = "out of range") -> None:
    """
    Preforms a validation check

    Args:
        exp (bool): Boolean expression such as ``self.value > 3``
        err_msg (str): Error Message
        msg2 (str, optional): Extended error message. Defaults to "out of range".

    Raises:
        AssertionError: If Validation fails
    """
    if not exp:
        # print(f"Type failure: {err_msg} {msg2}")
        # Normally, you'd use assert to throw an exception
        raise AssertionError(f"Type failure: {err_msg} {msg2}")