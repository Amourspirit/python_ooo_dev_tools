
# coding: utf-8

class TableHelper:

    @classmethod
    def make_cell_name(cls, row:int, col: int) -> str:
        """
        Convert given row and column number to ``A1`` style cell name.

        Args:
            row (int): Row number. This is a 1 based value.
            col (int): Column Number. This is 1 based value.

        Raises:
            ValueError: If row or col value < 1

        Returns:
            str: row and col as cell name such as A1, AB3
        """
        if row < 1:
            raise ValueError(f"row is one based. Value cannot be less then 1: {row}")
        return f"{cls.make_column_name(col)}{row}"
    
    @staticmethod
    def make_column_name(col: int) -> str: # col is 1 based
        """
        Makes a cell style name. eg: A, B, C, ... AA, AB, AC

        Args:
            col (int): Column number. This is a one based value.

        Raises:
            ValueError: If col value < 1

        Returns:
            str: column name. eg: A, B, C, ... AA, AB, AC
        """
        if col < 1:
            raise ValueError(f"col is one based. Value cannot be less then 1: {col}")
        str_col = str()
        div = col 
        while div:
            (div, mod) = divmod(div-1, 26) # will return (x, 0 .. 25)
            str_col = chr(mod + 65) + str_col
        return str_col