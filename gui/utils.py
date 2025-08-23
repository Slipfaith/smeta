from typing import Union


def format_rate(value: Union[int, float]) -> str:
    """Format rate according to GUI rules.

    Numbers are formatted with up to three decimals. If the formatted
    string ends with two or three zeros after the decimal separator,
    trim one trailing zero so that ``3.300`` becomes ``3.30`` and
    ``3.000`` becomes ``3.00``.
    """
    num = float(value)
    text = f"{num:.3f}"
    if text.endswith("000"):
        return text[:-1]
    if text.endswith("00"):
        return text[:-1]
    return text
