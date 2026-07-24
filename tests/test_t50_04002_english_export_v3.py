from cal_translator.formats.t50_04002.english_export_v3 import (
    try_enable_shape_printing,
)


class WritableShape:
    def __init__(self) -> None:
        self.PrintObject = False


class ReadOnlyShape:
    @property
    def PrintObject(self) -> bool:
        return True

    @PrintObject.setter
    def PrintObject(self, value: bool) -> None:
        raise AttributeError("PrintObject is read-only")


def test_shape_printing_is_enabled_when_property_is_writable() -> None:
    shape = WritableShape()

    assert try_enable_shape_printing(shape) is True
    assert shape.PrintObject is True


def test_read_only_shape_print_property_uses_excel_default() -> None:
    shape = ReadOnlyShape()

    assert try_enable_shape_printing(shape) is False
    assert shape.PrintObject is True
