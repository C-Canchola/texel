class SheetType:

    def __init__(self, index, name, description):
        self.index = index
        self.name = name
        self.description = description

    def __hash__(self):
        return hash(self.index)

    def __eq__(self, other):

        if isinstance(other, int):
            return self.index == other
        if isinstance(other, float):
            return self.index == int(other)
        if isinstance(other, SheetType):
            return self.index == other.index
        return False


SCALAR_INPUT = SheetType(
    0, "SCALAR_INPUT", "Tab made up entirely from scalar inputs")

STANDARD_ROW_OPERATION = SheetType(
    1, 'STANDARD_ROW_OPERATION', 'Tab made up of columns where each calculation is deteremined by some data in the '
                                 'same row. '
)
TWO_DIMENSIONAL_LOOK_UP = SheetType(
    2, 'TWO_DIMENSIONAL_LOOK_UP', 'Tab where the column headers and rows are look up keys to the same data '
                                  'below and to the right.'
)
