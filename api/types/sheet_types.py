class SheetType:

    def __init__(self, index, name, description):
        self.index = 0
        self.name = name
        self.description = description

    def __hash__(self):
        return hash(self.index)


SCALAR_INPUT = SheetType(
    0, "SCALAR_INPUT", "Tab made up entirely from scalar inputs")

STANDARD_ROW_OPERATION = SheetType(
    1, 'STANDARD_ROW_OPERATION', 'Tab made up of columns where each calculation is deteremined by some data in the same row.'
)
