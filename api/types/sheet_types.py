class SheetType:

    def __init__(self, index, name, description):
        self.index = 0
        self.name = name
        self.description = description


SCALAR_INPUT = SheetType(
    0, "SCALAR_INPUT", "Tab made up entirely from scalar inputs")
