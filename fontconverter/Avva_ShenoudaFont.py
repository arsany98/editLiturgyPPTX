from .CopticFontInterface import CopticFontInterface


class Avva_ShenoudaFont(CopticFontInterface):
    def __init__(self):
        self.letters = (
            "abjde6z30iklmn7oprctvfxyw24qhgs5ABJDE^Z#)IKLMN&OPRCTVFXYW@$QHGS%`.=:"
        )
        self.is_combining = False
