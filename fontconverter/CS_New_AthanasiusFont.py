from .CopticFontInterface import CopticFontInterface


class CS_New_AthanasiusFont(CopticFontInterface):
    def __init__(self):
        self.letters = (
            "abgde6zy;iklmnxoprctuv,'wsfqhj[]ABGDE6ZY:IKLMNXOPRCTUV<\"WSFQHJ{}`.=@"
        )
        self.is_combining = False
