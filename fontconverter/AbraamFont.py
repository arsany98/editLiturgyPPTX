from .CopticFontInterface import CopticFontInterface


class AbraamFont(CopticFontInterface):
    def __init__(self):
        self.letters = (
            "abgde,zhqiklmn[oprctuvxyw]f'\\js;ABGDE<ZHQIKLMN{OPRCTUVXYW}F\"|JS:`.?>"
        )
        self.is_combining = False
