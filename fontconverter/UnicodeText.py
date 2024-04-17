from .CopticFontInterface import CopticFontInterface


class UnicodeText(CopticFontInterface):
    def __init__(self):
        self.letters = (
            "ⲁⲃⲅⲇⲉⲋⲍⲏⲑⲓⲕⲗⲙⲛⲝⲟⲡⲣⲥⲧⲩⲫⲭⲯⲱϣϥϧϩϫϭϯⲀⲂⲄⲆⲈⲊⲌⲎⲐⲒⲔⲖⲘⲚⲜⲞⲠⲢⲤⲦⲨⲪⲬⲮⲰϢϤϦϨϪϬϮ̀.̅:"
        )
        self.is_combining = True
