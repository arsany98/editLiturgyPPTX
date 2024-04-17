from .Avva_ShenoudaFont import Avva_ShenoudaFont
from .CS_New_AthanasiusFont import CS_New_AthanasiusFont
from .UnicodeText import UnicodeText


class CopticFontFactory:
    def getFont(fontName):
        if fontName == "Avva_Shenouda":
            return Avva_ShenoudaFont()
        elif fontName == "CS New Athanasius":
            return CS_New_AthanasiusFont()
        elif fontName == "Coptic New Athanasius":
            return UnicodeText()
        return None
