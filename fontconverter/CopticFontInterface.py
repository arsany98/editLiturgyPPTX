class CopticFontInterface:
    letters = ""
    is_combining = False

    def get_letters(self):
        return self.letters

    def get_is_combining(self):
        return self.is_combining
