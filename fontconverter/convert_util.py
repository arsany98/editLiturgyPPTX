from .CopticFontFactory import CopticFontFactory


def fix_punc(text, prev, next, srcFont, destFont):
    src = CopticFontFactory.getFont(srcFont)
    dest = CopticFontFactory.getFont(destFont)
    destLetters = dest.get_letters()
    punc = destLetters[-2] + destLetters[-4]
    if src.get_is_combining() == True and dest.get_is_combining() == False:
        if text[0] in punc:
            prev = prev[:-1] + text[0] + prev[-1]
            text = text[1:]
        text = shift_punc_left(text, punc)
    elif src.get_is_combining() == False and dest.get_is_combining() == True:
        if text[-1] in punc:
            next = text[-1] + next
            text = text[:-1]
        text = shift_punc_right(text, punc)
    return prev, text, next


def convert(text, srcFont, destFont):
    src = CopticFontFactory.getFont(srcFont)
    dest = CopticFontFactory.getFont(destFont)
    font1 = src.get_letters()
    font2 = dest.get_letters()
    result = ""
    for c in text:
        if font1.find(c) != -1:
            result += font2[font1.index(c)]
        else:
            result += c
    return result


def shift_punc_left(text, punc):
    text = list(text)
    for i in range(1, len(text)):
        if text[i] in punc:
            text[i], text[i - 1] = text[i - 1], text[i]
    return "".join(text)


def shift_punc_right(text, punc):
    text = list(text)
    i = 0
    while i < len(text) - 1:
        if text[i] in punc:
            text[i + 1], text[i] = text[i], text[i + 1]
            i += 1
        i += 1
    return "".join(text)
