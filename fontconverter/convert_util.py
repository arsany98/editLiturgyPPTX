from .CopticFontFactory import CopticFontFactory


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
    if src.get_is_combining() == True and dest.get_is_combining() == False:
        result = shift_punc_left(result, font2[-2] + font2[-4])
    elif src.get_is_combining() == False and dest.get_is_combining() == True:
        result = shift_punc_right(result, font2[-2] + font2[-4])

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
