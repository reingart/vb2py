from vb2py.vbfunctions import *



def SplitOnWord(Text, Letter):
    _ret = None
    posn = 1
    Words = vbObjectInitialize((0,), Variant)
    for i in vbForRange(1, Len(Text)):
        if Mid(Text, i, 1) == Letter or i == Len(Text):
            Words = vbObjectInitialize((UBound(Words) + 1,), Variant, Words)
            if i == Len(Text):
                ThisWord = Mid(Text, posn, i - posn + 1)
            else:
                ThisWord = Mid(Text, posn, i - posn)
            Words[UBound(Words) - 1] = ThisWord
            posn = i + 1
    #
    Words = vbObjectInitialize((UBound(Words) - 1,), Variant, Words)
    _ret = Words
    #
    return _ret

def Factorial(n):
    _ret = None
    if n == 0:
        _ret = 1
    else:
        _ret = n * Factorial(n - 1)
    return _ret

# VB2PY (UntranslatedCode) Attribute VB_Name = "Utils"
