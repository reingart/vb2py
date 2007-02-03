from vb2py.vbfunctions import *
from vb2py.vbdebug import *



def Log(Text):
    Debug.Print(Time + ' - ' + Text)

def test():
    a = vbObjectInitialize(objtype=String)
    b = UBound(a)
    a = vbObjectInitialize((10,), Variant)
    a[1] = 'hello'
    a[10] = 'bye'
    c = a(1)
    d = a(10)

def Factorial(X):
    _ret = None
    if X == 0:
        _ret = 1
    else:
        _ret = X * Factorial(X - 1)
    return _ret

def Add(X, Y):
    _ret = None
    _ret = X + Y
    return _ret

def EraseTest():
    a = vbObjectInitialize((10, 2,), Integer)
    for i in vbForRange(1, 10):
        for j in vbForRange(1, 2):
            a[i, j] = i + j
    Debug.Print(GetArrayRepr(a))
    Erase(a)
    Debug.Print(GetArrayRepr(a))

def GetArrayRepr(Arr):
    _ret = None
    total = 0
    for i in vbForRange(1, UBound(Arr, 1)):
        for j in vbForRange(1, UBound(Arr, 2)):
            total = total + Arr(i, j)
    _ret = total
    return _ret

# VB2PY (UntranslatedCode) Attribute VB_Name = "Globals"
