from unittest import *
from testframework import *

# << Erase tests >>
# Simple test of erase
tests.append(("""
Function _GetArrayRepr(Arr)
total = 0
For i = 1 To UBound(Arr, 1)
    For j = 1 To UBound(Arr, 2)
       total = total + Arr(i, j)
    Next j
Next i
_GetArrayRepr = total
End Function

Dim _a(10, 2) As Integer
For _i = 1 To 10
    For _j = 1 To 2
        _a(_i, _j) = _i + _j
    Next _j
Next _i
t1 = _GetArrayRepr(_a)
Erase _a
t2 = _GetArrayRepr(_a)
""", {"t1" : 140, "t2" : 0}))

# Simple test of erase with two arrays
tests.append(("""
Function _GetArrayRepr(Arr)
total = 0
For i = 1 To UBound(Arr, 1)
    For j = 1 To UBound(Arr, 2)
       total = total + Arr(i, j)
    Next j
Next i
_GetArrayRepr = total
End Function

Dim _a(10, 2) As Integer
Dim _b(10, 2) As Integer
For _i = 1 To 10
    For _j = 1 To 2
        _a(_i, _j) = _i + _j
        _b(_i, _j) = _i + _j
    Next _j
Next _i
t1a = _GetArrayRepr(_a)
t1b = _GetArrayRepr(_b)
Erase _a, _b
t2a = _GetArrayRepr(_a)
t2b = _GetArrayRepr(_b)
""", {"t1a" : 140, "t2a" : 0, "t1b" : 140, "t2b" : 0}))
# -- end -- << Erase tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
    main()
