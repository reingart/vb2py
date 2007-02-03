from unittest import *
from testframework import *

# << Dim tests >> (1 of 5)
# Untyped
tests.append(("""
Dim a, b, c
a = 10
b = "hello"
c = 123.5
""", {"a" : 10, "b" : "hello", "c" : 123.5}
))

# Typed
tests.append(("""
Dim a As Integer, b As String, c As Single
a = 10
b = "hello"
c = 123.5
""", {"a" : 10, "b" : "hello", "c" : 123.5}
))

# Array of integers
tests.append(("""
Dim a(3) As Integer
a(0) = 1
a(1) = 2
a(2) = 3
a(3) = 4
""", {"a" : [1,2,3,4]}
))

# Array of integers with non-zero offset
tests.append(("""
Dim a(1 To 3) As Integer
a(1) = 2
a(2) = 3
a(3) = 4
""", {"a" : [2,3,4]}
))

# String size indicator
tests.append(("""
Dim _a As String * 20
b = Len(_a)
""", {"b" : 20}
))
# << Dim tests >> (2 of 5)
# Redims
tests.append(("""
Dim _a() As Single
ReDim _a(10)
b = len(_a)
""", {"b" : 11}))

# Redims without preserving
tests.append(("""
Dim _a(10) As Single
_a(5) = 10
ReDim _a(10)
b = 0 + _a(5)
""", {"b" : 0}))

# Redims with preserving
tests.append(("""
Dim _a(10) As Single
_a(5) = 10
ReDim Preserve _a(10)
b = 0 + _a(5)
""", {"b" : 10}))

# Redims with preserving
tests.append(("""
Dim _a(10, 5) As Single
_a(5, 1) = 10
_a(8, 4) = 20
ReDim Preserve _a(20, 10)
b = _a(5, 1)
c = _a(8, 4)
""", {"b" : 10, "c" : 20}))

# Redims with preserving
tests.append(("""
Dim _a(10) As Single
_a(5) = 10
_a(8) = 10
ReDim Preserve _a(5)
b = _a(5)
""", {"b" : 10}))
# << Dim tests >> (3 of 5)
# Double dim!
tests.append(("""
Dim _a(10), _b(10)
for _i = 1 To 10
   _a(_i) = _i+1
   _b(_i) = _i*10
Next _i
'
c = _b(_a(1))
d = _b(_a(2))
""", {"c" : 20, "d" : 30}
))

# Double dim set 
tests.append(("""
Dim _a(10), _b(10)
for _i = 1 To 10
   _a(_i) = _i+1
   _b(_i) = _i*10
Next _i
'
_b(_a(1)) = 101
_b(_a(2)) = 202
'
c = _b(2)
d = _b(3)
""", {"c" : 101, "d" : 202}
))

# Get Dim Fn!
tests.append(("""
Dim  _b(10)
for _i = 1 To 10
   _b(_i) = _i*10
Next _i
'
Function _a(x)
   _a = x+1
End Function
'
c = _b(_a(1))
d = _b(_a(2))
""", {"c" : 20, "d" : 30}
))

# Set Dim Fn
tests.append(("""
Dim  _b(10)
for _i = 1 To 10
   _b(_i) = _i*10
Next _i
'
Function _a(x)
   _a = x+1
End Function
'
_b(_a(1)) = 101
_b(_a(2)) = 202
'
c = _b(2)
d = _b(3)
""", {"c" : 101, "d" : 202}
))
# << Dim tests >> (4 of 5)
# Bug #810403  - Empty Dim should still create an array
tests.append(("""
Global _a() As String
b = (_a="") ' Check we got an array not a string
ReDim _a(10)
_a(1) = "hello"
_a(10) = "bye"
c = _a(1)
d = _a(10)
""", {"b" : 0, "c" : "hello", "d" : "bye"}
))
# << Dim tests >> (5 of 5)
# Sub
tests.append(("""

Sub _f()
Dim aa(1)
aa(1) = 10
End Sub


""", {}
))

# Function
tests.append(("""
Function _f()
Dim aa(1)
aa(1) = 10
_f = aa(1)
End Function

a = _f

""", {"a" : 10, }
))
# -- end -- << Dim tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
    main()
