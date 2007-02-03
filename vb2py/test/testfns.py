from unittest import *
from testframework import *

# << Fn tests >> (1 of 5)
# Simple function
tests.append(("""
Function _square(x)
    _square = x*x
End Function
a = _square(10)
""", {"a" : 100}))

# Simple function with a type
tests.append(("""
Function _square(x) As Single
    _square = x*x
End Function
a = _square(10)
""", {"a" : 100}))

# Accidental non-return
tests.append(("""
Function _square(x)
End Function
a = _square(10)
""", {"a" : None}))

# Function calling a function
tests.append(("""
Function _double(x)
    _double = 2*x
End Function

Function _squaredouble(x)
    _squaredouble = _double(x) * _double(x)
End Function

a = _squaredouble(10)
""", {"a" : 400}))

# Function with exit
tests.append(("""
Function _square(x)
    If x > 0 Then
        _square = x*x
        Exit Function
    End If
    _square = -(x*x)
End Function
a = _square(10)
b = _square(-10)
""", {"a" : 100, "b" : -100}))

# Simple function with no arguments, must detect that it is a function call
tests.append(("""
Function _square()
    _square = 100
End Function
a = _square
""", {"a" : 100}))
# << Fn tests >> (2 of 5)
# Recursive function
tests.append(("""
Function _factorial(x)
    If x = 0 Then
        _factorial = 1
    Else
        _factorial = _factorial(x-1) * x
    End If
End Function
a = _factorial(10)
""", {"a" : 3628800}))
# << Fn tests >> (3 of 5)
# Lots of arguments
tests.append(("""
Function _sum(a,b,c,d,e,f,g,h,i)
    _sum = a+b+c+d+e+f+g+h+i
End Function
a = _sum(1,2,3,4,5,6,7,8,9)
""", {"a" : 45}))

# Lots of arguments with types
tests.append(("""
Function _sum(a As Integer,b As Single,c,d As String,e,f,g As Single,h,i)
    _sum = a+b+c+d+e+f+g+h+i
End Function
a = _sum(1,2,3,4,5,6,7,8,9)
""", {"a" : 45}))

# Some arguments with options
tests.append(("""
Function _sum(a, b, Optional c)
    If IsMissing(c) Then c = 10
    _sum = a+b+c
End Function
a = _sum(1,2,3)
b = _sum(1,2)
""", {"a" : 6, "b" : 13}))

# Some arguments with options and defaults
tests.append(("""
Function _sum(a, b, Optional c=10)
    _sum = a+b+c
End Function
a = _sum(1,2,3)
b = _sum(1,2)
""", {"a" : 6, "b" : 13}))
# << Fn tests >> (4 of 5)
# Function with named arguments
tests.append(("""
Function _sum(Optional x=1, Optional y=2, Optional z=3)
    _sum = x + y + z
End Function
a = _sum(10, 20, 30)
b = _sum(x:=10)
c = _sum(z:=10)
d = _sum()
f = _sum(x:=10, y:=20, z:=30)
""", {"a" : 60, "b" : 15, "c" : 13, "d" : 6, "f" : 60}))
# << Fn tests >> (5 of 5)
# Function with named arguments
tests.append(("""
Function _sum(Optional x=1, Optional y=2, Optional z=3)
    _sum = x + y + z
End Function
a = _sum(, , 30)
b = _sum(10,,30)
""", {"a" : 33, "b" : 42,}))
# -- end -- << Fn tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
    main()
