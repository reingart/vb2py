# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

from unittest import *
from testframework import *

# << Sub tests >> (1 of 4)
# Simple subroutine - use "global" to see results of the subroutine
tests.append(("""
Dim _lst(10) As Single
Sub _SetValue(Index As Integer, Value As String)
	_lst(Index) = Value
End Sub

_SetValue 5, "hello"
a = _lst(5)
""", {"a" : "hello"}))

# Simple subroutine with an exit
tests.append(("""
Dim _lst(10) As Single
Sub _SetValue(Index As Integer, Value1 As String, Value2 As String)
	_lst(Index) = Value1
	Exit Sub
	_lst(Index) = Value2
End Sub

_SetValue 5, "hello", "bye"
a = _lst(5)
""", {"a" : "hello"}))

# Simple sub calling a sub
tests.append(("""
Dim _lst(10) As Single
Sub _SetValue(Index As Integer, Value As String)
	_lst(Index) = Value
End Sub

Sub _SetFive(Value)
	_SetValue 5, Value
End Sub

_SetFive "hello"
a = _lst(5)
""", {"a" : "hello"}))

# Subroutine empty but for a comment - this can be a syntax error in Python
tests.append(("""
Sub _SetValue()
	' Nothing to see here
End Sub
""", {}))
# << Sub tests >> (2 of 4)
# Recursive sub
tests.append(("""
Dim _lst(10) As Single
Sub _SetValue(Index As Integer, Value)
	_lst(Index) = Value
	If Index < 10 Then 
		_SetValue Index+1, Value+1
	End If
End Sub

_SetValue 1, 1
a = _lst(5)
""", {"a" : 5}))
# << Sub tests >> (3 of 4)
# Optional arguments
tests.append(("""
Dim _lst(10) As Single
Sub _SetValue(Index As Integer, Optional Value=10)
	_lst(Index) = Value
End Sub

_SetValue 5, "hello"
_SetValue 6
a = _lst(5)
b = _lst(6)
""", {"a" : "hello", "b" : 10}))

# Optional arguments
tests.append(("""
Dim _lst(10) As Single
Sub _SetValue(Index As Integer, Optional Value)
	If IsMissing(Value) Then Value = 10
	_lst(Index) = Value
End Sub

_SetValue 5, "hello"
_SetValue 6
a = _lst(5)
b = _lst(6)
""", {"a" : "hello", "b" : 10}))

# Optional arguments with hex numbers
tests.append(("""
Dim _lst(10) As Single
Sub _SetValue(Index As Integer, Optional Value=&HA)
	_lst(Index) = Value
End Sub

_SetValue 5, "hello"
_SetValue 6
a = _lst(5)
b = _lst(6)
""", {"a" : "hello", "b" : 10}))
# << Sub tests >> (4 of 4)
# Sub with named arguments
tests.append(("""
Dim _vals(10)
Sub _sum(Optional x=1, Optional y=2, Optional z=3)
	_vals(1) = x + y + z
End Sub

_sum 10, 20, 30
a = _vals(1)
_sum x:=10
b = _vals(1)
_sum z:=10 
c = _vals(1)
_sum
d = _vals(1)
_sum x:=10, y:=20, z:=30
f = _vals(1)
""", {"a" : 60, "b" : 15, "c" : 13, "d" : 6, "f" : 60}))
# -- end -- << Sub tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
	main()
