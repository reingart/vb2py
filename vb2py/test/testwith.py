# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

from unittest import *
from testframework import *

# << With tests >> (1 of 2)
# Simple test
tests.append(("""
Set _a = New Collection
With _a
	b = .Count()
End With
""", {"b" : 0}))

# Nested test
tests.append(("""
Dim _a As New Collection, _b As New Collection
_a.Add 24
_a.Add 25
_b.Add 1
With _a
	aa = .Count()
	With _b
		bb = .Count()
	End With
End With
""", {"aa" : 2, "bb" : 1}))

# Nested test with LHS implicit objects
tests.append(("""
Dim _a As New Collection, _b As New Collection
With _a
	.Add 24
	.Add 25
	aa = .Count()
	With _b
		.Add 1
		bb = .Count()
	End With
End With
""", {"aa" : 2, "bb" : 1}))

# Nested test with LHS implicit objects 2
tests.append(("""
Dim _a As New Collection, _b As New Collection
With _a
	.Add 24
	.Add 25
	With _b
		.Add 1
		bb = .Count()
	End With
	aa = .Count()
End With
""", {"aa" : 2, "bb" : 1}))
# << With tests >> (2 of 2)
# Tricky little nesting problem
tests.append(("""
Type _Container3
   Value As Integer
End Type

Type _Container2
   Value As Integer
   Obj As _Container3
End Type

Type _Container1
   Value As Integer
   Obj As _Container2
End Type


Dim _a As _Container1
Dim _b As _Container2
Dim _c As _Container3

Set _a.Obj = _b
Set _b.Obj = _c
_c.Value = 10

With _a
	With .Obj ' ie _b
	   With .Obj ' ie _c
			val = .Value
	   End With 
	End With
End With
""", {"val" : 10}))
# -- end -- << With tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
	main()
