# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

from unittest import *
from testframework import *

# << For tests >> (1 of 4)
# Simple test - note we are not interested in the value of the loop variable at the end of the loop
# hence the _var name
tests.append(("""
j = 0
For _var = 1 To 10
	j = j + _var
Next _var
""", {"j" : 55}))

# Simple test with step
tests.append(("""
j = 0
For _var = 1 To 10 Step 2
	j = j + _var
Next _var
""", {"j" : 25}))

# Simple test with reverse step
tests.append(("""
j = 0
For _var = 10 To 1 Step -1
	j = j + _var
Next _var
""", {"j" : 55}))

# Empty loop
tests.append(("""
j = 0
For _var = 10 To 1
	j = j + _var
Next _var
""", {"j" : 0}))

# Breaking out of the loop
tests.append(("""
j = 0
For _var = 1 To 10
	j = j + _var
	Exit For
Next _var
""", {"j" : 1}))
# << For tests >> (2 of 4)
# Nested loop
tests.append(("""
j = 0
k = 0
For _var = 1 To 10
	j = j + _var
	For _other = 1 To 10
		k = k + 1
	Next _other
Next _var
""", {"j" : 55, "k" : 100}))

# Nested loop - break from inner
tests.append(("""
j = 0
k = 0
For _var = 1 To 10
	j = j + _var
	For _other = 1 To 10
		k = k + 1
		Exit For
	Next _other
Next _var
""", {"j" : 55, "k" : 10}))

# Nested loop - break from outer
tests.append(("""
j = 0
k = 0
For _var = 1 To 10
	j = j + _var
	For _other = 1 To 10
		k = k + 1
	Next _other
	Exit For
Next _var
""", {"j" : 1, "k" : 10}))
# << For tests >> (3 of 4)
# Simple for each with collection
tests.append(("""
Dim _c As New Collection
_c.Add 10
_c.Add 20
_c.Add 30
t = 0
For Each _v In _c
	t = t + _v
Next _v
""", {"t" : 60}))

# Nested for each with collection
tests.append(("""
Dim _c As New Collection
_c.Add 10
_c.Add 20
_c.Add 30
Dim _d As New Collection
_d.Add 1
_d.Add 2
_d.Add 3
t = 0
For Each _v In _c
	t = t + _v
	For Each _x In _d
		t = t + _x
	Next _x
Next _v
""", {"t" : 78}))

# Simple for each with variant
tests.append(("""
Dim _c(10)
For _i = 1 To 10
	_c(_i) = _i
Next _i
t = 0
For Each _v In _c
	t = t + _v
Next _v
""", {"t" : 55}))
# << For tests >> (4 of 4)
# Simple test - note we are not interested in the value of the loop variable at the end of the loop
# hence the _var name
tests.append(("""
j = 0
For _var = 0.1 To 1.0
	j = j + _var
Next _var
""", {"j" : 0.1}))

# Simple test with step
tests.append(("""
j = 0
For _var = .1 To 1.0 Step .2
	j = j + _var
Next _var
""", {"j" : 2.5}))

# Simple test with reverse step
tests.append(("""
j = 0
For _var = 1.0 To .1 Step -.1
	j = j + _var
Next _var
j = int(j*10)
""", {"j" : 55}))

# Empty loop
tests.append(("""
j = 0
For _var = 1.0 To .1
	j = j + _var
Next _var
""", {"j" : 0}))

# Breaking out of the loop
tests.append(("""
j = 0
For _var = .1 To 1.0
	j = j + _var
	Exit For
Next _var
""", {"j" : .1}))
# -- end -- << For tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
	main()
