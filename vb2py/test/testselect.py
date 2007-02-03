# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

from unittest import *
from testframework import *

# << Select tests >> (1 of 5)
test_string = """
		a = %d
		Select Case a
		Case 1
			b = 10
		Case 2
			b = 20
		Case 3
			b = 30
		Case Else
			b = 40
		End Select
"""

for aval, result in ((1,10), (2,20), (3,30), (4, 40)):
	tests.append((test_string % aval,
				 {"a" : aval, "b" : result}))

tests.append(("""
	a = 10
	Select Case a
	Case 10
	Case 20
	End Select
	""", {"a":10}))
# << Select tests >> (2 of 5)
# These three tests are subtly different and try to catch errors in dealing with inline Case's
test_string = """
		a = %d
		Select Case a
		Case 1 To 9
			b = 0
		Case 10 To 19
			b = 1
		Case 20 To 29
			b = 2
		Case Else
			b = 3
		End Select
"""

for aval in range(1, 40):
	result = aval // 10
	tests.append((test_string % aval,
				 {"a" : aval, "b" : result}))

test_string = """
		a = %d
		Select Case a
		Case 1 To 9:
			b = 0
		Case 10 To 19:
			b = 1
		Case 20 To 29:
			b = 2
		Case Else:
			b = 3
		End Select
"""

for aval in range(1, 40):
	result = aval // 10
	tests.append((test_string % aval,
				 {"a" : aval, "b" : result}))

test_string = """
		a = %d
		Select Case a
		Case 1 To 9: c=0
			b = 0
		Case 10 To 19: c=1
			b = 1
		Case 20 To 29: c=2
			b = 2
		Case Else: c=3
			b = 3
		End Select
"""

for aval in range(1, 40):
	result = aval // 10
	tests.append((test_string % aval,
				 {"a" : aval, "b" : result, "c" : result}))
# << Select tests >> (3 of 5)
test_string = """
		a = %d
		Select Case a
		Case 1, 5, 10
			b = 10
		Case 2, 20 To 30, 15
			b = 20
		Case 3
			b = 30
		Case Else
			b = 40
		End Select
"""

for aval, result in ((1, 10), (5, 10), (10, 10),
					 (2, 20), (20, 20), (25, 20), (30, 20), (15, 20),
					 (3, 30),
					 (4, 40), (-20, 40), (1000, 40)):
	tests.append((test_string % aval,
				 {"a" : aval, "b" : result}))
# << Select tests >> (4 of 5)
test_string = """
		a = %d
		Select Case a
		Case 1, 5, 10
			b = 10
		Case 2, 20 To 30, 15
			Select Case a
						Case -10 To 25
				  b = 21
			   Case Else
				  b = 22
			End Select
		Case 3
			b = 30
		Case Else
			b = 40
		End Select
"""

for aval, result in ((1, 10), (5, 10), (10, 10),
					 (2, 21), (20, 21), (25, 21), (30, 22), (15, 21),
					 (3, 30),
					 (4, 40), (-20, 40), (1000, 40)):
	tests.append((test_string % aval,
				 {"a" : aval, "b" : result}))
# << Select tests >> (5 of 5)
test_string = """
		a = %d
		Select Case a
		Case 1, 5, 10
			b = 10
		Case 2, 20 To 30, 15
			b = 20
		Case 3
			b = 30
		Case Is > 2000
			b = 35
		Case -2000, -1000, Is < -5000
			b = 36
		Case Else
			b = 40
		End Select
"""

for aval, result in ((1, 10), (5, 10), (10, 10),
					 (2, 20), (20, 20), (25, 20), (30, 20), (15, 20),
					 (3, 30),
					 (2001, 35), (2002, 35), (10000, 35),
					 (-2000, 36), (-1000, 36), (-6000, 36),
					 (1000, 40), (-20, 40), (2000, 40), (-4000, 40)):
	tests.append((test_string % aval,
				 {"a" : aval, "b" : result}))
# -- end -- << Select tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
	main()
