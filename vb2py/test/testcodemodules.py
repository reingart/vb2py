# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

from unittest import *
from complexframework import *


# << CodeModule tests >>
#
# Module level variables should be global (in the Python sense of the word)
tests.append((
		VBCodeModule(),
		"""
		Public my_a As String

		Public Sub SetA(Value As Integer)
			my_a = Value
		End Sub
		Public Function GetA()
			GetA = my_a
		End Function
		""",
		("SetA('hello')\n"
		 "assert GetA() == 'hello', 'GetA was (%s)' % (GetA(),)\n",)
))
# -- end -- << CodeModule tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
	main()
