#
# Turn off logging in extensions (too loud!)
import vb2py.extensions
vb2py.extensions.disableLogging()

from unittest import *
from vb2py.vbparser import buildParseTree, VBParserError

#
# Set some config options which are appropriate for testing
import vb2py.config
Config = vb2py.config.VB2PYConfig()
Config.setLocalOveride("General", "ReportPartialConversion", "No")


tests = []

# << Parsing tests >>
# Simple for
tests.append("""
For i = 0 To 1000
  a = a + 1
Next i
""")

# Simple for
tests.append("""
For i=0 To 1000
  a = a + 1
Next i
""")

# Empty for
tests.append("""
For i = 0 To 1000
Next i
""")

# Simple for with unnamed Next
tests.append("""
For i = 0 To 1000
  a = a + 1
Next
""")

# For with step
tests.append("""
For i = 0 To 1000 Step 2
  a = a + 1
Next i
""")

# For with exit
tests.append("""
For i = 0 To 1000
  a = a + 1
  Exit For
Next i
""")

# Nested for
tests.append("""
For i = 0 To 1000
  a = a + 1
  For j = 1 To i
     b = b + j
  Next j
Next i
""")

# Dotted names - what does this even mean?
tests.append("""
For me.you = 0 To 1000 Step 2
  a = a + 1
Next me.you
""")
# -- end -- << Parsing tests >>

class ParsingTest(TestCase):
    """Holder class which gets built into a whole test case"""


def getTestMethod(vb):
    """Create a test method"""
    def testMethod(self):
        try:
            buildParseTree(vb)
        except VBParserError:
            raise "Unable to parse ...\n%s" % vb
    return testMethod

#
# Add tests to main test class
for idx in range(len(tests)):
    setattr(ParsingTest, "test%d" % idx, getTestMethod(tests[idx]))


if __name__ == "__main__":
    main()
