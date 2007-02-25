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
# Simple expressions
tests.extend([
'a = 10',
'a = 20+30',
'a = "hello there"',
'a = 10',
'a = Array(10,20)',
'a = myfunction.mymethod(10)',
'a = &HFF',
'a = &HFF&',
'a = #1/10/2000#',
'a = #1/10#',
'a = 10 Mod 2',
])


# Nested expressions
tests.extend(["a = 10+(10+(20+(30+40)))",
              "a = (10+20)+(30+40)",
              "a = ((10+20)+(30+40))",
])

# Conditional expressions
tests.extend(["a = a = 1",
              "a = a <> 10",
              "a = a > 10",
              "a = a < 10",
              "a = a <= 10",
              "a = a >= 10",
              "a = a = 1 And b = 2",
              "a = a = 1 Or b = 2",
              "a = a Or b",
              "a = a Or Not b",
              "a = Not a = 1",
              "a = Not a",
              "a = a Xor b",
              "a = b Is Nothing",
              "a = b \ 2",
              "a = b Like c",
              'a = "hello" Like "goodbye"',
])

# Things that failed
tests.extend([
            "a = -(x*x)",
            "a = -x*10",
            "a = 10 Mod 6",
            "Set NewEnum = mCol.[_NewEnum]",
            "a = 10 ^ -bob",
])

# Functions
tests.extend([
            "a = myfunction",
            "a = myfunction()",
            "a = myfunction(1,2,3,4)",
            "a = myfunction(1,2,3,z:=4)",
            "a = myfunction(x:=1,y:=2,z:=4)",
            "a = myfunction(b(10))",
            "a = myfunction(b _\n(10))",
])

# String Functions
tests.extend([
            'a = Trim$("hello")',
            'a = Left$("hello", 4)',
])			

# Things that failed
tests.extend([
            "a = -(x*x)",
            "a = -x*10",
            "a = 10 Mod 6",
])

# Address of
tests.extend([
        "a = fn(AddressOf fn)",
        "a = fn(a, b, c, AddressOf fn)",
        "a = fn(a, AddressOf b, AddressOf c, AddressOf fn)",
        "a = fn(a, AddressOf b.m.m, AddressOf c.k.l, AddressOf fn)",
])

# Type of
tests.extend([
        "a = fn(TypeOf fn)",
        "a = fn(a, b, c, TypeOf fn)",
        "a = fn(a, TypeOf b, TypeOf c, TypeOf fn)",
        "a = fn(a, TypeOf b.m.m, TypeOf c.k.l, TypeOf fn)",
        "a = TypeOf Control Is This",
        "a = TypeOf Control Is This Or TypeOf Control Is That",])
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
