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
# open statements
tests.extend([
    "Open fn For Output As 12",
    "Open fn For Output As #12",
    "Open fn For Input As 12",
    "Open fn For Input As #12",
    "Open fn.gk.gl() For Input As #NxtChn()",
    "Open fn For Append Lock Write As 23",
    "Open fn For Random As 23 Len = 1234",
    "Close 1",
    "Close #1",
    "Close channel",
    "Close #channel",
    "Close",
    "Close\na=1",
    "Closet = 10",
])


# Bug #810968 Close #1, #2 ' fails to parse 
tests.extend([
    "Close #1, #2, #3, #4",
    "Close 1, 2, 3, 4",
    "Close #1, 2, #3, 4",
    "Close #one, #two, #three, #four",
    "Close one, two, three, four",
    "Close #1,#2,#3,#4",
    "Close   #1   ,   #2   ,   #3   ,   #4   ",
])
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
