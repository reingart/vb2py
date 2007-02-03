# Created by Leo from: C:\Development\Python23\Lib\site-packages\vb2py\vb2py.leo

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
# Recordset notation
tests.extend([
"RS!diskID = DriveID",
"RS!diskID = DriveID+10",
'RS!diskID = "DriveID"',
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
