#
# Turn off logging in extensions (too loud!)
import vb2py.extensions
vb2py.extensions.disableLogging()

from unittest import *
from vb2py.vbparser import convertVBtoPython
import vb2py.vbfunctions as vbfunctions
import vb2py.vbfunctions

#
# Import script testing
try:
    import scripttest
except ImportError:
    scripttest = None

#
# Private data hiding may obscure some of the testing so we turn it off
import vb2py.config
Config = vb2py.config.VB2PYConfig()
Config.setLocalOveride("General", "RespectPrivateStatus", "No")
Config.setLocalOveride("General", "ReportPartialConversion", "No")

tests = []

def BasicTest():
    """Return a new class - we do it this way to allow this to work properly for multiple tests"""
    class _BasicTest(TestCase):
        """Holder class which gets built into a whole test case"""
    return _BasicTest

# << Test functions >> (1 of 2)
def getTestMethod(vb, result):
    """Create a test method"""
    def testMethod(self):
        local_dict = {"convertVBtoPython" : convertVBtoPython,
                      "vbfunctions" : vbfunctions}
        # << Parse VB >>
        try:					  
            python = convertVBtoPython(vb.replace("\r\n", "\n"))
        except Exception, err:
            self.fail("Error while parsing (%s)\n%s" % (err, vb))
        # -- end -- << Parse VB >>
        # << Execute the Python code >>
        try:
            exec "from vb2py.vbfunctions import *" in local_dict
            exec python in local_dict
        except Exception, err:
            if not result.has_key("FAIL"):
                self.fail("Error (%s):\n%s\n....\n%s" % (err, vb, python))
        else:
            if result.has_key("FAIL"):
                self.fail("Should have failed:%s\n\n%s" % (vb, python))
        # -- end -- << Execute the Python code >>
        # << Work out what is expected >>
        expected = {}
        exec "" in expected
        expected.update(result)
        expected["convertVBtoPython"] = convertVBtoPython
        expected["vbfunctions"] = vbfunctions

        # Variables which we don't worry about
        skip_variables = ["vbclasses", "logging", "logger"]
        # -- end -- << Work out what is expected >>
        # << Check for discrepancies >>
        reason = ""
        for key in local_dict:
            if not (key.startswith("_") or hasattr(vb2py.vbfunctions, key)):
                try:
                    if expected[key] <> local_dict[key]:
                        reason += "%s (exp, act): '%s' <> '%s'\n" % (key, expected[key], local_dict[key])
                except KeyError:
                    if key not in skip_variables:
                        reason += "Variable didn't exist: '%s'\n" % key
        # -- end -- << Check for discrepancies >>
        #
        self.assert_(reason == "", "Failed: %s\n%s\n\n%s" % (reason, vb, python))

    return testMethod
# << Test functions >> (2 of 2)
def getScriptTestMethod(vb):
    """Create a test method using the script testing method"""
    if scripttest is None:
        raise ImportError("Could not import script test - must be run on Windows with win32com support")
    def testMethod(self):
        scripttest.testCode(vb)
    return testMethod
# -- end -- << Test functions >>

#
# Add tests to main test class
def addTestsTo(TestClassFactory, tests):
    """Add all the tests to the test class"""
    TestClass = TestClassFactory()
    for idx in range(len(tests)):
        setattr(TestClass, "test%d" % idx, getTestMethod(*tests[idx]))
    return TestClass

#
# Add tests to main test class using script testing
def addScriptTestsTo(TestClassFactory, tests):
    """Add all the tests to the test class"""
    TestClass = TestClassFactory()
    for idx in range(len(tests)):
        setattr(TestClass, "test%d" % idx, getScriptTestMethod(tests[idx]))
    return TestClass
