# Created by Leo from: C:\Development\Python23\Lib\site-packages\vb2py\vb2py.leo

#
# Turn off logging in extensions (too loud!)
import vb2py.extensions
vb2py.extensions.disableLogging()

from unittest import *
from vb2py.vbparser import convertVBtoPython, VBClassModule, VBModule, VBFormModule, VBCodeModule
import vb2py.vbfunctions as vbfunctions
import vb2py.vbfunctions

tests = []

def BasicTest():
    """Return a new class - we do it this way to allow this to work properly for multiple tests"""
    class _BasicTest(TestCase):
        """Holder class which gets built into a whole test case"""
    return _BasicTest


def getTestMethod(container, vb, assertions):
    """Create a test method"""
    def testMethod(self):
        local_dict = {"convertVBtoPython" : convertVBtoPython,
                      "vbfunctions" : vbfunctions}
        # << Parse VB >>
        try:					  
            python = convertVBtoPython(vb, container=container)
        except Exception, err:
            self.fail("Error while parsing (%s)\n%s" % (err, vb))
        # -- end -- << Parse VB >>
        # << Execute the Python code >>
        try:
            exec "from vb2py.vbfunctions import *" in local_dict
            exec python in local_dict
        except Exception, err:
            self.fail("Error (%s):\n%s\n....\n%s" % (err, vb, python))
        # -- end -- << Execute the Python code >>
        # << Check assertions >>
        #	        Go through each assertion (a Python statement) to see if it holds
        
        reason = ""

        internal_dict = {"python" : python, "vb" : vb}

        for assertion in assertions:
            if assertion.startswith("$"):
                dct = internal_dict
                assertion = assertion[1:]
            else:
                dct = local_dict
            try:
                exec assertion in dct
            except Exception, err:
                reason += "%s (%s)\n" % (Exception, err)
        # -- end -- << Check assertions >>
        #print vb, "\n\n", python, "\n\n--------------------------------"
        #
        self.assert_(reason == "", "Failed: %s\n%s\n\n%s" % (reason, vb, python))

    return testMethod

#
# Add tests to main test class
def addTestsTo(TestClassFactory, tests):
    """Add all the tests to the test class"""
    TestClass = TestClassFactory()
    for idx in range(len(tests)):
        setattr(TestClass, "test%d" % idx, getTestMethod(*tests[idx]))
    return TestClass
