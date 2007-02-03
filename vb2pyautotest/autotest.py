"""An automatic testing framework to compare VB and Python versions of a function"""	

import os
import pprint
from vb2py import converter, vbparser, config
Config = config.VB2PYConfigObject("autotest.ini", "../vb2pyautotest")

# << AutoTest Functions >> (1 of 3)
class TestMaker:
    """A Base Class to help in making unit tests"""

    default_filename = "testscript.txt"

    # << TestMaker methods >> (1 of 8)
    def __init__(self, filename):
        """Initialize the test maker"""
        self.filename = filename
    # << TestMaker methods >> (2 of 8)
    def parseVB(self):
        """Parse the VB code"""
        self.parser =  converter.VBConverter(converter.importTarget("PythonCard"), converter.FileParser)	
        self.parser.doConversion(self.filename)
    # << TestMaker methods >> (3 of 8)
    def createTests(self):
        """Create the test signatures"""
        self.tests = self.extractSignatures(self.parser)
    # << TestMaker methods >> (4 of 8)
    def makeTestForFunction(self, testid, modulename, fnname, paramlist):
        """Make a test script for a function

        The function resides in a module and takes a list of parameters which
        are specified in paramlist. Paramlist also includes a list of values
        to pass as that particular parameter.

        The results of the test are written to a file testid

        """
        raise NotImplementedError
    # << TestMaker methods >> (5 of 8)
    def extractSignatures(self, project):
        """Extract function test signatures from a project"""
        fns = []
        for module in project.modules:
            for defn in module.code_structure.locals:
                if isinstance(defn, vbparser.VBFunction):
                    test = ["test_%s_%s" % (module.name, defn.identifier),
                            module.name,
                            defn.identifier]
                    ranges = []
                    for param in defn.parameters:
                        try:
                            thisrange = Config["DefaultRanges", param.type]
                        except config.ConfigParser.NoOptionError:
                            thisrange = []
                        ranges.append((param.identifier, eval(thisrange)))
                    test.append(ranges)
                    fns.append(test)
        return fns
    # << TestMaker methods >> (6 of 8)
    def createTestScript(self):
        """Create a script containing the tests"""
        ret = []
        for test in self.tests:
            ret.append(self.makeTestForFunction(*test))
        return "\n".join(ret)
    # << TestMaker methods >> (7 of 8)
    def writeTestsToFile(self, filename=None):
        """Write the test script to a file"""
        script = self.createTestScript()
        if filename is None:
            filename = self.default_filename
        f = open(filename, "w")
        f.write(script)
        f.close()
    # << TestMaker methods >> (8 of 8)
    def makeTestFile(self, filename=None):
        """Translate VB and make tests"""
        self.parseVB()
        self.createTests()
        self.writeTestsToFile(filename)
    # -- end -- << TestMaker methods >>
# << AutoTest Functions >> (2 of 3)
class PythonTestMaker(TestMaker):
    """A Class to help in making Python unit tests"""

    default_filename = "testscript.py"

    # << PythonTestMaker methods >>
    def makeTestForFunction(self, testid, modulename, fnname, paramlist):
        """Make a Python test script for a function

        The function resides in a module and takes a list of parameters which
        are specified in paramlist. Paramlist also includes a list of values
        to pass as that particular parameter.

        The results of the test are written to a file testid

        """
        ret = []
        ret.append("from %s import %s" % (modulename, fnname))
        ret.append("results = []")
        tabs = ""
        #
        for param, values in paramlist:
            ret.append("%sfor %s in %s:" % (tabs, param, values))
            tabs += "\t"
        #
        arg_list = ",".join([param[0] for param in paramlist])
        ret.append("%sresults.append((%s(%s), %s))" % (tabs, fnname, arg_list, arg_list))
        ret.extend(("f = open('%s_py.txt', 'w')" % testid,
                   r"f.write('# vb2Py Autotest results\n')",
                   r"f.write('\n'.join([', '.join(map(str, x)) for x in results]))",
                   r"f.close()"))											 
        #
        return "\n".join(ret)
    # -- end -- << PythonTestMaker methods >>
# << AutoTest Functions >> (3 of 3)
class VBTestMaker(TestMaker):
    """A Class to help in making VB unit tests"""

    default_filename = "testscript.bas"

    # << VBTestMaker methods >>
    def makeTestForFunction(self, testid, modulename, fnname, paramlist):
        """Make a VB test script for a function

        The function resides in a module and takes a list of parameters which
        are specified in paramlist. Paramlist also includes a list of values
        to pass as that particular parameter.

        The results of the test are written to a file testid

        """
        ret = []
        ret.append("Dim Results()")
        ret.append("ReDim Results(0)")
        tabs = ""
        #
        for param, values in paramlist:
            ret.append("%sfor each %s in Array%s" % (tabs, param, tuple(values)))
            tabs += "\t"
        #
        arg_list = ",".join([param[0] for param in paramlist])
        ret.append("%sRedim Preserve Results(UBound(Results)+1)" % tabs)
        ret.append("%sAnswer = %s(%s)" % (tabs, fnname, arg_list))
        #
        result_list = ' & "," & '.join(["Str(%s)" % x for x in ["Answer"] + [param[0] for param in paramlist]])
        ret.append("%sResults(Ubound(Results)) = %s" % (tabs, result_list))
        #
        for param, values in paramlist:
            tabs = tabs[:-1]
            ret.append("Next %s" % param)
        #
        ret.extend(("Chn = NextFile",
                    "Open '%s_vb.txt' For Output As #Chn" % testid,
                    'Print #Chn, "# vb2Py Autotest results"',
                    "For Each X In Results",
                    "    Print #Chn, X",
                    "Next X",
                    "Close #Chn"))											 
        #
        return "\n".join(ret)
    # -- end -- << VBTestMaker methods >>
# -- end -- << AutoTest Functions >>

if __name__ == "__main__":
    filename = 'c:\\development\\python22\\lib\\site-packages\\vb2py\\vb\\test3\\Globals.bas'
    p = PythonTestMaker(filename)
    v = VBTestMaker(filename)
