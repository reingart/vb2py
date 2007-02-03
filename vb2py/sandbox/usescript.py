"""A test of using the Windows scripting host to compare the VB and converted Python code"""

import win32com.client
import re

# Get the host
vbhost = win32com.client.Dispatch("ScriptControl")
vbhost.language = "vbscript"

import vb2py.vbparser

code = """
' VB2PY-Test: y = [1,2,10,-20,50,0, "blah"]

Function Squared(x)
Squared = x^2
End Function

a = Squared(y)

"""

code1 = """
a=1
b=2
c=a+b

"""

# << Error Classes >>
class ScriptTestError(Exception): """Base class for all script testing errors"""
class ValueNotEqual(ScriptTestError): """The VB and Python result variables didn't match"""
class VBFailed(ScriptTestError): """The VB exec of the code failed but the Python version didn't"""
class PythonFailed(ScriptTestError): """The Python exec of the code failed but the VB version didn't"""
# -- end -- << Error Classes >>
# << Functions >> (1 of 4)
def runTestCode(code, host):
    """Run some code in VB"""
    host.addcode(code)
# << Functions >> (2 of 4)
def findVars(code):
    """Find all the testing variables in a sample of code"""
    finder = re.compile("^(\w+)\s*=.*$", re.MULTILINE)
    return finder.findall(code)
# << Functions >> (3 of 4)
def findDirectives(code):
    """Find testing directives"""
    finder = re.compile("^'\s*VB2PY-Test\s*:\s*(\w+)\s*=\s*(.*)$", re.MULTILINE)
    return finder.findall(code)
# << Functions >> (4 of 4)
def testCode(code, verbose=0):
    """Test some code using the script host"""

    test = findDirectives(code)
    if not test:
        test = [("_dummy", "[1]")]

    for var, value_code in test:		
        #
        values = eval(value_code)
        for value in values:
            #
            if var <> "_dummy":
                line = "%s = %s" % (var, repr(value).replace("'", '"'))
                if verbose:
                    print "Doing '%s'" % line
                #
                vbhost.addcode(line)

            # Run VB	
            try:
                runTestCode(code, vbhost)
            except Exception, vberr:
                vbcode_failed = 1
            else:
                vbcode_failed = 0

            # Run Python
            python = vb2py.vbparser.convertVBtoPython(code)
            vars = findVars(python)
            namespace = {var : value}
            try:
                exec python in namespace
            except Exception, pythonerr:
                pythoncode_failed = 1
            else:
                pythoncode_failed = 0

            # Check for failures
            if vbcode_failed and not pythoncode_failed:
                raise VBFailed("VB failed on '%s' but Python didn't (%s)" % (line, vberr))
            elif pythoncode_failed and not vbcode_failed:
                raise PythonFailed("Python failed on '%s' but VB didn't (%s)" % (line, pythonerr))
            elif pythoncode_failed and vbcode_failed:
                if verbose:
                    print "Both failed VB\n%s\nPython\n%s" % (vberr, pythonerr)
            else:				
                if verbose:
                    print "\tVB\tPython"
                for name in vars:
                    if verbose:
                        print "%s\t%s\t%s" % (name, vbhost.eval(name), namespace[name])	
                    else:
                        if vbhost.eval(name) <> namespace[name]:
                            raise ValueNotEqual("%s: VB(%s), Python(%s)" % (name, vbhost.eval(name), namespace[name]))

    if verbose:
        # Compare
        print "VB Code\n%s\n\n" % code
        print "Python Code\n%s\n\n" % python
# -- end -- << Functions >>

testCode(code, verbose=1)			
print "Finished"
