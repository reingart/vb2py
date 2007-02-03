# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

"""Test some code using the windows scripting host"""

try:
	import win32com.client
except ImportError:
	raise ImportError("Script test requires access to the Windows Scripting host")

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
# << Functions >> (1 of 5)
def runTestCode(code, host):
	"""Run some code in VB"""
	host.addcode(code)
# << Functions >> (2 of 5)
def findVars(code):
	"""Find all the testing variables in a sample of code"""
	finder = re.compile("^(\w+)\s*=.*$", re.MULTILINE)
	return finder.findall(code)
# << Functions >> (3 of 5)
def findDirectives(code):
	"""Find testing directives"""
	finder = re.compile("^'\s*VB2PY-Test\s*:\s*(\w+)\s*=\s*(.*)$", re.MULTILINE)
	return finder.findall(code)
# << Functions >> (4 of 5)
def sendOutput(text, lines, verbose):
	"""Send some output"""
	if verbose == 1:
		print text
	lines.append(text)
# << Functions >> (5 of 5)
def testCode(code, verbose=0):
	"""Test some code using the script host"""

	output = []
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
					sendOutput("Doing '%s'" % line, output, verbose)
				#
				vbhost.addcode(line)
			else:
				line = ""

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
				raise VBFailed("VB failed on '%s' but Python didn't (%s)\n%s" % (line, vberr, python))
			elif pythoncode_failed and not vbcode_failed:
				raise PythonFailed("Python failed on '%s' but VB didn't (%s)\n%s" % (line, pythonerr, python))
			elif pythoncode_failed and vbcode_failed:
				if verbose:
					sendOutput("Both failed\n\nVB\n%s\n\nPython\n%s\n%s" % (vberr, pythonerr, python), output, verbose)
			else:				
				if verbose:
					sendOutput("\tVB\tPython", output, verbose)
				for name in vars:
					if not name.endswith("_"):
						if verbose:
							sendOutput("%s\t%s\t%s" % (name, repr(vbhost.eval(name)), repr(namespace[name])), output, verbose)
						else:
							if vbhost.eval(name) <> namespace[name]:
								raise ValueNotEqual("%s: VB(%s), Python(%s)\n%s" % (
										name, repr(vbhost.eval(name)), repr(namespace[name]), python))

	if verbose==1:
		# Compare
		print "VB Code\n%s\n\n" % code
		print "Python Code\n%s\n\n" % python

	return output
# -- end -- << Functions >>

if __name__ == "__main__":
	testCode(code, verbose=1)			
	print "Finished"
