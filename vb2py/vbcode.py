# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

# << Documentation >>
#	"""
#	This module implements a set of classes for describing VB code. The 
#	structure is,
#	
#	- namespace
#	- module
#	- function
#	- block
#	- line
#	- variable
#	
#	- modules, form modules and classes all inherit from BaseModule
#	- subs and functions all inherit from BaseFunction
#	- For, While, Do, Select blocks all inherit from BaseBlock
#	
#	"""


# -- end -- << Documentation >>
# << Declarations >>
import re
from Plex import *
from StringIO import StringIO
# -- end -- << Declarations >>
# << VBCode classes >> (1 of 3)
class BaseVariable(object):
	"""A representation of a VB object"""

	# << BaseVariable methods >>
	def __init__(self, name, vartype=""):
		"""Initialize the variable"""
		if vartype == "":
			vartype = "Variant"
		self.name = name
		self.vartype = vartype
	# -- end -- << BaseVariable methods >>
# << VBCode classes >> (2 of 3)
class BaseNameSpace(object):
	"""A representation of a VB namespace"""

	# << BaseNameSpace declarations >>
	letter = Range("AZaz")
	digit = Range("09")
	name = letter + Rep(letter | digit)
	number = Rep1(digit)
	space = Any(" \t\n")
	delim = Str(",") + Opt(space)

	scope = Str("Dim") | Str("Private") | Str("Public")
	const = Str("Const")

	array = Str("(") + Rep(name + Rep(delim + name))  + Str(")")	
	var_with_type = name + Opt(array) + space + Opt(Str("As") + space + name)
	var_with_no_type = name + Opt(array)

	var = var_with_type | var_with_no_type	
	declare = scope + space + var + Rep(delim + var_with_type) + Eol

	lexicon = Lexicon([
		(declare, "declare"),
		(Str("\n"), ""),
	])

	sub_lex = Lexicon([
		(scope, "scope"),
		(var_with_type, "type"),
		(delim, ""),
		(space, ""),
		(const, "const"),
		(Eol, ""),
	])
	# -- end -- << BaseNameSpace declarations >>
	# << BaseNameSpace methods >> (1 of 2)
	def __init__(self, name="Namespace"):
		"""Initialize the module"""
		self.public_names = []
		self.private_names = []
	# << BaseNameSpace methods >> (2 of 2)
	def initVariablesFromText(self, text):
		"""Initialize the variables from some text"""
		#
		# Declares
		s = StringIO(text)
		scan = Scanner(self.lexicon, s)
		while 1:
			tok = scan.read()
			print tok
			if tok[0] == "declare":
				v = StringIO(tok[1])
				inner_scan = Scanner(self.sub_lex, v)
				while 1:
					inner_tok = inner_scan.read()
					print inner_tok
					if inner_tok[0] is None:
						break
			if tok[0] is None:
				break
	# -- end -- << BaseNameSpace methods >>
# << VBCode classes >> (3 of 3)
class BaseModule(BaseNameSpace):
	"""A representation of a VB module"""

	# << BaseModule methods >> (1 of 3)
	def __init__(self, *args, **kw):
		"""Initialize the module"""
		super(BaseModule, self).__init__(*args, **kw)
		self.functions = []
	# << BaseModule methods >> (2 of 3)
	def initFromText(self, text):
		"""Initialize the structure from some text"""
		self.initVariablesFromText(text)
		self.initFunctionsFromText(text)
	# << BaseModule methods >> (3 of 3)
	def initFunctionsFromText(self, text):
		"""Initialize the functions from some text"""
	# -- end -- << BaseModule methods >>
# -- end -- << VBCode classes >>

if __name__ == "__main__":
	text = open("vb\\test1\\test.bas", "r").read()
	m = BaseModule()
	m.initFromText(text)
	for v in m.public_names:
		print "Public ", v.name, v.vartype
	for v in m.private_names:
		print "Private ", v.name, v.vartype
