# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

# << Imports >>
#
# Configuration options
import config
Config = config.VB2PYConfig()

from pprint import pprint as pp
from simpleparse.common import chartypes
import sys, os

declaration = open(os.path.join(sys.exec_prefix, "lib", "site-packages", "vb2py", "vbgrammar.txt"), "r").read()

from simpleparse import generator
try:
	from TextTools import TextTools
except ImportError:
	from mx.TextTools import TextTools


import logging   # For logging output and debugging 
logging.basicConfig() # Console output for logging

log = logging.getLogger("VBParser")
log.setLevel(int(Config["General", "LoggingLevel"]))
# -- end -- << Imports >>
# << Error Classes >>
class VBParserError(Exception): 
	"""An error occured during parsing"""

class UnhandledStructureError(VBParserError): 
	"""A structure was parsed but could not be handled by class"""
class InvalidOption(VBParserError): 
	"""An invalid config option was detected"""
class NestingError(VBParserError): 
	"""An error occured while handling a nested structure"""
class UnresolvableName(VBParserError):
	"""We were asked to resolve a name but couldn't because we don't know it"""

class SystemPluginFailure(VBParserError): 
	"""A system level plugin failed"""

class DirectiveError(VBParserError): 
	"""An unknown directive was found"""
# -- end -- << Error Classes >>
# << Definitions >>
pass
# -- end -- << Definitions >>

# << Utility functions >> (1 of 7)
def renderListAsCode(lst):
	for item in lst:
		print item.renderAsCode()
# << Utility functions >> (2 of 7)
def convertToElements(details, txt):
	"""Convert a parse tree to elements"""
	ret = []
	if details:
		for item in details:
			ret.append(VBElement(item, txt))
	return ret
# << Utility functions >> (3 of 7)
def indentList(lst, indent):
	"""Join and indent a list of strings"""
	raise NotImplementedError("This function is deprecated")
	joiner = " "*indent*INDENT_AMOUNT
	return joiner.join(lst)
# << Utility functions >> (4 of 7)
def buildParseTree(vbtext, starttoken="line", verbose=0):
	"""Parse some VB"""

	#
	# Build a parser
	parser = generator.buildParser(declaration).parserbyname(starttoken)

	txt = applyPlugins("preProcessVBText", vbtext)

	nodes = []
	while 1:
		success, tree, next = TextTools.tag(txt, parser)
		if not success:
			if txt.strip():
				raise VBParserError("Failed, %d, '%s'" % (next, txt.split("\n")[0]))
			break
		if verbose:
			print success, next
			pp(tree)
			print "."
		nodes.extend(convertToElements(tree, txt))
		txt = txt[next:]

	return nodes
# << Utility functions >> (5 of 7)
def parseVB(vbtext, container=None, starttoken="line", verbose=0):
	"""Parse some VB"""

	nodes = buildParseTree(vbtext, starttoken, verbose)

	if container is None:
		m = VBModule()
	else:
		m = container

	for idx, node in zip(xrange(sys.maxint), nodes):
		if verbose:
			print idx,
		try:
			m.processElement(node)
		except UnhandledStructureError:
			log.warn("Unhandled: %s\n%s" % (node.structure_name, node.text))

	return m
# << Utility functions >> (6 of 7)
def convertVBtoPython(vbtext, *args, **kw):
	"""Convert some VB text to Python"""
	m = parseVB(vbtext, *args, **kw)
	return applyPlugins("postProcessPythonText", m.renderAsCode())
# << Utility functions >> (7 of 7)
def applyPlugins(methodname, txt):
	"""Apply the method of all active plugins to this text"""
	use_user_plugins = Config["General", "LoadUserPlugins"] == "Yes"
	for plugin in plugins:
		if plugin.isEnabled() and plugin.system_plugin or use_user_plugins:
			try:
				txt = getattr(plugin, methodname)(txt)	
			except Exception, err:
				if plugin.system_plugin:
					raise SystemPluginFailure(
						"System plugin '%s' had an exception (%s) while doing %s. Unable to continue" % (
							plugin.name, err, methodname))
				else:                        
					log.warn("Plugin '%s' had an exception (%s) while doing %s and will be disabled" % (
							plugin.name, err, methodname))
					plugin.disable()
	return txt
# -- end -- << Utility functions >>
# << Classes >> (1 of 53)
class VBElement(object):
	"""An element of VB code"""

	# << VBElement methods >> (1 of 2)
	def __init__(self, details, text):
		"""Initialize from the details"""
		self.name = details[0]
		self.text = text[details[1]:details[2]]
		self.elements = convertToElements(details[3], text)
	# << VBElement methods >> (2 of 2)
	def printTree(self, offset=0):
		"""Print out this tree"""
		print "%s%s : '%s'" % (" "*offset, self.name, self.text.split("\n")[:20])
		for subelement in self.elements:
			subelement.printTree(offset+1)
	# -- end -- << VBElement methods >>
# << Classes >> (2 of 53)
class VBNamespace(object):
	"""Handles a VB Namespace"""

	# << VBNamespace declarations >>
	#
	# Autohandlers is a list of token names. These will be extracted and the values added as
	# attributes to our object
	auto_handlers = []
	auto_class_handlers = None

	#
	# Skip handlers are automatically by-passed. This is useful for quickly ignoring a 
	# handler in a base class
	skip_handlers = []

	#
	# Used to translate () into [] under certain circumstances (LHS of an assign)
	brackets_are_indexes = 0


	default_scope = "Private"
	# -- end -- << VBNamespace declarations >>
	# << VBNamespace methods >> (1 of 18)
	def __init__(self, scope="Private"):
		"""Initialize the namespace"""
		self.locals = []
		self.local_default_scope = self.default_scope
		self.auto_class_handlers = {
			"object_definition" : (VBVariableDefinition, self.locals),
			"const_statement"   : (VBConstant, self.locals),
			"user_type_definition" : (VBUserType, self.locals),
		}
		#
		# This dictionary stores names which are to be substituted if found 
		self.name_substitution = {}

		# << Get indenting options >>
		char_spec = Config["General", "IndentCharacter"]
		if char_spec == "Space":
			self._indent_char = " "
		elif char_spec == "Tab":
			self._indent_char = "\t"
		else:
			raise InvalidOption("Indent character option not understood: '%s'" % char_spec)

		self._indent_amount = int(Config["General", "IndentAmount"])
		# -- end -- << Get indenting options >>
	# << VBNamespace methods >> (2 of 18)
	def processElement(self, element):
		"""Process our tree"""
		handler = self.getHandler(element)
		if handler:
			handler(element)
		else:
			if element.elements:
				for subelement in element.elements:
					self.processElement(subelement)
			else:
				log.info("Unhandled element '%s' from %s\n%s" % (element.name, self, element.text))
	# << VBNamespace methods >> (3 of 18)
	def getHandler(self, element):
		"""Find a handler for the element"""
		if element.name in self.skip_handlers:
			return None
		elif element.name in self.auto_handlers:
			log.info("'%s' found auto handler for '%s'" % (self, element.name))
			return self.createExtractHandler(element.name)
		elif element.name in self.auto_class_handlers:
			log.info("'%s' found auto class handler for '%s'" % (self, element.name))
			# << Create class handler >>
			obj_class, add_to = self.auto_class_handlers[element.name]

			if obj_class == self.__class__:
				# Ooops, recursive handling - we should handle the sub elements
				def class_handler(element):
					for sub_element in element.elements:
						self.handleSubObject(sub_element, obj_class, add_to)
			else:	
				def class_handler(element):
					self.handleSubObject(element, obj_class, add_to)

			return class_handler
			# -- end -- << Create class handler >>
		try:
			return getattr(self, "handle_%s" % element.name)
		except AttributeError:
			return None
	# << VBNamespace methods >> (4 of 18)
	def createExtractHandler(self, token):
		"""Create a handler which will extract a certain token value"""
		def handler(element):
			log.info("Grabbed attribute '%s' for %s as '%s'" % (token, self, element.text))
			setattr(self, token, element.text)
		return handler
	# << VBNamespace methods >> (5 of 18)
	def asString(self):
		"""Convert to a nice representation"""
		return repr(self)
	# << VBNamespace methods >> (6 of 18)
	def handleSubObject(self, element, obj_class, add_to):
		"""Handle an object which creates a sub object"""
		v = obj_class(self.local_default_scope)
		v.processElement(element)
		v.parent = self
		#
		# Assume that we are supposed to add this to a list of items
		# if this fails then perhaps this is an attribute we are supposed to set
		try:
			add_to.append(v)	
		except AttributeError:
			setattr(self, add_to, v)
		#
		log.info("Added new %s to %s" % (obj_class, self.asString()))
	# << VBNamespace methods >> (7 of 18)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		return self.getIndent(indent) + "# Unrendered object %s\n" % (self.asString(), )
	# << VBNamespace methods >> (8 of 18)
	def getParentProperty(self, name):
		"""Get a property from our nearest ancestor who has it"""
		try:
			return getattr(self, name)
		except AttributeError:
			try:
				parent = self.parent
				return parent.getParentProperty(name)
			except AttributeError:
				raise NestingError("Reached outer level when trying to access a parent property: '%s'" % name)
	# << VBNamespace methods >> (9 of 18)
	def searchParentProperty(self, name):
		"""Search for any ancestor who has the named parameter set to tru"""
		try:
			if getattr(self, name):
				return 1
		except AttributeError:
			pass
		try:
			parent = self.parent
			return parent.searchParentProperty(name)
		except AttributeError:
			return 0
	# << VBNamespace methods >> (10 of 18)
	def getLocalNameFor(self, name):
		"""Get the local version of a name

		We look for any ancestor with a name conversion in operation for this name and
		return the first one that has it. If there are none then we just use the name

		"""
		try:
			return self.name_substitution[name]
		except KeyError:
			try:
				return self.parent.getLocalNameFor(name)
			except AttributeError:
				return name
	# << VBNamespace methods >> (11 of 18)
	def getIndent(self, indent):
		"""Return some spaces to do indenting"""
		return self._indent_char*indent*self._indent_amount
	# << VBNamespace methods >> (12 of 18)
	def getWarning(self, warning_type, text, indent=0, crlf=0):
		"""Construct a warning comment"""
		ret = "%s# %s (%s) %s" % (
				self.getIndent(indent),
				Config["General", "AttentionMarker"],
				warning_type,
				text)
		if crlf:
			ret += "\n"
		return ret
	# << VBNamespace methods >> (13 of 18)
	def checkOptionChoice(self, section, name, choices):
		"""Return the index of a config option in a list of choices

		We return the actual choice name which may seem odd but is done to make
		the code readable. The main purpose of this method is to allow the choice
		to be selected with the error trapping hidden.

		"""
		value = Config[section, name]
		try:
			return choices[list(choices).index(value)]
		except ValueError:
			raise InvalidOption("Invalid option for %s.%s, must be one of %s" % (
										section, name, choices))
	# << VBNamespace methods >> (14 of 18)
	def checkOptionYesNo(self, section, name):
		"""Return the yes/no value of an option checking for invalid answers"""
		return self.checkOptionChoice(section, name, ("Yes", "No"))
	# << VBNamespace methods >> (15 of 18)
	def resolveName(self, name):
		"""Convert a local name to a fully resolved name

		We traverse up through the nested namespaces until someone knows
		what to do with the name. If nobody knows then we know if must be
		a local so it keeps the same name.

		"""
		try:
			return self.resolveLocalName(name)
		except UnresolvableName:
			try:
				return self.parent.resolveName(name)
			except AttributeError:
				return name	# Nobody knew the name so it must be local
	# << VBNamespace methods >> (16 of 18)
	def resolveLocalName(self, name):
		"""Convert a local name to a fully resolved name"""
		raise UnresolvableName("Name '%s' is not known in this namespace" % name)
	# << VBNamespace methods >> (17 of 18)
	def handle_scope(self, element):
		"""Handle a scope definition"""
		self.local_default_scope = element.text
		log.info("Changed default scope to %s" % self.local_default_scope)
	# << VBNamespace methods >> (18 of 18)
	def handle_line_end(self, element):
		"""Handle the end of a line"""
		self.local_default_scope = self.default_scope
	# -- end -- << VBNamespace methods >>
# << Classes >> (3 of 53)
class VBConsumer(VBNamespace):
	"""Consume and store elements"""

	def processElement(self, element):
		"""Eat this element"""
		self.element = element
		log.info("Consumed element: %s" % element)
# << Classes >> (4 of 53)
class VBUnrendered(VBConsumer):
	"""Represents an unrendered statement"""

	def renderAsCode(self, indent):
		"""Render the unrendrable!"""
		show_warning = Config["General", "WarnAboutUnrenderedCode"]
		if show_warning == "Yes":
			return self.getWarning("UntranslatedCode", self.element.text.replace("\n", "\\n"), indent, crlf=1)
		elif show_warning == "No":
			return ""
		else:
			raise InvalidOption("WarnAboutUnrenderedCode option not understood: '%s'" % show_warning)
# << Classes >> (5 of 53)
class VBCodeBlock(VBNamespace):
	"""A block of VB code"""

	# << VBCodeBlock methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the block"""
		super(VBCodeBlock, self).__init__()
		self.blocks = []
		self.auto_class_handlers.update({
			"assignment_statement" : (VBAssignment, self.blocks),
			"set_statement" : (VBSet, self.blocks),
			"comment_body" : (VBComment, self.blocks),
			"vb2py_directive" : (VB2PYDirective, self.blocks),
			"if_statement" : (VBIf, self.blocks),
			"inline_if_statement" : (VBInlineIf, self.blocks),
			"select_statement" : (VBSelect, self.blocks),
			"for_statement" : (VBFor, self.blocks),
			"for_each_statement" : (VBForEach, self.blocks),
			"exit_statement" : (VBExitStatement, self.blocks),
			"while_statement" : (VBWhile, self.blocks),
			"do_statement" : (VBDo, self.blocks),
			"redim_statement" : (VBReDim, self.blocks),
			"implicit_call_statement" : (VBCall, self.blocks),
			"label_statement" : (VBLabel, self.blocks),
			"with_statement" : (VBWith, self.blocks),

			"open_statement" : (VBOpen, self.blocks),
			"close_statement" : (VBClose, self.blocks),
			"input_statement" : (VBInput, self.blocks),
			"print_statement" : (VBPrint, self.blocks),
			"line_input_statement" : (VBLineInput, self.blocks),

			"name_statement" : (VBUnrendered, self.blocks),
			"resume_statement" : (VBUnrendered, self.blocks),
			"goto_statement" : (VBUnrendered, self.blocks),
			"on_statement" : (VBUnrendered, self.blocks),
			"external_declaration" : (VBUnrendered, self.blocks),
			"get_statement" : (VBUnrendered, self.blocks),
			"put_statement" : (VBUnrendered, self.blocks),

		})
	# << VBCodeBlock methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		return "".join([block.renderAsCode(indent) for block in self.blocks])
	# -- end -- << VBCodeBlock methods >>
# << Classes >> (6 of 53)
class VBVariable(VBNamespace):
	"""Handles a VB Variable"""

	# << VBVariable declarations >>
	auto_handlers = [
			"scope",
			"size_definition",
			"type",
			"identifier",
			"value",
			"optional",
			"expression",
			"new_keyword",
			"preserve_keyword",
	]

	skip_handlers = [
			"object_definition",
			"const_statement",
	]
	# -- end -- << VBVariable declarations >>
	# << VBVariable methods >> (1 of 3)
	def __init__(self, scope="Private"):
		"""Initialize the variable"""
		super(VBVariable, self).__init__(scope)
		self.identifier = None
		self.scope = scope
		self.type = "Variant"
		self.size_definition = None
		self.value = None
		self.optional = None
		self.expression = "VBMissingArgument"
		self.new_keyword = None
		self.preserve_keyword = None
	# << VBVariable methods >> (2 of 3)
	def asString(self):
		"""Convert to a nice representation"""
		return "%s\t%s\t%s\t%s\t%s" % (self.type, self.identifier, 
					self.size_definition, self.scope, self.value)
	# << VBVariable methods >> (3 of 3)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		if self.optional:
			return "%s=%s" % (self.identifier, self.expression)
		else:
			return self.identifier
	# -- end -- << VBVariable methods >>
# << Classes >> (7 of 53)
class VBObject(VBNamespace):
	"""Handles a VB Object"""

	am_on_lhs = 0 # Set to 1 if the object is on the LHS of an assignment

	# << VBObject methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the object"""
		super(VBObject, self).__init__(scope)

		self.primary = None
		self.modifiers = []
		self.implicit_object = None

		self.auto_class_handlers.update({
			"primary" : (VBConsumer, "primary"),
			"attribute" : (VBAttribute, self.modifiers),
			"parameter_list" : (VBParameterList, self.modifiers),
		})

		self.auto_handlers = (
			"implicit_object",
		)
	# << VBObject methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""Render this subroutine"""
		#
		# Check for implicit object and if we are one then find the nearest "With"
		if self.implicit_object:
			implicit_name = "%s." % self.getParentProperty("with_object")
		else:
			implicit_name = ""
		#
		# For the LHS objects we need to look for the local name for Function return arguments
		if self.am_on_lhs:
			obj_name = self.getLocalNameFor(self.primary.element.text)
		else:
			obj_name = self.primary.element.text
		#
		return "%s%s%s" % (implicit_name,
						   self.resolveName(obj_name),
						   "".join([item.renderAsCode() for item in self.modifiers]))
	# -- end -- << VBObject methods >>
# << Classes >> (8 of 53)
class VBLHSObject(VBObject):
	"""Handles a VB Object appearing on the LHS of an assignment"""

	am_on_lhs = 1 # Set to 1 if the object is on the LHS of an assignment
# << Classes >> (9 of 53)
class VBAttribute(VBConsumer):
	"""An attribute of an object"""

	def renderAsCode(self, indent=0):
		"""Render this attribute"""
		return ".%s" % self.element.text
# << Classes >> (10 of 53)
class VBParameterList(VBCodeBlock):
	"""An parameter list for an object"""

	# << VBParameterList methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the object"""
		super(VBParameterList, self).__init__(scope)

		self.expressions = []
		self.auto_class_handlers.update({
			"expression" : (VBExpression, self.expressions),
		})
	# << VBParameterList methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""Render this attribute"""
		content = ", ".join([item.renderAsCode() for item in self.expressions])
		if self.searchParentProperty("brackets_are_indexes"):
			return "[%s]" % content		
		else:
			return "(%s)" % content
	# -- end -- << VBParameterList methods >>
# << Classes >> (11 of 53)
class VBExpression(VBNamespace):
	"""Represents an comment"""

	# << VBExpression methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the assignment"""
		super(VBExpression, self).__init__(scope)
		self.parts = []
		self.auto_class_handlers.update({
			"sign"	: (VBExpressionPart, self.parts),
			"pre_not" : (VBExpressionPart, self.parts),
			"par_expression" : (VBParExpression, self.parts),
			"operation" : (VBOperation, self.parts),
			"pre_named_argument" : (VBExpressionPart, self.parts),
		})
	# << VBExpression methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		return " ".join([item.renderAsCode(indent) for item in self.parts])
	# -- end -- << VBExpression methods >>
# << Classes >> (12 of 53)
class VBParExpression(VBNamespace):
	"""A block in an expression"""

	auto_handlers = [
		"l_bracket",
		"r_bracket",
	]

	# << VBParExpression methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize"""
		super(VBParExpression, self).__init__(scope)
		self.parts = []
		self.named_argument = ""
		self.auto_class_handlers.update({
			"integer" : (VBExpressionPart, self.parts),
			"stringliteral" : (VBStringLiteral, self.parts),
			"floatnumber" : (VBExpressionPart, self.parts),
			"longinteger" : (VBExpressionPart, self.parts),
			#"object" : (VBExpressionPart, self.parts),
			"object" : (VBObject, self.parts),
			"par_expression" : (VBParExpression, self.parts),
			"operation" : (VBOperation, self.parts),
			"named_argument" : (VBConsumer, "named_argument"),
		})

		self.l_bracket = self.r_bracket = ""
	# << VBParExpression methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		if self.named_argument:
			arg = "%s=" % self.named_argument.element.text
		else:
			arg = ""
		ascode = " ".join([item.renderAsCode(indent) for item in self.parts])
		return "%s%s%s%s" % (arg, self.l_bracket, ascode, self.r_bracket)
	# -- end -- << VBParExpression methods >>
# << Classes >> (13 of 53)
class VBExpressionPart(VBConsumer):
	"""Part of an expression"""

	# << VBExpressionPart methods >>
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		if self.element.name == "object":
			#
			# Check for implicit object (inside a with)
			if self.element.text.startswith("."):
				return "%s%s" % (self.getParentProperty("with_object"),
								 self.element.text)
		elif self.element.name == "pre_named_argument":
			return "%s=" % (self.element.text.split(":=")[0],)
		elif self.element.name == "pre_not":
			self.element.text = "not"

		return self.element.text
	# -- end -- << VBExpressionPart methods >>
# << Classes >> (14 of 53)
class VBOperation(VBExpressionPart):
	"""An operation in an expression"""

	translation = {
		"&" : "+",
		"=" : "==",
		"\\" : "//",  # TODO: Is this right?
		"Is" : "is",
	}

	# << VBOperation methods >>
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		if self.element.text in self.translation:
			return self.translation[self.element.text]
		else:
			return super(VBOperation, self).renderAsCode(indent)
	# -- end -- << VBOperation methods >>
# << Classes >> (15 of 53)
class VBStringLiteral(VBExpressionPart):
	"""Represents a string literal"""

	# << VBStringLiteral methods >>
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		#
		# Remember to replace the double quotes with single ones
		body = self.element.text[1:-1]
		body = body.replace('""', '"')
		body = body.replace("'", "\\'")
		return "'%s'" % body
	# -- end -- << VBStringLiteral methods >>
# << Classes >> (16 of 53)
class VBModule(VBCodeBlock):
	"""Handles a VB Module"""

	skip_handlers = [
	]

	convert_functions_to_methods = 0  # If this is 1 then local functions will become methods
	indent_all_blocks = 0

	# << VBModule methods >> (1 of 6)
	def __init__(self, scope="Private"):
		"""Initialize the module"""
		super(VBModule, self).__init__(scope)
		self.auto_class_handlers.update({
			"sub_definition" : (VBSubroutine, self.locals),
			"fn_definition" : (VBFunction, self.locals),
			"property_definition" : (VBProperty, self.locals),
			"enumeration_definition" : (VBEnum, self.locals),
		})
		self.local_names = []
	# << VBModule methods >> (2 of 6)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		return "%s\n%s\n%s\n%s" % (
					   self.importStatements(indent),
					   self.renderModuleHeader(indent),
					   self.renderDeclarations(indent+self.indent_all_blocks), 
					   self.renderBlocks(indent+self.indent_all_blocks))
	# << VBModule methods >> (3 of 6)
	def importStatements(self, indent=0):
		"""Render the standard import statements for this block"""
		return "from vb2py.vbfunctions import *"
	# << VBModule methods >> (4 of 6)
	def renderDeclarations(self, indent):
		"""Render the declarations as code

		Most of the rendering is delegated to the individual declaration classes. However,
		we cannot do this with properties since they need to be grouped into a single assignment.
		We do the grouping here and delegate the rendering to them.

		"""
		#
		ret = []
		#
		# Handle non-properties and group properties together
		properties = {}
		for declaration in self.locals:
			# Check for property
			if isinstance(declaration, VBProperty):
				log.info("Collected property '%s', decorator '%s'" % (
							declaration.identifier, declaration.property_decorator_type))
				decorators = properties.setdefault(declaration.identifier, {})
				decorators[declaration.property_decorator_type] = declaration
			else:
				ret.append(declaration.renderAsCode(indent))
		#
		# Now render all the properties
		for property in properties:
			if properties[property]:
					ret.append(properties[property].values()[0].renderPropertyGroup(indent, property, **properties[property]))
		#		
		return "".join(ret)
	# << VBModule methods >> (5 of 6)
	def renderBlocks(self, indent=0):
		"""Render this module's blocks"""
		return "".join([block.renderAsCode(indent) for block in self.blocks])
	# << VBModule methods >> (6 of 6)
	def renderModuleHeader(self, indent=0):
		"""Render a header for the module"""
		return ""
	# -- end -- << VBModule methods >>
# << Classes >> (17 of 53)
class VBClassModule(VBModule):
	"""Handles a VB Class"""

	convert_functions_to_methods = 1  # If this is 1 then local functions will become methods
	indent_all_blocks = 1

	# << VBClassModule methods >> (1 of 2)
	def renderModuleHeader(self, indent=0, classname="MyClass"):
		"""Render this element as code"""
		if self.checkOptionYesNo("Classes", "UseNewStyleClasses") == "Yes":
			return "class %s(Object):\n" % classname
		else:
			return "class %s:\n" % classname
	# << VBClassModule methods >> (2 of 2)
	def resolveLocalName(self, name):
		"""Convert a local name to a fully resolved name

		We search our local variables to see if we know the name. If we do then we
		need to add a self.

		"""
		if name in self.local_names:
			return "self.%s" % name
		for obj in self.locals:
			if obj.identifier == name:
				return "self.%s" % name
		raise UnresolvableName("Name '%s' is not known in this namespace" % name)
	# -- end -- << VBClassModule methods >>
# << Classes >> (18 of 53)
class VBCodeModule(VBModule):
	"""Handles a VB Code module"""
# << Classes >> (19 of 53)
class VBFormModule(VBClassModule):
	"""Handles a VB Form module"""

	convert_functions_to_methods = 1  # If this is 1 then local functions will become methods
# << Classes >> (20 of 53)
class VBVariableDefinition(VBVariable):
	"""Handles a VB Dim of a Variable"""

	# << VBVariableDefinition methods >>
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		local_name = self.identifier
		if self.size_definition:
			if self.preserve_keyword:
				preserve = ", %s" % (self.identifier, )
			else:
				preserve = ""
			return "%s%s = vbObjectInitialize(%s, %s%s)\n" % (
							self.getIndent(indent),
							local_name,
							self.size_definition,
							self.type,
							preserve)
		elif self.new_keyword:
			return "%s%s = %s()\n" % (
							self.getIndent(indent),
							local_name,
							self.type)
		else:
			return "%s%s = %s()\n" % (
							self.getIndent(indent),
							local_name,
							self.type)
	# -- end -- << VBVariableDefinition methods >>
# << Classes >> (21 of 53)
class VBConstant(VBVariable):
	"""Represents a constant in VB"""
# << Classes >> (22 of 53)
class VBReDim(VBCodeBlock):
	"""Represents a Redim statement"""

	# << VBReDim methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the Redim"""
		super(VBReDim, self).__init__(scope)
		#
		self.variables = []
		self.preserve = None
		#
		self.auto_class_handlers = {
			"object_definition" : (VBVariableDefinition, self.variables),
			"preserve_keyword" : (VBConsumer, "preserve"),
		}
	# << VBReDim methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		for var in self.variables:
			var.preserve_keyword = self.preserve
		return "".join([var.renderAsCode(indent) for var in self.variables])
	# -- end -- << VBReDim methods >>
# << Classes >> (23 of 53)
class VBAssignment(VBNamespace):
	"""An assignment statement"""

	auto_handlers = [
	]

	# << VBAssignment methods >> (1 of 3)
	def __init__(self, scope="Private"):
		"""Initialize the assignment"""
		super(VBAssignment, self).__init__(scope)
		self.parts = []
		self.object = None
		self.auto_class_handlers.update({
			"expression" : (VBExpression, self.parts),
			"object" : (VBLHSObject, "object")
		})
	# << VBAssignment methods >> (2 of 3)
	def asString(self):
		"""Convert to a nice representation"""
		return "%s = %s" % (self.object, self.parts)
	# << VBAssignment methods >> (3 of 3)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		self.object.brackets_are_indexes = 1 # Convert brackets on LHS to []
		return "%s%s = %s\n" % (self.getIndent(indent),
								self.object.renderAsCode(), 
								self.parts[0].renderAsCode(indent))
	# -- end -- << VBAssignment methods >>
# << Classes >> (24 of 53)
class VBSet(VBAssignment):
	"""A set statement"""

	auto_handlers = [
		"new_keyword",
	]

	new_keyword = ""

	# << VBSet methods >>
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		if not self.new_keyword:
			return super(VBSet, self).renderAsCode(indent)
		else:
			return "%s%s = %s()\n" % (
						self.getIndent(indent),
						self.object.renderAsCode(), 
						self.parts[0].renderAsCode(indent))
	# -- end -- << VBSet methods >>
# << Classes >> (25 of 53)
class VBCall(VBCodeBlock):
	"""A set statement"""

	auto_handlers = [
	]


	# << VBCall methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the assignment"""
		super(VBCall, self).__init__(scope)
		self.parameters = []
		self.object = None
		self.auto_class_handlers = ({
			"expression" : (VBParExpression, self.parameters),
			"object" : (VBObject, "object")
		})
	# << VBCall methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		if self.parameters:
			params = ", ".join([par.renderAsCode() for par in self.parameters])
		else:
			params = ""
		#
		return "%s%s(%s)\n" % (self.getIndent(indent),
							 self.object.renderAsCode(), 
							 params)
	# -- end -- << VBCall methods >>
# << Classes >> (26 of 53)
class VBExitStatement(VBConsumer):
	"""Represents an exit statement"""

	# << VBExitStatement methods >>
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		indenter = self.getIndent(indent)
		if self.element.text == "Exit Function":
			return "%sreturn %s\n" % (indenter, Config["Functions", "ReturnVariableName"])
		elif self.element.text == "Exit Sub":
			return "%sreturn\n" % indenter
		else:
			return "%sbreak\n" % indenter
	# -- end -- << VBExitStatement methods >>
# << Classes >> (27 of 53)
class VBComment(VBConsumer):
	"""Represents an comment"""

	# << VBComment methods >>
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		return self.getIndent(indent) + "#%s\n" % self.element.text
	# -- end -- << VBComment methods >>
# << Classes >> (28 of 53)
class VBLabel(VBUnrendered):
	"""Represents a label"""

	def renderAsCode(self, indent):
		"""Render the label"""
		if Config["Labels", "IgnoreLabels"] == "Yes":
			return ""
		else:
			return super(VBLabel, self).renderAsCode(indent)
# << Classes >> (29 of 53)
class VBOpen(VBCodeBlock):
	"""Represents an open statement"""

	# << VBOpen methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the open"""
		super(VBOpen, self).__init__(scope)
		#
		self.filename = None
		self.open_modes = []
		self.channel = None
		#
		self.auto_class_handlers = ({
			"filename" : (VBParExpression, "filename"),
			"open_mode" : (VBConsumer, self.open_modes),
			"channel" : (VBParExpression, "channel"),
		})
		#
		self.open_mode_lookup = {
			"Input" : "r",
			"Output" : "w",
			"Append" : "a",
			"Binary" : "b",
		}
	# << VBOpen methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		file_mode = ""
		todo = []
		for mode in self.open_modes:
			m = mode.element.text.strip()
			try:
				file_mode += self.open_mode_lookup[m.strip()]
			except KeyError:
				todo.append("'%s'" % m.strip())
		if todo:
			todo_warning = self.getWarning("UnknownFileMode", ", ".join(todo))	
		else:
			todo_warning = ""
		#
		return "%sVBFiles.openFile(%s, %s, '%s') %s\n" % (
					self.getIndent(indent),
					self.channel.renderAsCode(),
					self.filename.renderAsCode(),
					file_mode,
					todo_warning)
	# -- end -- << VBOpen methods >>
# << Classes >> (30 of 53)
class VBClose(VBCodeBlock):
	"""Represents a close statement"""

	# << VBClose methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the open"""
		super(VBClose, self).__init__(scope)
		#
		self.channel = None
		#
		self.auto_class_handlers = ({
			"expression" : (VBParExpression, "channel"),
		})
	# << VBClose methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		return "%sVBFiles.closeFile(%s)\n" % (
					self.getIndent(indent),
					self.channel.renderAsCode())
	# -- end -- << VBClose methods >>
# << Classes >> (31 of 53)
class VBInput(VBCodeBlock):
	"""Represents an input statement"""

	input_type = "Input"

	# << VBInput methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the open"""
		super(VBInput, self).__init__(scope)
		#
		self.channel = None
		self.variables = []
		#
		self.auto_class_handlers = ({
			"channel_id" : (VBParExpression, "channel"),
			"expression" : (VBExpression, self.variables),
		})
	# << VBInput methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		return "%s%s = VBFiles.get%s(%s, %d)\n" % (
					self.getIndent(indent),
					", ".join([var.renderAsCode() for var in self.variables]),
					self.input_type,
					self.channel.renderAsCode(),
					len(self.variables))
	# -- end -- << VBInput methods >>
# << Classes >> (32 of 53)
class VBLineInput(VBInput):
	"""Represents an input statement"""

	input_type = "LineInput"
# << Classes >> (33 of 53)
class VBPrint(VBCodeBlock):
	"""Represents a print statement"""

	# << VBPrint methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the print"""
		super(VBPrint, self).__init__(scope)
		#
		self.channel = None
		self.variables = []
		self.hold_cr = None
		#
		self.auto_class_handlers = ({
			"channel_id" : (VBParExpression, "channel"),
			"expression" : (VBExpression, self.variables),
			"print_separator" : (VBPrintSeparator, self.variables),
		})
	# << VBPrint methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		print_list = ", ".join([var.renderAsCode() for var in self.variables if var.renderAsCode()])
		if self.variables:
			if self.variables[-1].renderAsCode() not in (None, "\t"):
				print_list += ", '\\n'"
		return "%sVBFiles.writeText(%s, %s)\n" % (
					self.getIndent(indent),
					self.channel.renderAsCode(),
					print_list)
	# -- end -- << VBPrint methods >>
# << Classes >> (34 of 53)
class VBPrintSeparator(VBConsumer):
	"""Represents a print statement separator"""

	# << VBPrintSeparator methods >>
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		if self.element.text == ";":
			return None
		elif self.element.text == ",":
			return '"\\t"'
		else:
			raise UnhandledStructureError("Unknown print separator '%s'" % self.element.text)
	# -- end -- << VBPrintSeparator methods >>
# << Classes >> (35 of 53)
class VBUserType(VBCodeBlock):
	"""Represents a select block"""

	auto_handlers = [
	]

	select_variable_index = 0

	# << VBUserType methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the Select"""
		super(VBUserType, self).__init__(scope)
		#
		self.variables = []
		self.identifier = None
		#
		self.auto_class_handlers = {
			"identifier" : (VBConsumer, "identifier"),
			"object_definition" : (VBVariable, self.variables),
		}
	# << VBUserType methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		vars = []
		for var in self.variables:
			vars.append("%sself.%s = %s()" % (
							self.getIndent(indent+2),
							var.identifier,
							var.type))
		#
		return ("%sclass %s:\n"
				"%sdef __init__(self):\n%s\n\n" % (
					self.getIndent(indent),
					self.identifier.element.text,
					self.getIndent(indent+1),
					"\n".join(vars)))
	# -- end -- << VBUserType methods >>
# << Classes >> (36 of 53)
class VBSubroutine(VBCodeBlock):
	"""Represents a subroutine"""

	# << VBSubroutine methods >> (1 of 4)
	def __init__(self, scope="Private"):
		"""Initialize the subroutine"""
		super(VBSubroutine, self).__init__(scope)
		self.identifier = None
		self.scope = scope
		self.block = VBPass()
		self.parameters = []
		self.type = None
		#
		self.auto_class_handlers.update({
			"formal_param" : (VBVariable, self.parameters),
			"block" : (VBCodeBlock, "block"),
			"type_definition" : (VBUnrendered, "type"),
		})

		self.auto_handlers = [
				"identifier",
				"scope",
		]

		self.skip_handlers = [
				"sub_definition",
		]
	# << VBSubroutine methods >> (2 of 4)
	def renderAsCode(self, indent=0):
		"""Render this subroutine"""

		ret = "%sdef %s(%s):\n%s\n" % (
					self.getIndent(indent),
					self.identifier,
					self.renderParameters(),
					self.block.renderAsCode(indent+1))
		return ret
	# << VBSubroutine methods >> (3 of 4)
	def renderParameters(self):
		"""Render the parameter list"""
		params = [param.renderAsCode() for param in self.parameters]
		if self.getParentProperty("convert_functions_to_methods"):
			params.insert(0, "self")
		return ", ".join(params)
	# << VBSubroutine methods >> (4 of 4)
	def resolveLocalName(self, name):
		"""Convert a local name to a fully resolved name

		We search our local variables and parameters to see if we know the name. If we do then we
		return the original name.

		"""
		names = [obj.identifier for obj in self.block.locals + self.parameters]
		if name in names:
			return name
		else:
			raise UnresolvableName("Name '%s' is not known in this namespace" % name)
	# -- end -- << VBSubroutine methods >>
# << Classes >> (37 of 53)
class VBFunction(VBSubroutine):
	"""Represents a function"""

	# << VBFunction methods >>
	def renderAsCode(self, indent=0):
		"""Render this subroutine"""
		#
		# Set a name conversion to capture the function name
		# Assignments to this function name should go to the _ret parameter
		return_var = Config["Functions", "ReturnVariableName"]
		self.name_substitution[self.identifier] = return_var
		#
		if self.block:
			block = self.block.renderAsCode(indent+1)
		else:
			block = self.getIndent(indent+1) + "pass\n"
		#
		if Config["Functions", "PreInitializeReturnVariable"] == "Yes":
			pre_init = "%s%s = None\n" % (				
					self.getIndent(indent+1),
					return_var)
		else:
			pre_init = ""

		ret = "%sdef %s(%s):\n%s%s%sreturn %s\n\n" % (
					self.getIndent(indent),
					self.identifier, 
					self.renderParameters(),
					pre_init,
					block,
					self.getIndent(indent+1),
					return_var)
		return ret
	# -- end -- << VBFunction methods >>
# << Classes >> (38 of 53)
class VBIf(VBCodeBlock):
	"""Represents an if block"""

	auto_handlers = [
	]

	skip_handlers = [
			"if_statement",
	]


	# << VBIf methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the If"""
		super(VBIf, self).__init__(scope)
		#
		self.condition = None
		self.if_block = VBPass()
		self.elif_blocks = []
		self.else_block = None
		#
		self.auto_class_handlers = {
			"condition" : (VBExpression, "condition"),
			"if_block" : (VBCodeBlock, "if_block"),
			"else_if_statement" : (VBElseIf, self.elif_blocks),
			"else_block" : (VBCodeBlock, "else_block"),
		}
	# << VBIf methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		ret = self.getIndent(indent) + "if %s:\n" % self.condition.renderAsCode()
		ret += self.if_block.renderAsCode(indent+1)
		if self.elif_blocks:
			for elif_block in self.elif_blocks:
				ret += elif_block.renderAsCode(indent)
		if self.else_block:
			ret += self.getIndent(indent) + "else:\n"
			ret += self.else_block.renderAsCode(indent+1)
		return ret
	# -- end -- << VBIf methods >>
# << Classes >> (39 of 53)
class VBElseIf(VBIf):
	"""Represents an ElseIf statement"""

	# << VBElseIf methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the If"""
		super(VBIf, self).__init__(scope)
		#
		self.condition = None
		self.elif_block = VBPass()
		#
		self.auto_class_handlers = {
			"condition" : (VBExpression, "condition"),
			"else_if_block" : (VBCodeBlock, "elif_block"),
		}
	# << VBElseIf methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		ret = self.getIndent(indent) + "elif %s:\n" % self.condition.renderAsCode()
		ret += self.elif_block.renderAsCode(indent+1)
		return ret
	# -- end -- << VBElseIf methods >>
# << Classes >> (40 of 53)
class VBInlineIf(VBCodeBlock):
	"""Represents an if block"""

	auto_handlers = [
	]

	skip_handlers = [
			"if_statement",
	]


	# << VBInlineIf methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the If"""
		super(VBInlineIf, self).__init__(scope)
		#
		self.condition = None
		self.statements = []
		#
		self.auto_class_handlers = {
			"condition" : (VBExpression, "condition"),
			"statement" : (VBCodeBlock, self.statements),
			"implicit_call_statement" : (VBCodeBlock, self.statements),  # TODO: remove me
		}
	# << VBInlineIf methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		assert self.statements, "Inline If has no statements!"

		ret = "%sif %s:\n%s" % (
					self.getIndent(indent),
					self.condition.renderAsCode(),
					self.statements[0].renderAsCode(indent+1),)
		#
		if len(self.statements) == 2:
			ret += "%selse:\n%s" % (
					self.getIndent(indent),
					self.statements[1].renderAsCode(indent+1))
		elif len(self.statements) > 2:
			raise VBParserError("Inline if with more than one clause not supported")
		#
		return ret
	# -- end -- << VBInlineIf methods >>
# << Classes >> (41 of 53)
class VBSelect(VBCodeBlock):
	"""Represents a select block"""

	auto_handlers = [
	]

	_select_variable_index = 0

	# << VBSelect methods >> (1 of 3)
	def __init__(self, scope="Private"):
		"""Initialize the Select"""
		super(VBSelect, self).__init__(scope)
		#
		self.blocks = []
		#
		self.auto_class_handlers = {
			"expression" : (VBExpression, "expression"),
			"case_item_block" : (VBCaseItem, self.blocks),
			"case_else_block" : (VBCaseElse, self.blocks),
		}
		#
		# Change the variable index if we are a select
		if self.__class__ == VBSelect:
			self.select_variable_index = VBSelect._select_variable_index
			VBSelect._select_variable_index = VBSelect._select_variable_index + 1
	# << VBSelect methods >> (2 of 3)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		#
		# Change if/elif status on the first child
		if self.blocks:
			self.blocks[0].if_or_elif = "if"
		#
		if Config["Select", "EvaluateVariable"] <> "EachTime":
			ret = "%s%s = %s\n" % (self.getIndent(indent),
									 self.getSelectVariable(),
									 self.expression.renderAsCode())
		else:
			ret = ""
		ret += "".join([item.renderAsCode(indent) for item in self.blocks])
		return ret
	# << VBSelect methods >> (3 of 3)
	def getSelectVariable(self):
		"""Return the name of the select variable"""
		eval_variable = Config["Select", "EvaluateVariable"]
		if eval_variable == "Once":
			if Config["Select", "UseNumericIndex"] == "Yes":
				select_var = "%s%d" % (Config["Select", "SelectVariablePrefix"], 
									   self.getParentProperty("select_variable_index"))
			else:
				select_var = Config["Select", "SelectVariablePrefix"]
		elif eval_variable == "EachTime":
			select_var = "%s" % self.getParentProperty("expression").renderAsCode()
		else:
			raise InvalidOption("Evaluate variable option not understood: '%s'" % eval_variable)
		return select_var
	# -- end -- << VBSelect methods >>
# << Classes >> (42 of 53)
class VBCaseBlock(VBSelect):
	"""Represents a select block"""

	if_or_elif = "elif" # Our parent will change this if we are the first

	# << VBCaseBlock methods >>
	def __init__(self, scope="Private"):
		"""Initialize the Select"""
		super(VBCaseBlock, self).__init__(scope)
		#
		self.lists = []
		self.expressions = []
		self.block = VBPass()
		#
		self.auto_class_handlers = {
			"case_list" : (VBCaseItem, self.lists),
			"expression" : (VBExpression, self.expressions),
			"block" : (VBCodeBlock, "block"),
		}
	# -- end -- << VBCaseBlock methods >>
# << Classes >> (43 of 53)
class VBCaseItem(VBCaseBlock):
	"""Represents a select block"""

	# << VBCaseItem methods >>
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		select_variable_index = self.getParentProperty("select_variable_index")
		if self.lists:
			expr = " or ".join(["(%s)" % item.renderAsCode() for item in self.lists])
			return "%s%s %s:\n%s" % (
						   self.getIndent(indent),
						   self.if_or_elif,
						   expr,
						   self.block.renderAsCode(indent+1))						   
		elif len(self.expressions) == 1:
			return "%s == %s" % (
										   self.getSelectVariable(),
										   self.expressions[0].renderAsCode())
		elif len(self.expressions) == 2:
			return "%s <= %s <= %s" % (
										   self.expressions[0].renderAsCode(),
										   self.getSelectVariable(),
										   self.expressions[1].renderAsCode())
		raise VBParserError("Error rendering case item")
	# -- end -- << VBCaseItem methods >>
# << Classes >> (44 of 53)
class VBCaseElse(VBCaseBlock):
	"""Represents a select block"""

	# << VBCaseElse methods >>
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		return "%selse:\n%s" % (self.getIndent(indent),
								 self.block.renderAsCode(indent+1))
	# -- end -- << VBCaseElse methods >>
# << Classes >> (45 of 53)
class VBFor(VBCodeBlock):
	"""Represents a for statement"""

	# << VBFor methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the Select"""
		super(VBFor, self).__init__(scope)
		#
		self.block = VBPass()
		self.expressions = []
		#
		self.auto_class_handlers = {
			"expression" : (VBExpression, self.expressions),
			"block" : (VBCodeBlock, "block"),
		}

		self.auto_handlers = [
			"identifier",
		]
	# << VBFor methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		range_statement = ", ".join([item.renderAsCode() for item in self.expressions])
		return "%sfor %s in vbForRange(%s):\n%s" % (
								 self.getIndent(indent),
								 self.identifier,
								 range_statement,
								 self.block.renderAsCode(indent+1))
	# -- end -- << VBFor methods >>
# << Classes >> (46 of 53)
class VBForEach(VBFor):
	"""Represents a for each statement"""

	# << VBForEach methods >>
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		return "%sfor %s in %s:\n%s" % (
								 self.getIndent(indent),
								 self.identifier,
								 self.expressions[0].renderAsCode(),
								 self.block.renderAsCode(indent+1))
	# -- end -- << VBForEach methods >>
# << Classes >> (47 of 53)
class VBWhile(VBCodeBlock):
	"""Represents a while statement"""

	# << VBWhile methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the Select"""
		super(VBWhile, self).__init__(scope)	
		#
		self.block = VBPass()
		self.expression = None
		#
		self.auto_class_handlers = {
			"expression" : (VBExpression, "expression"),
			"block" : (VBCodeBlock, "block"),
		}
	# << VBWhile methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		return "%swhile %s:\n%s" % (
							self.getIndent(indent),
							self.expression.renderAsCode(),
							self.block.renderAsCode(indent+1))
	# -- end -- << VBWhile methods >>
# << Classes >> (48 of 53)
class VBDo(VBCodeBlock):
	"""Represents a do statement"""

	# << VBDo methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the Select"""
		super(VBDo, self).__init__(scope)
		#
		self.block = VBPass()
		self.pre_while = None
		self.pre_until = None
		self.post_while = None
		self.post_until = None
		#
		self.auto_class_handlers = {
			"while_clause" : (VBExpression, "pre_while"),
			"until_clause" : (VBExpression, "pre_until"),
			"post_while_clause" : (VBExpression, "post_while"),
			"post_until_clause" : (VBExpression, "post_until"),
			"block" : (VBCodeBlock, "block"),
		}
	# << VBDo methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""Render this element as code

		There are five different kinds of do loop
			pre_while
			pre_until
			post_while
			post_until
			no conditions

		"""
		if self.pre_while:
			return "%swhile %s:\n%s" % (
							self.getIndent(indent),
							self.pre_while.renderAsCode(),
							self.block.renderAsCode(indent+1))
		elif self.pre_until:
			return "%swhile not (%s):\n%s" % (
							self.getIndent(indent),
							self.pre_until.renderAsCode(),
							self.block.renderAsCode(indent+1))
		elif self.post_while:
			return "%swhile 1:\n%s%sif not (%s):\n%sbreak\n" % (
							self.getIndent(indent),
							self.block.renderAsCode(indent+1),
							self.getIndent(indent+1),
							self.post_while.renderAsCode(),
							self.getIndent(indent+2))
		elif self.post_until:
			return "%swhile 1:\n%s%sif %s:\n%sbreak\n" % (
							self.getIndent(indent),
							self.block.renderAsCode(indent+1),
							self.getIndent(indent+1),
							self.post_until.renderAsCode(),
							self.getIndent(indent+2))						
		else:
			return "%swhile 1:\n%s" % (
							self.getIndent(indent),
							self.block.renderAsCode(indent+1))
	# -- end -- << VBDo methods >>
# << Classes >> (49 of 53)
class VBWith(VBCodeBlock):
	"""Represents a with statement"""

	_with_variable_index = 0

	# << VBWith methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the Select"""
		super(VBWith, self).__init__(scope)
		#
		self.block = None
		self.expression = None
		#
		self.auto_class_handlers = {
			"expression" : (VBExpression, "expression"),
			"block" : (VBCodeBlock, "block"),
		}
		#
		self.with_variable_index = VBWith._with_variable_index
		VBWith._with_variable_index = VBWith._with_variable_index + 1
	# << VBWith methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""Render this element as code"""
		if self.checkOptionChoice("With", "EvaluateVariable", ("EveryTime", "Once")) == "EveryTime":
			self.with_object = self.expression.renderAsCode()
			return self.block.renderAsCode(indent)
		else:
			if self.checkOptionYesNo("With", "UseNumericIndex") == "Yes":
				varname = "%s%d" % (Config["With", "WithVariablePrefix"],
									self.with_variable_index)
			else:
				varname = Config["With", "WithVariablePrefix"]

			self.with_object = varname

			return "%s%s = %s\n%s" % (
							self.getIndent(indent),
							varname,
							self.expression.renderAsCode(),
							self.block.renderAsCode(indent))
	# -- end -- << VBWith methods >>
# << Classes >> (50 of 53)
class VBProperty(VBSubroutine):
	"""Represents a property definition"""

	# << VBProperty methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the Select"""
		super(VBProperty, self).__init__(scope)
		self.property_decorator_type = None
		#
		self.auto_handlers.append("property_decorator_type")
	# << VBProperty methods >> (2 of 2)
	def renderPropertyGroup(self, indent, name, Let=None, Set=None, Get=None):
		"""Render a group of property statements"""
		if Let and Set:
			raise UnhandledStructureError("Cannot handle both Let and Set properties for an object")

		log.info("Rendering property group '%s'" % name)

		ret = []
		params = []
		pset = Let or Set
		pget = Get

		if pset:
			self.getParentProperty("local_names").append(pset.identifier) # Store property name for namespace analysis
			pset.identifier = "%s%s" % (Config["Properties", "LetSetVariablePrefix"], pset.identifier)		
			ret.append(pset.renderAsCode(indent))
			params.append("fset=%s" % pset.identifier)
		if pget:
			self.getParentProperty("local_names").append(pget.identifier) # Store property name for namespace analysis
			pget.__class__ = VBFunction # Needs to be a function
			pget.name_substitution[pget.identifier] = Config["Functions", "ReturnVariableName"]
			pget.identifier = "%s%s" % (Config["Properties", "GetVariablePrefix"], pget.identifier)		
			ret.append(pget.renderAsCode(indent))
			params.append("fget=%s" % pget.identifier)

		return "%s%s%s = property(%s)\n\n" % (
					"".join(ret),
					self.getIndent(indent),
					name,
					", ".join(params))
	# -- end -- << VBProperty methods >>
# << Classes >> (51 of 53)
class VBEnum(VBCodeBlock):
	"""Represents an enum definition"""

	# << VBEnum methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the Select"""
		super(VBEnum, self).__init__(scope)
		self.enumerations = []
		self.identifier = None
		#
		self.auto_class_handlers = {
				"enumeration_item" : (VBConsumer, self.enumerations),
			}

		self.auto_handlers = ["identifier"]
	# << VBEnum methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""Render a group of property statements"""
		ret = ["%s%s = %d" % (self.getIndent(indent), enum.element.text, id) 
				for enum, id in zip(self.enumerations, xrange(sys.maxint))]

		return "%s# Enumeration '%s'\n%s\n" % (
							self.getIndent(indent),
							self.identifier,
							"\n".join(ret),
					)
	# -- end -- << VBEnum methods >>
# << Classes >> (52 of 53)
class VB2PYDirective(VBCodeBlock):
	"""Handles a vb2py directive"""

	skip_handlers = [
			"vb2py_directive",
	]

	# << VB2PYDirective methods >> (1 of 2)
	def __init__(self, scope="Private"):
		"""Initialize the module"""
		super(VB2PYDirective, self).__init__(scope)
		self.auto_handlers = (
			"directive_type",
			"config_name",
			"config_section",
			"expression",
		)
		self.directive_type = "Set"
		self.config_name = None
		self.config_section = None
		self.expression = None
	# << VB2PYDirective methods >> (2 of 2)
	def renderAsCode(self, indent=0):
		"""We use the rendering to do our stuff"""
		if self.directive_type == "Set":
			Config.setLocalOveride(self.config_section, self.config_name, self.expression)
			log.info("Doing a set: %s" % str((self.config_section, self.config_name, self.expression)))
		elif self.directive_type == "Unset":
			Config.removeLocalOveride(self.config_section, self.config_name)
			log.info("Doing an uset: %s" % str((self.config_section, self.config_name)))
		else:
			raise DirectiveError("Directive not understood: '%s'" % self.directive_type)
		return ""
	# -- end -- << VB2PYDirective methods >>
# << Classes >> (53 of 53)
class VBPass(VBCodeBlock):
	"""Represents an empty statement"""

	def renderAsCode(self, indent):
		"""Render it!"""
		return "%spass\n" % (self.getIndent(indent),)
# -- end -- << Classes >>

# Plug-ins
import extensions
plugins = extensions.loadAllPlugins()


if __name__ == "__main__":	
	from testparse import txt
	m = parseVB(txt)
