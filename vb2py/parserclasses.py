"""A set of classes used during the parsing of VB code"""

# << Definitions >>
StopSearch = -9999 # Used to terminate searches for parent properties
# -- end -- << Definitions >>
# << Classes >> (1 of 74)
class VBElement(object):
    """An element of VB code"""

    # << VBElement methods >> (1 of 2)
    def __init__(self, details, text):
        """Initialize from the details"""
        #import pdb; pdb.set_trace()
        self.name = details[0]
        self.text = makeUnicodeFromSafe(text[details[1]:details[2]])
        self.elements = convertToElements(details[3], text)
    # << VBElement methods >> (2 of 2)
    def printTree(self, offset=0):
        """Print out this tree"""
        print "%s%s : '%s'" % (" "*offset, self.name, self.text.split("\n")[:20])
        for subelement in self.elements:
            subelement.printTree(offset+1)
    # -- end -- << VBElement methods >>
# << Classes >> (2 of 74)
class VBFailedElement(object):
    """An failed element of VB code"""

    # << VBFailedElement methods >>
    def __init__(self, name, text):
        """Initialize from the details"""
        self.name = name
        self.text = text
        self.elements = []
    # -- end -- << VBFailedElement methods >>
# << Classes >> (3 of 74)
class VBNamespace(object):
    """Handles a VB Namespace"""

    # << VBNamespace declarations >>
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

    # 
    # Set this to 1 if the object is a function (ie requires () when accessing)
    is_function = 0

    #
    # Set to 1 for types which would mark the end of the docstrings
    would_end_docstring = 1


    #
    # Intrinsic VB functions - we need to know these to be able to convert
    # bare references (eg Dir) to function references (Dir())
    intrinsic_functions = [
        "Dir", "FreeFile", "Rnd", "Timer",
    ]
    # -- end -- << VBNamespace declarations >>
    # << VBNamespace methods >> (1 of 28)
    def __init__(self, scope="Private"):
        """Initialize the namespace"""
        self.locals = []
        self.local_default_scope = self.default_scope
        self.auto_class_handlers = {
            "object_definition" : (VBVariableDefinition, self.locals),
            "const_definition"   : (VBConstant, self.locals),
            "user_type_definition" : (VBUserType, self.locals),
            "event_definition" : (VBUnrendered, self.locals),
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
    # << VBNamespace methods >> (2 of 28)
    def amGlobal(self, scope):
        """Decide if a variable will be considered a global

            The algorithm works by asking our parent for a 'public_is_global' flag. If this
            is true and the scope is either 'public' or 'global' then we are a global. It is
            up to each parent to decide if publics are global. Things like code modules will have this
            set whereas things like subroutines will not.

            """
        #
        # First throw out anything which is private
        log.info("Checking if global: '%s' scope is '%s'" % (self, scope))
        if scope in ("Public", "Global"):
            if self.getParentProperty("public_is_global", 0):
                log.info("We are global!")
                return 1
        return 0
    # << VBNamespace methods >> (3 of 28)
    def assignParent(self, parent):
        """Set our parent

            This is kept as a separate method because it is a useful hook for subclasses.
            Once this method is called, the object is fully initialized.

            """
        self.parent = parent
    # << VBNamespace methods >> (4 of 28)
    def asString(self):
        """Convert to a nice representation"""
        return repr(self)
    # << VBNamespace methods >> (5 of 28)
    def checkIfFunction(self, name):
        """Check if the name is a function or not"""
        for loc in self.locals:
            if loc.identifier == name:
                return loc.is_function
        raise UnresolvableName("Name '%s' is not known in this context" % name)
    # << VBNamespace methods >> (6 of 28)
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
    # << VBNamespace methods >> (7 of 28)
    def checkOptionYesNo(self, section, name):
        """Return the yes/no value of an option checking for invalid answers"""
        return self.checkOptionChoice(section, name, ("Yes", "No"))
    # << VBNamespace methods >> (8 of 28)
    def containsStatements(self):
        """Check if we contain statements"""
        #
        # TODO: This needs refactoring - it is horrible
        if isinstance(self, NonCodeBlocks):
            return 0
        if not hasattr(self, "blocks"):
            return 1
        elif self.blocks:
            for item in self.blocks:
                if item.containsStatements():
                    return 1
            return 0
        else:
            return 1
    # << VBNamespace methods >> (9 of 28)
    def createExtractHandler(self, token):
        """Create a handler which will extract a certain token value"""
        def handler(element):
            log.info("Grabbed attribute '%s' for %s as '%s'" % (token, self, element.text))
            setattr(self, token, element.text)
        return handler
    # << VBNamespace methods >> (10 of 28)
    def filterListByClass(self, sequence, cls): 
        """Return all elements of sequence that are an instance of the given class"""
        return [item for item in sequence if isinstance(item, cls)]
    # << VBNamespace methods >> (11 of 28)
    def finalizeObject(self):
        """Finalize the object

            This method is called once the object has been completely parsed and can
            be used to do any processing required.

            """
    # << VBNamespace methods >> (12 of 28)
    def findParentOfClass(self, cls):
        """Return our nearest parent who is a subclass of cls"""
        try:
            parent = self.parent
        except AttributeError:
            raise NestingError("Reached outer layer when looking for parent of class")
        if isinstance(parent, cls):
            return parent
        else:
            return parent.findParentOfClass(cls)
    # << VBNamespace methods >> (13 of 28)
    def getHandler(self, element):
        """Find a handler for the element"""
        if element.name in self.skip_handlers:
            return None
        elif element.name in self.auto_handlers:
            log.info("Found auto handler for '%s' ('%s')" % (element.name, self))
            return self.createExtractHandler(element.name)
        elif element.name in self.auto_class_handlers:
            log.info("Found auto handler for '%s' ('%s')" % (element.name, self))
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
    # << VBNamespace methods >> (14 of 28)
    def getIndent(self, indent):
        """Return some spaces to do indenting"""
        return self._indent_char*indent*self._indent_amount
    # << VBNamespace methods >> (15 of 28)
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
    # << VBNamespace methods >> (16 of 28)
    def getParentProperty(self, name, default=None):
        """Get a property from our nearest ancestor who has it"""
        try:
            return getattr(self, name)
        except AttributeError:
            try:
                parent = self.parent
                return parent.getParentProperty(name)
            except AttributeError:
                if default is not None:
                    return default
                raise NestingError("Reached outer level when trying to access a parent property: '%s'" % name)
    # << VBNamespace methods >> (17 of 28)
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
    # << VBNamespace methods >> (18 of 28)
    def handleSubObject(self, element, obj_class, add_to):
        """Handle an object which creates a sub object"""
        v = obj_class(self.local_default_scope)
        v.processElement(element)
        v.assignParent(self)
        v.finalizeObject()
        #
        # Assume that we are supposed to add this to a list of items
        # if this fails then perhaps this is an attribute we are supposed to set
        try:
            add_to.append(v)	
        except AttributeError:
            setattr(self, add_to, v)
        #
        log.info("Added new %s to %s" % (obj_class, self.asString()))
    # << VBNamespace methods >> (19 of 28)
    def isAFunction(self, name):
        """Check if the name is a function or not

            We traverse up through the nested namespaces until someone knows
            the name and then see if they are a function.

            """
        if name in self.intrinsic_functions:
            return 1
        try:
            return self.checkIfFunction(name)
        except UnresolvableName:
            try:
                return self.parent.isAFunction(name)
            except (AttributeError, UnresolvableName):
                return 0	# Nobody knew the name so we can't know if it is or not
    # << VBNamespace methods >> (20 of 28)
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
    # << VBNamespace methods >> (21 of 28)
    def registerAsGlobal(self):
        """Register ourselves as a global object

            We try to add ourselves to our parents "global_objects" table. This may fail
            if we are not owned by anything that has a global_obects table, as would be
            the case for converting a simple block of text.

            """
        try:
            global_objects = self.getParentProperty("global_objects")
        except NestingError:
            log.warn("Tried to register global object but there was no suitable object table")
        else:
            global_objects[self.identifier] = self
            log.info("Registered a new global object: '%s'" % self)
    # << VBNamespace methods >> (22 of 28)
    def registerImportRequired(self, modulename):
        """Register a need to import a certain module

            When we need to use a variable from another module we need to tell our module-like
            containner to add an 'import' statement. So we search for such a container and try
            to add the module name to the import list.

            It is possible (but unlikely) that we need the import but we are not in a container.
            If this happens we just warning and carry on.

            """
        try:
            module_imports = self.getParentProperty("module_imports")
        except NestingError:
            log.warn("Tried to request a module import (%s) but couldn't find a suitable container" % modulename)
        else:
            if modulename not in module_imports:
                module_imports.append(modulename)
            log.info("Registered a new module import: '%s'" % modulename)
    # << VBNamespace methods >> (23 of 28)
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        return self.getIndent(indent) + "# Unrendered object %s\n" % (self.asString(), )
    # << VBNamespace methods >> (24 of 28)
    def resolveLocalName(self, name, rendering_locals=0, requestedby=None):
        """Convert a local name to a fully resolved name"""
        raise UnresolvableName("Name '%s' is not known in this namespace" % name)
    # << VBNamespace methods >> (25 of 28)
    def resolveName(self, name, rendering_locals=None, requestedby=None):
        """Convert a local name to a fully resolved name

            We traverse up through the nested namespaces until someone knows
            what to do with the name. If nobody knows then we know if must be
            a local so it keeps the same name.

            """
        if rendering_locals is None:
            rendering_locals = self.getParentProperty("rendering_locals")
        if not requestedby:
            requestedby = self		
        try:
            return self.resolveLocalName(name, rendering_locals, requestedby=requestedby)
        except UnresolvableName:
            try:
                return self.parent.resolveName(name, rendering_locals, requestedby=requestedby)
            except AttributeError:
                return name	# Nobody knew the name so it must be local
    # << VBNamespace methods >> (26 of 28)
    def searchParentProperty(self, name):
        """Search for any ancestor who has the named parameter set to true

            Stop searching if someone has the property set to StopSearch

            """
        try:
            if getattr(self, name) == StopSearch:
                return 0
            elif getattr(self, name):
                return 1
        except AttributeError:
            pass
        try:
            parent = self.parent
            return parent.searchParentProperty(name)
        except AttributeError:
            return 0
    # << VBNamespace methods >> (27 of 28)
    def handle_scope(self, element):
        """Handle a scope definition"""
        self.local_default_scope = element.text
        log.info("Changed default scope to %s" % self.local_default_scope)
    # << VBNamespace methods >> (28 of 28)
    def handle_line_end(self, element):
        """Handle the end of a line"""
        self.local_default_scope = self.default_scope
    # -- end -- << VBNamespace methods >>
# << Classes >> (4 of 74)
class VBConsumer(VBNamespace):
    """Consume and store elements"""

    def processElement(self, element):
        """Eat this element"""
        self.element = element
        log.info("Consumed element: %s" % element)
# << Classes >> (5 of 74)
class VBUnrendered(VBConsumer):
    """Represents an unrendered statement"""

    would_end_docstring = 0 

    def renderAsCode(self, indent):
        """Render the unrendrable!"""
        if self.checkOptionYesNo("General", "WarnAboutUnrenderedCode") == "Yes":
            return self.getWarning("UntranslatedCode", self.element.text.replace("\n", "\\n"), indent, crlf=1)
        else:
            return ""
# << Classes >> (6 of 74)
class VBMessage(VBUnrendered):
    """Allows a message to be placed in the python output"""

    def __init__(self, scope="Private", message="No message", messagetype="Unknown"):
        """Initialise the message"""
        super(VBMessage, self).__init__(scope)
        self.message = message
        self.messagetype = messagetype

    def renderAsCode(self, indent=0):
        """Render the message"""
        return self.getWarning(self.messagetype, 
                               self.message, indent, crlf=1)
# << Classes >> (7 of 74)
class VBMissingArgument(VBConsumer):
    """Represents an missing argument"""

    def renderAsCode(self, indent=0):
        """Render the unrendrable!"""
        return "VBMissingArgument"
# << Classes >> (8 of 74)
class VBCodeBlock(VBNamespace):
    """A block of VB code"""

    # << VBCodeBlock methods >> (1 of 2)
    def __init__(self, scope="Private"):
        """Initialize the block"""
        super(VBCodeBlock, self).__init__()
        self.blocks = []
        self.auto_class_handlers.update({
            "assignment_statement" : (VBAssignment, self.blocks),
            "lset_statement" : (VBLSet, self.blocks),
            "rset_statement" : (VBRSet, self.blocks),
            "set_statement" : (VBSet, self.blocks),
            "comment_body" : (VBComment, self.blocks),
            "vb2py_directive" : (VB2PYDirective, self.blocks),
            "if_statement" : (VBIf, self.blocks),
            "inline_if_statement" : (VBInlineIf, self.blocks),
            "select_statement" : (VBSelect, self.blocks),
            "exit_statement" : (VBExitStatement, self.blocks),
            "while_statement" : (VBWhile, self.blocks),
            "do_statement" : (VBDo, self.blocks),
            "redim_statement" : (VBReDim, self.blocks),
            "implicit_call_statement" : (VBCall, self.blocks),
            "inline_implicit_call" : (VBCall, self.blocks),
            "label_statement" : (VBLabel, self.blocks),
            "with_statement" : (VBWith, self.blocks),
            "end_statement" : (VBEnd, self.blocks),

            "for_statement" : (VBFor, self.blocks),
            "inline_for_statement" : (VBFor, self.blocks),
            "for_each_statement" : (VBForEach, self.blocks),

            "open_statement" : (VBOpen, self.blocks),
            "close_statement" : (VBClose, self.blocks),
            "input_statement" : (VBInput, self.blocks),
            "print_statement" : (VBPrint, self.blocks),
            "line_input_statement" : (VBLineInput, self.blocks),
            "seek_statement" : (VBSeek, self.blocks),
            "name_statement" : (VBName, self.blocks),

            "attribute_statement" : (VBUnrendered, self.blocks),
            "resume_statement" : (VBUnrendered, self.blocks),
            "goto_statement" : (VBUnrendered, self.blocks),
            "on_statement" : (VBUnrendered, self.blocks),
            "external_declaration" : (VBUnrendered, self.blocks),
            "get_statement" : (VBUnrendered, self.blocks),
            "put_statement" : (VBUnrendered, self.blocks),
            "option_statement" : (VBUnrendered, self.blocks),
            "class_header_block" : (VBUnrenderedBlock, self.blocks),

            "parser_failure" : (VBParserFailure, self.blocks),

        })
    # << VBCodeBlock methods >> (2 of 2)
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        #
        # Watch out for the block not containing any statements (could be all comments!)
        if not self.containsStatements():
            self.blocks.append(VBPass())
        #
        return "".join([block.renderAsCode(indent) for block in self.blocks])
    # -- end -- << VBCodeBlock methods >>
# << Classes >> (9 of 74)
class VBUnrenderedBlock(VBCodeBlock):
    """Represents an unrendered block"""

    would_end_docstring = 0 

    def renderAsCode(self, indent):
        """Render the unrendrable!"""
        return ""
# << Classes >> (10 of 74)
class VBOptionalCodeBlock(VBCodeBlock):
    """A block of VB code which can be empty and still sytactically correct"""

    # << VBOptionalCodeBlock methods >>
    def containsStatements(self, indent=0):
        """Return true if this block contains statements

            We always return 1 here because it doesn't matter if we contain statements of not

            """
        return 1
    # -- end -- << VBOptionalCodeBlock methods >>
# << Classes >> (11 of 74)
class VBVariable(VBNamespace):
    """Handles a VB Variable"""

    # << VBVariable declarations >>
    auto_handlers = [
            "scope",
            "type",
            "string_size_indicator",
            "value",
            "identifier",
            "optional",
            "new_keyword",
            "preserve_keyword",
            "implicit_object",
    ]

    skip_handlers = [
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
        self.size_definitions = []
        self.value = None
        self.optional = None
        self.expression = VBMissingArgument()
        self.new_keyword = None
        self.preserve_keyword = None
        self.string_size_indicator = None
        self.object = None
        self.implicit_object = None
        self.unsized_definition = None

        self.auto_class_handlers = {
            "expression"	: (VBExpression, "expression"),
            "size"	: (VBSizeDefinition, self.size_definitions),
            "size_range"	: (VBSizeDefinition, self.size_definitions),
            "unsized_definition"	: (VBConsumer, "unsized_definition"),
        }
    # << VBVariable methods >> (2 of 3)
    def finalizeObject(self):
        """We can use this opportunity to now determine if we are a global"""
        if self.amGlobal(self.scope):
            self.registerAsGlobal()
    # << VBVariable methods >> (3 of 3)
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        if self.optional:
            return "%s=%s" % (self.identifier, self.expression.renderAsCode())
        else:
            return self.identifier
    # -- end -- << VBVariable methods >>
# << Classes >> (12 of 74)
class VBSizeDefinition(VBNamespace):
    """Handles a VB Variable size definition"""

    # << VBSizeDefinition methods >> (1 of 2)
    def __init__(self, scope="Private"):
        """Initialize the size definition"""
        super(VBSizeDefinition, self).__init__(scope)
        #
        self.expression = None
        self.sizes = []
        self.size_ranges = []
        #
        self.auto_class_handlers = {
            "size"	: (VBExpression, self.sizes),
            "size_range"	: (VBSizeDefinition, self.size_ranges),
        }
    # << VBSizeDefinition methods >> (2 of 2)
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        if self.sizes:
            return ", ".join([item.renderAsCode() for item in self.sizes])
        else:
            return "(%s)" % ", ".join([item.renderAsCode() for item in self.size_ranges])
    # -- end -- << VBSizeDefinition methods >>
# << Classes >> (13 of 74)
class VBObject(VBNamespace):
    """Handles a VB Object"""

    am_on_lhs = 0 # Set to 1 if the object is on the LHS of an assignment

    # << VBObject methods >> (1 of 6)
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
    # << VBObject methods >> (2 of 6)
    def renderAsCode(self, indent=0):
        """Render this subroutine"""
        return self._renderPartialObject(indent)
    # << VBObject methods >> (3 of 6)
    def finalizeObject(self):
        """Finalize the object

            Check for any type markers.

            """
        for obj in [self.primary] + self.modifiers:
            try:
                ending = obj.element.text[-1:] or " "
            except AttributeError:
                pass # It isn't a consumer so we can't check it
            else:
                if ending in "#$%&":
                    log.info("Removed type identifier from '%s'" % obj.element.text)
                    obj.element.text = obj.element.text[:-1]
    # << VBObject methods >> (4 of 6)
    def asString(self):
        """Return a string representation"""
        if self.implicit_object:
            log.info("Ooops an implicit object in definition")
        ret = [self.primary.element.text] + [item.asString() for item in self.modifiers]
        return ".".join(ret)
    # << VBObject methods >> (5 of 6)
    def fnPart(self): 
        """Return the function part of this object (ie without any parameters"""
        return self._renderPartialObject(indent=0, modifier=VBAttribute)
    # << VBObject methods >> (6 of 6)
    def _renderPartialObject(self, indent=0, modifier=None): 
        """Render this object but only including modifiers of a certain class"""
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
        resolved_name = self.resolveName(obj_name)
        #
        # Check if this looks like a function
        # TODO: This isn't very rigorous
        if not self.modifiers:
            if self.isAFunction(obj_name):
                resolved_name += "()"
        #
        if modifier is None:
            valid_modifiers = self.modifiers
        else:
            valid_modifiers = self.filterListByClass(self.modifiers, modifier)
        #
        return "%s%s%s" % (implicit_name,
                           resolved_name,
                           "".join([item.renderAsCode() for item in valid_modifiers]))
    # -- end -- << VBObject methods >>
# << Classes >> (14 of 74)
class VBLHSObject(VBObject):
    """Handles a VB Object appearing on the LHS of an assignment"""

    am_on_lhs = 1 # Set to 1 if the object is on the LHS of an assignment
# << Classes >> (15 of 74)
class VBAttribute(VBConsumer):
    """An attribute of an object"""

    def renderAsCode(self, indent=0):
        """Render this attribute"""
        return ".%s" % self.element.text
# << Classes >> (16 of 74)
class VBParameterList(VBCodeBlock):
    """An parameter list for an object"""

    # << VBParameterList methods >> (1 of 2)
    def __init__(self, scope="Private"):
        """Initialize the object"""
        super(VBParameterList, self).__init__(scope)

        self.expressions = []
        self.auto_class_handlers.update({
            "expression" : (VBExpression, self.expressions),
            "missing_positional" : (VBMissingPositional, self.expressions),
        })
    # << VBParameterList methods >> (2 of 2)
    def renderAsCode(self, indent=0):
        """Render this attribute"""
        #
        # Check if we should replace () with [] - needed on the LHS of an assignment but not
        # elsewhere since __call__ is mapped to __getitem__ for array types
        if self.searchParentProperty("brackets_are_indexes"):
            fmt = "[%s]"
            self.brackets_are_indexes = StopSearch	# Prevents double accounting in a(b(5)) expressions where b is a function		
        else:
            fmt = "(%s)"
        #	
        # Construct the list of parameters - this is harder than it looks because 
        # for any missing positional parameters we have to do some introspection
        # to dig out the default value
        param_list = []
        for idx, element in zip(xrange(1000), self.expressions):
            element.parameter_index_position = idx # Needed so that the element can get its default
            param_list.append(element.renderAsCode())
        #
        content = ", ".join(param_list)
        return fmt % content
    # -- end -- << VBParameterList methods >>
# << Classes >> (17 of 74)
class VBMissingPositional(VBCodeBlock):
    """A positional argument that is missing from the argument list"""

    # << VBMissingPositional methods >> (1 of 2)
    def __init__(self, scope="Private"):
        """Initialize the object"""
        super(VBMissingPositional, self).__init__(scope)
    # << VBMissingPositional methods >> (2 of 2)
    def renderAsCode(self, indent=0):
        """Render this attribute"""
        #
        # The parameter_index_position attribute will be set
        # by our parent. We also need to look for the function name
        # which depends on our context
        try:
            function_name = self.findParentOfClass(VBObject).fnPart()
        except NestingError:
            try:
                function_name = self.getParentProperty("object").fnPart()
            except NestingError:
                raise UnresolvableName("Could not locate function name when supplying missing argument")
        #
        return "VBGetMissingArgument(%s, %d)" % (
                         function_name,
                         self.parameter_index_position)
    # -- end -- << VBMissingPositional methods >>
# << Classes >> (18 of 74)
class VBExpression(VBNamespace):
    """Represents an comment"""

    # << VBExpression methods >> (1 of 3)
    def __init__(self, scope="Private"):
        """Initialize the assignment"""
        super(VBExpression, self).__init__(scope)
        self.parts = []
        self.auto_class_handlers.update({
            "sign"	: (VBExpressionPart, self.parts),
            "pre_not" : (VBExpressionPart, self.parts),
            "par_expression" : (VBParExpression, self.parts),
            "point" : (VBPoint, self.parts),
            "operation" : (VBOperation, self.parts),
            "pre_named_argument" : (VBExpressionPart, self.parts),
            "pre_typeof" : (VBUnrendered, self.parts),
        })
        self.operator_groupings = [] # operators who requested regrouping (eg 'a Like b' -> 'Like(a,b)')
    # << VBExpression methods >> (2 of 3)
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        self.checkForOperatorGroupings()
        return " ".join([item.renderAsCode(indent) for item in self.parts])
    # << VBExpression methods >> (3 of 3)
    def checkForOperatorGroupings(self):
        """Look for operators who requested regrouping

            Some operator cannot be translated in place (eg Like) since they must
            be converted to functions. This means that we have to re-order the 
            parts of the expression.

            """
        for item in self.operator_groupings:
            idx = self.parts.index(item)
            rh, lh = self.parts.pop(idx+1), self.parts.pop(idx-1)
            item.rh, item.lh = rh, lh
    # -- end -- << VBExpression methods >>
# << Classes >> (19 of 74)
class VBParExpression(VBNamespace):
    """A block in an expression"""

    auto_handlers = [
        "l_bracket",
        "r_bracket",
    ]

    # << VBParExpression methods >> (1 of 3)
    def __init__(self, scope="Private"):
        """Initialize"""
        super(VBParExpression, self).__init__(scope)
        self.parts = []
        self.named_argument = ""
        self.auto_class_handlers.update({
            "integer" : (VBExpressionPart, self.parts),
            "hexinteger" : (VBExpressionPart, self.parts),
            "stringliteral" : (VBStringLiteral, self.parts),
            "dateliteral" : (VBDateLiteral, self.parts),
            "floatnumber" : (VBExpressionPart, self.parts),
            "longinteger" : (VBExpressionPart, self.parts),
            "object" : (VBObject, self.parts),
            "par_expression" : (VBParExpression, self.parts),
            "operation" : (VBOperation, self.parts),
            "named_argument" : (VBConsumer, "named_argument"),
            "pre_not" : (VBExpressionPart, self.parts),
            "pre_typeof" : (VBUnrendered, self.parts),
            "point" : (VBPoint, self.parts),
        })

        self.l_bracket = self.r_bracket = ""
        self.operator_groupings = [] # operators who requested regrouping (eg 'a Like b' -> 'Like(a,b)')
    # << VBParExpression methods >> (2 of 3)
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        self.checkForOperatorGroupings()
        if self.named_argument:
            arg = "%s=" % self.named_argument.element.text
        else:
            arg = ""
        ascode = " ".join([item.renderAsCode(indent) for item in self.parts])
        return "%s%s%s%s" % (arg, self.l_bracket, ascode, self.r_bracket)
    # << VBParExpression methods >> (3 of 3)
    def checkForOperatorGroupings(self):
        """Look for operators who requested regrouping

            Some operator cannot be translated in place (eg Like) since they must
            be converted to functions. This means that we have to re-order the 
            parts of the expression.

            """
        # Destructively scan the list so we don't try this a second time later!
        while self.operator_groupings:
            item = self.operator_groupings.pop()
            idx = self.parts.index(item)
            rh, lh = self.parts.pop(idx+1), self.parts.pop(idx-1)
            item.rh, item.lh = rh, lh
    # -- end -- << VBParExpression methods >>
# << Classes >> (20 of 74)
class VBPoint(VBExpression):
    """A block in an expression"""

    skip_handlers = [
        "point",
    ]

    # << VBPoint methods >>
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        return "(%s)" % ", ".join([item.renderAsCode() for item in self.parts])
    # -- end -- << VBPoint methods >>
# << Classes >> (21 of 74)
class VBExpressionPart(VBConsumer):
    """Part of an expression"""

    # << VBExpressionPart methods >> (1 of 2)
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        if self.element.name == "object":
            #
            # Check for implicit object (inside a with)
            if self.element.text.startswith("."):
                return "%s%s" % (self.getParentProperty("with_object"),
                                 self.element.text)
        elif self.element.text.lower() == "like":
            return "Like(%s, %s)" % (self.lh.renderAsCode(), self.rh.renderAsCode())
        elif self.element.name == "pre_named_argument":
            return "%s=" % (self.element.text.split(":=")[0],)
        elif self.element.name == "pre_not":
            self.element.text = "not"
        elif self.element.name == "hexinteger":
            if self.element.text.endswith("&"):
                return "0x%s" % self.element.text[2:-1]
            else:
                return "0x%s" % self.element.text[2:]

        return self.element.text
    # << VBExpressionPart methods >> (2 of 2)
    def finalizeObject(self):
        """Finalize the object

            Check for any type markers.

            """
        ending = self.element.text[-1:] or " "
        if ending in "#$%&":
            log.info("Removed type identifier from '%s'" % self.element.text)
            self.element.text = self.element.text[:-1]
    # -- end -- << VBExpressionPart methods >>
# << Classes >> (22 of 74)
class VBOperation(VBExpressionPart):
    """An operation in an expression"""

    translation = {
        "&" : "+",
        "^" : "**",
        "=" : "==",
        "\\" : "//",  # TODO: Is this right?
        "is" : "is",
        "or" : "or",
        "and" : "and", # TODO: are there any more?
        "xor" : "^",
        "mod" : "%", 
    }

    # << VBOperation methods >> (1 of 2)
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        if self.element.text.lower() in self.translation:
            return self.translation[self.element.text.lower()]
        else:
            return super(VBOperation, self).renderAsCode(indent)
    # << VBOperation methods >> (2 of 2)
    def finalizeObject(self):
        """Finalize the object"""
        if self.element.text.lower() in ("like", ):
            log.info("Found regrouping operator, reversing order of operands")
            self.parent.operator_groupings.append(self)
    # -- end -- << VBOperation methods >>
# << Classes >> (23 of 74)
class VBStringLiteral(VBExpressionPart):
    """Represents a string literal"""

    # << VBStringLiteral methods >>
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        #
        # Remember to replace the double quotes with single ones
        body = self.element.text[1:-1]
        body = body.replace('""', '"')
        #
        if self.checkOptionYesNo("General", "AlwaysUseRawStringLiterals") == "Yes":
            body = body.replace("'", "\'")
            return "r'%s'" % body
        else:		
            body = body.replace('\\', '\\\\')
            body = body.replace("'", "\\'")
            return "'%s'" % body
    # -- end -- << VBStringLiteral methods >>
# << Classes >> (24 of 74)
class VBDateLiteral(VBParExpression):
    """Represents a date literal"""

    skip_handlers = [
        "dateliteral",
    ]

    # << VBDateLiteral methods >>
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        return "MakeDate(%s)" % ", ".join([item.renderAsCode() for item in self.parts])
    # -- end -- << VBDateLiteral methods >>
# << Classes >> (25 of 74)
class VBProject(VBNamespace):
    """Handles a VB Project"""

    # << VBProject methods >> (1 of 2)
    def __init__(self, scope="Private"):
        """Initialize the module"""
        super(VBProject, self).__init__(scope)
        self.global_objects = {} # This is where global variables live
    # << VBProject methods >> (2 of 2)
    def resolveLocalName(self, name, rendering_locals=0, requestedby=None):
        """Convert a local name to a fully resolved name

            We search our local modules to see if they have a matching global variable
            and if they do then we can construct the local name from it.

            """
        #import pdb; pdb.set_trace()
        if name in self.global_objects:
            # Found as another module's public var - so mark it up and request an import
            modulename = self.global_objects[name].getParentProperty("modulename")
            if requestedby:
                requestedby.registerImportRequired(modulename)
            return "%s.%s" % (modulename,
                              name)
        else:
            raise UnresolvableName("Name '%s' is not known in this namespace" % name)
    # -- end -- << VBProject methods >>
# << Classes >> (26 of 74)
class VBModule(VBCodeBlock):
    """Handles a VB Module"""

    skip_handlers = [
    ]

    convert_functions_to_methods = 0  # If this is 1 then local functions will become methods
    indent_all_blocks = 0
    allow_new_style_class = 1 # Can be used to dissallow new style classes

    public_is_global = 0 # Public objects defined here will not be globals

    # Put methods and attribute names in here which always need to be public
    # like Class_Initialize and Class_Terminate for classes
    always_public_attributes = []

    # << VBModule methods >> (1 of 11)
    def __init__(self, scope="Private", modulename="unknownmodule", classname="MyClass",
                 superclasses=None):
        """Initialize the module"""
        super(VBModule, self).__init__(scope)
        self.auto_class_handlers.update({
            "sub_definition" : (VBSubroutine, self.locals),
            "fn_definition" : (VBFunction, self.locals),
            "property_definition" : (VBProperty, self.locals),
            "enumeration_definition" : (VBEnum, self.locals),
        })
        self.local_names = []
        self.modulename = modulename
        self.classname = classname
        self.superclasses = superclasses or []
        #
        self.rendering_locals = 0
        self.docstrings = []
        self.module_imports = [] # The additional modules we need to import
    # << VBModule methods >> (2 of 11)
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        self.setCustomModulesAsGlobals()
        if self.checkOptionYesNo("General", "TryToExtractDocStrings") == "Yes":
            self.extractDocStrings()
        #
        # Pre-render the following before the import statments in case any
        # of them ask us to do additional imports
        header = self.renderModuleHeader(indent)
        docstrings = self.renderDocStrings(indent)
        declarations = self.renderDeclarations(indent+self.indent_all_blocks)
        blocks = self.renderBlocks(indent+self.indent_all_blocks)
        #
        return "%s\n\n%s%s\n%s\n%s" % (
                       self.importStatements(indent),
                       header,
                       docstrings,
                       declarations,
                       blocks,
                      )
    # << VBModule methods >> (3 of 11)
    def importStatements(self, indent=0):
        """Render the standard import statements for this block"""
        other = [""]+["import %s" % item for item in self.module_imports] # Leading [""] gives a newline
        if self.checkOptionYesNo("General", "IncludeDebugCode") == "Yes":
            debug = "\nfrom vb2py.vbdebug import *"
        else:
            debug = ""
        return "from vb2py.vbfunctions import *%s%s" % (debug, "\n".join(other))
    # << VBModule methods >> (4 of 11)
    def renderDeclarations(self, indent):
        """Render the declarations as code

            Most of the rendering is delegated to the individual declaration classes. However,
            we cannot do this with properties since they need to be grouped into a single assignment.
            We do the grouping here and delegate the rendering to them.

            """
        #
        ret = []
        self.rendering_locals = 1 # Used for switching behaviour (eg adding 'self')
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
        self.rendering_locals = 0
        #
        return "".join(ret)
    # << VBModule methods >> (5 of 11)
    def renderBlocks(self, indent=0):
        """Render this module's blocks"""
        return "".join([block.renderAsCode(indent) for block in self.blocks])
    # << VBModule methods >> (6 of 11)
    def extractDocStrings(self, indent=0):
        """Extract doc strings from this module

            We look for comments in the body of the module and take all the ones before
            anything that isn't a comment.

            """
        for line in self.blocks[:]:
            if isinstance(line, VBComment):
                self.docstrings.append(line)
                self.blocks.remove(line)
            elif line.would_end_docstring:
                break
    # << VBModule methods >> (7 of 11)
    def renderDocStrings(self, indent=0):
        """Render this module's docstrings"""
        local_indent = indent+self.indent_all_blocks
        if not self.docstrings:
            return ""
        elif len(self.docstrings) == 1:
            return '%s"""%s"""\n'  % (
                    self.getIndent(local_indent),
                    self.docstrings[0].asString())
        else:
            joiner = "\n%s" % self.getIndent(local_indent)
            return '%s"""%s\n%s%s\n%s"""\n' % (
                    self.getIndent(local_indent),
                    self.docstrings[0].asString(),
                    self.getIndent(local_indent),
                    joiner.join([item.asString() for item in self.docstrings[1:]]),
                    self.getIndent(local_indent),
                    )
    # << VBModule methods >> (8 of 11)
    def renderModuleHeader(self, indent=0):
        """Render a header for the module"""
        return ""
    # << VBModule methods >> (9 of 11)
    def resolveLocalName(self, name, rendering_locals=0, requestedby=None):
        """Convert a local name to a fully resolved name

            We search our local variables to see if we know the name. If we do then we
            just report it.

            """
        if name in self.local_names:
            return name
        for obj in self.locals:
            if obj.identifier == name:
                return self.enforcePrivateName(obj)
        raise UnresolvableName("Name '%s' is not known in this namespace" % name)
    # << VBModule methods >> (10 of 11)
    def enforcePrivateName(self, obj):
        """Enforce the privacy for this object name if required"""
        if obj.scope == "Private" and self.checkOptionYesNo("General", "RespectPrivateStatus") == "Yes" \
           and not obj.identifier in self.always_public_attributes:
            return "%s%s" % (Config["General", "PrivateDataPrefix"], obj.identifier)
        else:
            return obj.identifier
    # << VBModule methods >> (11 of 11)
    def setCustomModulesAsGlobals(self): 
        """Set all the custom import modules as global modules

            If the user has specified custom imports (eg Comctllib) then
            we need to import these as globals in the project. We force
            them into the project (if there is one) global object
            table so that they can be resolved at run time.

            """
        #
        # Get global object table if there is one
        try:
            global_objects = self.getParentProperty("global_objects")
        except NestingError:
            return
        #
        log.info("Processing custom modules now")
        custom_modules = Config.getItemNames("CustomIncludes")
        #
        # Do for all custom modules
        for module_id in custom_modules:
            #
            # Import this module
            module_name = Config["CustomIncludes", module_id]
            log.info("Processing custom module %s (%s)" % (module_id, module_name))
            module = __import__("vb2py.custom.%s" % module_name, globals(), locals(), ["custom"])
            #
            # Get a container to store the values in
            vbmodule = VBCodeModule(modulename="vb2py.custom.%s" % module_name)
            #
            # Now set all items in the module to be global (if they don't seem to be
            # hidden)
            for item_name in dir(module):
                if not item_name.startswith("_"):
                    log.info("Registered new custom global '%s'" % item_name)
                    global_objects[item_name] = vbmodule
    # -- end -- << VBModule methods >>
# << Classes >> (27 of 74)
class VBClassModule(VBModule):
    """Handles a VB Class"""

    convert_functions_to_methods = 1  # If this is 1 then local functions will become methods
    indent_all_blocks = 1

    # Put methods and attribute names in here which always need to be public
    # like Class_Initialize and Class_Terminate for classes
    always_public_attributes = ["Class_Initialize", "Class_Terminate"]

    # << VBClassModule methods >> (1 of 4)
    def __init__(self, *args, **kw):
        """Initialize the class module"""
        super(VBClassModule, self).__init__(*args, **kw)
        self.name_substitution = {"Me" : "self"}
    # << VBClassModule methods >> (2 of 4)
    def renderModuleHeader(self, indent=0):
        """Render this element as code"""
        supers = self.superclasses[:]
        if self.checkOptionYesNo("Classes", "UseNewStyleClasses") == "Yes" and \
                 self.allow_new_style_class:
            supers.insert(0, "Object")
        if supers:
            return "class %s(%s):\n" % (self.classname, ", ".join(supers))
        else:        
            return "class %s:\n" % self.classname
    # << VBClassModule methods >> (3 of 4)
    def resolveLocalName(self, name, rendering_locals=0, requestedby=None):
        """Convert a local name to a fully resolved name

            We search our local variables to see if we know the name. If we do then we
            need to add a self.

            """
        # Don't do anything for locals
        if rendering_locals:
            prefix = ""
        else:
            prefix = "self."
        #
        if name in self.local_names:
            return "%s%s" % (prefix, name)
        for obj in self.locals:
            if obj.identifier == name:
                return "%s%s" % (prefix, self.enforcePrivateName(obj))
        raise UnresolvableName("Name '%s' is not known in this namespace" % name)
    # << VBClassModule methods >> (4 of 4)
    def assignParent(self, parent):
        """Set our parent"""
        super(VBClassModule, self).assignParent(parent)
        self.identifier = self.classname 
        self.registerAsGlobal()
    # -- end -- << VBClassModule methods >>
# << Classes >> (28 of 74)
class VBCodeModule(VBModule):
    """Handles a VB Code module"""

    public_is_global = 1 # Public objects defined here will be globals

    # << VBCodeModule methods >>
    def enforcePrivateName(self, obj):
        """Enforce the privacy for this object name if required

            In a code module this is not required. Private variables and definitions in a code
            module are not really hidden in the same way as in a class module. They are accessible
            still. The main thing is that they are not global.

            """
        return obj.identifier
    # -- end -- << VBCodeModule methods >>
# << Classes >> (29 of 74)
class VBFormModule(VBClassModule):
    """Handles a VB Form module"""

    convert_functions_to_methods = 1  # If this is 1 then local functions will become methods
# << Classes >> (30 of 74)
class VBCOMExternalModule(VBModule):
    """Handles external COM references"""



    # << VBCOMExternalModule methods >> (1 of 2)
    def __init__(self, *args, **kw):
        """Initialize the COM module

        We always need win32com.client to be imported

        """
        super(VBCOMExternalModule, self).__init__(*args, **kw)
        self.module_imports.append("win32com.client")
        self.docstrings.append(
                VBRenderDirect("Automatically generated file based on project references"))
    # << VBCOMExternalModule methods >> (2 of 2)
    def renderDeclarations(self, indent):
        """Render all the declarations

        We have a list of libraries and objects in our names attribute
        so we create a series of dummy classes with callable
        attributes which return COM objects.

        """
        library_code = []
        for library, members in self.names.iteritems():
            member_code = []
            for member in members:
                member_code.append(
                        '    def %s(self):\n'
                        '        """Create the %s.%s object"""\n'
                        '        return win32com.client.Dispatch("%s.%s")\n'
                        '\n' % (member, library, member, library, member))
            library_code.append('class _%s:\n'
                                '    """COM Library"""\n'
                                '\n%s%s = _%s()\n' % (library, ''.join(member_code), library, library))

        return '\n\n'.join(library_code)
    # -- end -- << VBCOMExternalModule methods >>
# << Classes >> (31 of 74)
class VBVariableDefinition(VBVariable):
    """Handles a VB Dim of a Variable"""

    # << VBVariableDefinition methods >>
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        #
        local_name = self.resolveName(self.identifier)
        #
        # TODO: Can't handle implicit objects yet
        if self.implicit_object:
            warning = self.getWarning(
                    "UnhandledDefinition", 
                    "Dim of implicit 'With' object (%s) is not supported" % local_name, 
                    indent=indent, crlf=1)
        else:
            warning = ""
        #
        if self.string_size_indicator:
            size = self.string_size_indicator
            self.type = "FixedString"
        else:
            size = ""
        #
        # Make sure we resolve the type properly
        local_type = self.resolveName(self.type)
        #
        if self.unsized_definition: # This is a 'Dim a()' statement
            return "%s%s%s = vbObjectInitialize(objtype=%s)\n" % (
                            warning,
                            self.getIndent(indent),
                            local_name,
                            local_type)						
        elif self.size_definitions: # There is a size 'Dim a(10)'
            if self.preserve_keyword:
                preserve = ", %s" % (local_name, )
            else:
                preserve = ""
            if size:
                size = ", stringsize=" + size
            return "%s%s%s = vbObjectInitialize((%s,), %s%s%s)\n" % (
                            warning,
                            self.getIndent(indent),
                            local_name,
                            ", ".join([item.renderAsCode() for item in self.size_definitions]),
                            local_type,
                            preserve,
                            size)
        elif self.new_keyword: # It is an 'Dim a as new ...'
            return "%s%s%s = %s(%s)\n" % (
                            warning,
                            self.getIndent(indent),
                            local_name,
                            local_type,
                            size)
        else: # This is just 'Dim a as frob'
            return "%s%s%s = %s(%s)\n" % (
                            warning,
                            self.getIndent(indent),
                            local_name,
                            local_type,
                            size)
    # -- end -- << VBVariableDefinition methods >>
# << Classes >> (32 of 74)
class VBConstant(VBVariableDefinition):
    """Represents a constant in VB"""

    # << VBConstant methods >>
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        #local_name = self.getLocalNameFor(self.identifier)
        local_name = self.resolveName(self.identifier)
        return "%s%s = %s\n" % (
                            self.getIndent(indent),
                            local_name,
                            self.expression.renderAsCode())
    # -- end -- << VBConstant methods >>
# << Classes >> (33 of 74)
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
# << Classes >> (34 of 74)
class VBAssignment(VBNamespace):
    """An assignment statement"""

    auto_handlers = [
    ]

    # << VBAssignment methods >> (1 of 4)
    def __init__(self, scope="Private"):
        """Initialize the assignment"""
        super(VBAssignment, self).__init__(scope)
        self.parts = []
        self.object = None
        self.auto_class_handlers.update({
            "expression" : (VBExpression, self.parts),
            "object" : (VBLHSObject, "object")
        })
    # << VBAssignment methods >> (2 of 4)
    def asString(self):
        """Convert to a nice representation"""
        return "%s = %s" % (self.object, self.parts)
    # << VBAssignment methods >> (3 of 4)
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        self.checkForModuleGlobals()
        self.object.brackets_are_indexes = 1 # Convert brackets on LHS to []
        return "%s%s = %s\n" % (self.getIndent(indent),
                                self.object.renderAsCode(), 
                                self.parts[0].renderAsCode(indent))
    # << VBAssignment methods >> (4 of 4)
    def checkForModuleGlobals(self):
        """Check if this assignment requires a global statement

            We can use this opportunity to now check if we need to append a 'global' statement
            to our container. If we are in a CodeModule an assignment and the LHS of the assignment is a
            module level variable which is not locally shadowed then we need a global.

            So the procedure is,
             - look for our parent who is a subroutine type
             - if we don't have one then skip out
             - see if this parent knows us, if so then we are a subroutine local
             - also see if we are the subroutine name
             - look for our parent who is a module type
             - see if this parent knows us, if so then we are a module local
             - if we are then tell our subroutine parent that we need a global statement

            """        
        log.info("Checking whether to use a global statement for '%s'" % self.object.primary.element.text)
        #import pdb; pdb.set_trace()
        try:
            enclosing_sub = self.findParentOfClass(VBSubroutine)
        except NestingError:
            return # We are not in a subroutine

        log.info("Found sub")    
        try:
            name = enclosing_sub.resolveLocalName(self.object.primary.element.text)
        except UnresolvableName:
            if enclosing_sub.identifier == self.object.primary.element.text:
                return
        else:
            return # We are a subroutine local

        log.info("Am not local")    
        try:
            enclosing_module = self.findParentOfClass(VBCodeModule)
        except NestingError:
            return # We are not in a module

        log.info("Found code module")        
        try:
            name = enclosing_module.resolveLocalName(self.object.primary.element.text)
        except UnresolvableName:
            return # We are not known at the module level

        # If we get to here then we are a module level local!
        enclosing_sub.globals_required[self.resolveName(self.object.primary.element.text)] = 1

        log.info("Added a module level global: '%s'" % self.resolveName(self.object.primary.element.text))
    # -- end -- << VBAssignment methods >>
# << Classes >> (35 of 74)
class VBSpecialAssignment(VBAssignment):
    """A special assignment eg LSet, RSet where the assignment ends up as a function call"""

    fn_name = None

    # << VBSpecialAssignment methods >>
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        self.checkForModuleGlobals()
        self.object.brackets_are_indexes = 1 # Convert brackets on LHS to []
        return "%s%s = %s(%s, %s)\n" % (self.getIndent(indent),
                                self.object.renderAsCode(), 
                                self.fn_name, 
                                self.object.renderAsCode(), 
                                self.parts[0].renderAsCode(indent))
    # -- end -- << VBSpecialAssignment methods >>
# << Classes >> (36 of 74)
class VBLSet(VBSpecialAssignment):
    """An LSet statement"""

    fn_name = "LSet"
# << Classes >> (37 of 74)
class VBRSet(VBSpecialAssignment):
    """An RSet statement"""

    fn_name = "RSet"
# << Classes >> (38 of 74)
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
# << Classes >> (39 of 74)
class VBEnd(VBAssignment):
    """An end statement"""

    # << VBEnd methods >>
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        return "%ssys.exit(0)\n" % self.getIndent(indent)
    # -- end -- << VBEnd methods >>
# << Classes >> (40 of 74)
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
            "missing_positional" : (VBMissingPositional, self.parameters),
            "object" : (VBObject, "object")
        })
    # << VBCall methods >> (2 of 2)
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        if self.parameters:
            #	
            # Construct the list of parameters - this is harder than it looks because 
            # for any missing positional parameters we have to do some introspection
            # to dig out the default value
            param_list = []
            for idx, element in zip(xrange(1000), self.parameters):
                element.parameter_index_position = idx # Needed so that the element can get its default
                param_list.append(element.renderAsCode())
            params = ", ".join(param_list)
        else:
            params = ""
        #
        self.object.am_on_lhs = 1
        #
        return "%s%s(%s)\n" % (self.getIndent(indent),
                             self.object.renderAsCode(), 
                             params)
    # -- end -- << VBCall methods >>
# << Classes >> (41 of 74)
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
        elif self.element.text == "Exit Property":
            if self.getParentProperty("property_decorator_type") == "Get":
                return "%sreturn %s\n" % (indenter, Config["Functions", "ReturnVariableName"])
            else:
                return "%sreturn\n" % indenter		
        else:
            return "%sbreak\n" % indenter
    # -- end -- << VBExitStatement methods >>
# << Classes >> (42 of 74)
class VBComment(VBConsumer):
    """Represents an comment"""

    #
    # Used to indicate if this is a valid statement
    not_a_statement = 0

    # << VBComment methods >> (1 of 2)
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        return self.getIndent(indent) + "#%s\n" % self.element.text
    # << VBComment methods >> (2 of 2)
    def asString(self):
        """Render this element as a string"""
        return self.element.text
    # -- end -- << VBComment methods >>
# << Classes >> (43 of 74)
class VBLabel(VBUnrendered):
    """Represents a label"""

    def renderAsCode(self, indent):
        """Render the label"""
        if Config["Labels", "IgnoreLabels"] == "Yes":
            return ""
        else:
            return super(VBLabel, self).renderAsCode(indent)
# << Classes >> (44 of 74)
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
# << Classes >> (45 of 74)
class VBClose(VBCodeBlock):
    """Represents a close statement"""

    # << VBClose methods >> (1 of 2)
    def __init__(self, scope="Private"):
        """Initialize the open"""
        super(VBClose, self).__init__(scope)
        #
        self.channels = []
        #
        self.auto_class_handlers = ({
            "expression" : (VBParExpression, self.channels),
        })
    # << VBClose methods >> (2 of 2)
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        if not self.channels:
            return "%sVBFiles.closeFile()\n" % (
                        self.getIndent(indent))
        else:
            ret = []
            for channel in self.channels:
                ret.append("%sVBFiles.closeFile(%s)\n" % (
                        self.getIndent(indent),
                        channel.renderAsCode()))
            return "".join(ret)
    # -- end -- << VBClose methods >>
# << Classes >> (46 of 74)
class VBSeek(VBCodeBlock):
    """Represents a seek statement"""

    # << VBSeek methods >> (1 of 2)
    def __init__(self, scope="Private"):
        """Initialize the seek"""
        super(VBSeek, self).__init__(scope)
        #
        self.expressions = []
        #
        self.auto_class_handlers = ({
            "expression" : (VBParExpression, self.expressions),
        })
    # << VBSeek methods >> (2 of 2)
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        return "%sVBFiles.seekFile(%s, %s)\n" % (
                    self.getIndent(indent),
                    self.expressions[0].renderAsCode(),
                    self.expressions[1].renderAsCode(),)
    # -- end -- << VBSeek methods >>
# << Classes >> (47 of 74)
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
        # Make sure variables are converted as if they are on the LHS of an assignment
        for var in self.variables:
            var.brackets_are_indexes = 1
        #
        return "%s%s = VBFiles.get%s(%s, %d)\n" % (
                    self.getIndent(indent),
                    ", ".join([var.renderAsCode() for var in self.variables]),
                    self.input_type,
                    self.channel.renderAsCode(),
                    len(self.variables))
    # -- end -- << VBInput methods >>
# << Classes >> (48 of 74)
class VBLineInput(VBInput):
    """Represents an input statement"""

    input_type = "LineInput"
# << Classes >> (49 of 74)
class VBPrint(VBCodeBlock):
    """Represents a print statement"""

    # << VBPrint methods >> (1 of 2)
    def __init__(self, scope="Private"):
        """Initialize the print"""
        super(VBPrint, self).__init__(scope)
        #
        self.channel = VBRenderDirect("None")
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
# << Classes >> (50 of 74)
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
# << Classes >> (51 of 74)
class VBName(VBCodeBlock):
    """Represents a name statement"""

    # << VBName methods >> (1 of 2)
    def __init__(self, scope="Private"):
        """Initialize the print"""
        super(VBName, self).__init__(scope)
        #
        self.channel = VBRenderDirect("None")
        self.files = []
        #
        self.auto_class_handlers = ({
            "expression" : (VBExpression, self.files),
        })
    # << VBName methods >> (2 of 2)
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        self.registerImportRequired("os")
        file_list = ", ".join([fle.renderAsCode() for fle in self.files])
        return "%sos.rename(%s)\n" % (
                    self.getIndent(indent),
                    file_list)
    # -- end -- << VBName methods >>
# << Classes >> (52 of 74)
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
        if not self.variables:
            vars.append(VBPass().renderAsCode(indent+2))
        else:
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
# << Classes >> (53 of 74)
class VBSubroutine(VBCodeBlock):
    """Represents a subroutine"""

    public_is_global = 0 # Public objects defined here will not be globals

    # << VBSubroutine methods >> (1 of 6)
    def __init__(self, scope="Private"):
        """Initialize the subroutine"""
        super(VBSubroutine, self).__init__(scope)
        self.identifier = None
        self.scope = scope
        self.block = VBPass()
        self.parameters = []
        self.globals_required = {} # A list of objects required in a global statement
        self.type = None
        self.static = None
        #
        self.auto_class_handlers.update({
            "formal_param" : (VBVariable, self.parameters),
            "block" : (VBCodeBlock, "block"),
            "type_definition" : (VBUnrendered, "type"),
        })

        self.auto_handlers = [
                "identifier",
                "scope",
                "static",
        ]

        self.skip_handlers = [
                "sub_definition",
        ]

        self.rendering_locals = 0
    # << VBSubroutine methods >> (2 of 6)
    def renderAsCode(self, indent=0):
        """Render this subroutine"""
        code_block = self.block.renderAsCode(indent+1)
        locals = [declaration.renderAsCode(indent+1) for declaration in self.block.locals]
        if self.static:
            log.warn("Static function detected - static is not supported")
        ret = "\n%sdef %s(%s):\n%s%s%s" % (
                    self.getIndent(indent),
                    self.getParentProperty("enforcePrivateName")(self),
                    self.renderParameters(),
                    self.renderGlobalStatement(indent+1),
                    "\n".join(locals),
                    code_block)
        return ret
    # << VBSubroutine methods >> (3 of 6)
    def renderParameters(self):
        """Render the parameter list"""
        params = [param.renderAsCode() for param in self.parameters]
        if self.getParentProperty("convert_functions_to_methods"):
            params.insert(0, "self")
        return ", ".join(params)
    # << VBSubroutine methods >> (4 of 6)
    def resolveLocalName(self, name, rendering_locals=0, requestedby=None):
        """Convert a local name to a fully resolved name

            We search our local variables and parameters to see if we know the name. If we do then we
            return the original name.

            """
        names = [obj.identifier for obj in self.block.locals + self.parameters]
        if name in names:
            return name
        else:
            raise UnresolvableName("Name '%s' is not known in this namespace" % name)
    # << VBSubroutine methods >> (5 of 6)
    def renderGlobalStatement(self, indent=0):
        """Render the global statement if we need it"""
        if self.globals_required:
            return "%sglobal %s\n" % (self.getIndent(indent),
                                      ", ".join(self.globals_required.keys()))
        else:
            return ""
    # << VBSubroutine methods >> (6 of 6)
    def assignParent(self, *args, **kw):
        """Assign our parent

            We can use this opportunity to now determine if we are a global

            """
        super(VBSubroutine, self).assignParent(*args, **kw)
        #
        # Check if we will be considered a global for the project
        if hasattr(self, "parent"):
            if self.parent.amGlobal(self.scope):
                self.registerAsGlobal()
    # -- end -- << VBSubroutine methods >>
# << Classes >> (54 of 74)
class VBFunction(VBSubroutine):
    """Represents a function"""

    is_function = 1 # We need () if we are accessed directly

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
        locals = [declaration.renderAsCode(indent+1) for declaration in self.block.locals]
        #
        if Config["Functions", "PreInitializeReturnVariable"] == "Yes":
            pre_init = "%s%s = None\n" % (				
                    self.getIndent(indent+1),
                    return_var)
        else:
            pre_init = ""

        ret = "\n%sdef %s(%s):\n%s%s%s%s%sreturn %s\n" % (
                    self.getIndent(indent),
                    self.getParentProperty("enforcePrivateName")(self), 
                    self.renderParameters(),
                    self.renderGlobalStatement(indent+1),
                    pre_init,
                    "\n".join(locals),
                    block,
                    self.getIndent(indent+1),
                    return_var)
        return ret
    # -- end -- << VBFunction methods >>
# << Classes >> (55 of 74)
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
# << Classes >> (56 of 74)
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
# << Classes >> (57 of 74)
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
            "inline_implicit_call" : (VBCodeBlock, self.statements),  # TODO: remove me
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
# << Classes >> (58 of 74)
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
        self.comment_block = VBNothing()
        #
        self.auto_class_handlers = {
            "expression" : (VBExpression, "expression"),
            "case_item_block" : (VBCaseItem, self.blocks),
            "case_else_block" : (VBCaseElse, self.blocks),
            "case_comment_block" : (VBOptionalCodeBlock, "comment_block"),
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
        ret += self.comment_block.renderAsCode()
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
# << Classes >> (59 of 74)
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
# << Classes >> (60 of 74)
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
            # << Handle single expression >>
            expression_text = self.expressions[0].renderAsCode()
            # Now check for "Is"
            if expression_text.startswith("Is "):
                # This has "Is" - replace it and use the rest of the expression
                return "%s %s" % (
                                   self.getSelectVariable(),
                                   expression_text[3:])
            else:	
                # Standard case
                return "%s == %s" % (
                                   self.getSelectVariable(),
                                   expression_text)
            # -- end -- << Handle single expression >>
        elif len(self.expressions) == 2:
            return "%s <= %s <= %s" % (
                                           self.expressions[0].renderAsCode(),
                                           self.getSelectVariable(),
                                           self.expressions[1].renderAsCode())
        raise VBParserError("Error rendering case item")
    # -- end -- << VBCaseItem methods >>
# << Classes >> (61 of 74)
class VBCaseElse(VBCaseBlock):
    """Represents a select block"""

    # << VBCaseElse methods >>
    def renderAsCode(self, indent=0):
        """Render this element as code"""
        return "%selse:\n%s" % (self.getIndent(indent),
                                 self.block.renderAsCode(indent+1))
    # -- end -- << VBCaseElse methods >>
# << Classes >> (62 of 74)
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
            "block" : (VBCodeBlock, "block"), # Used for full 'for'
            "body" : (VBCodeBlock, "block"),  # Used for inline 'for'
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
# << Classes >> (63 of 74)
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
# << Classes >> (64 of 74)
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
# << Classes >> (65 of 74)
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
# << Classes >> (66 of 74)
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
        #
        # Don't even do anything if there is no body to the With
        if self.block:
            #
            # Before we render the expression we change its parent to our parent because
            # we don't want any ".implicit" objects to be evaluated using our With object
            self.expression.parent = self.parent
            #
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
        else:
            return ""
    # -- end -- << VBWith methods >>
# << Classes >> (67 of 74)
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

        #
        # Get the name for this property - respecting the hidden status
        obj = pset or pget # Need at least one!
        proper_name = self.getParentProperty("enforcePrivateName")(obj)

        if pset:
            self.getParentProperty("local_names").append(pset.identifier) # Store property name for namespace analysis
            pset.identifier = "%s%s" % (Config["Properties", "LetSetVariablePrefix"], pset.identifier)		
            ret.append(pset.renderAsCode(indent))
            params.append("fset=%s" % self.getParentProperty("enforcePrivateName")(pset))
        if pget:
            self.getParentProperty("local_names").append(pget.identifier) # Store property name for namespace analysis
            pget.__class__ = VBFunction # Needs to be a function
            pget.name_substitution[pget.identifier] = Config["Functions", "ReturnVariableName"]
            pget.identifier = "%s%s" % (Config["Properties", "GetVariablePrefix"], pget.identifier)		
            ret.append(pget.renderAsCode(indent))
            params.append("fget=%s" % self.getParentProperty("enforcePrivateName")(pget))

        return "\n%s%s%s = property(%s)\n" % (
                    "".join(ret),
                    self.getIndent(indent),
                    proper_name,
                    ", ".join(params))
    # -- end -- << VBProperty methods >>
# << Classes >> (68 of 74)
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
                "enumeration_item" : (VBEnumItem, self.enumerations),
            }

        self.auto_handlers = ["identifier"]
    # << VBEnum methods >> (2 of 2)
    def renderAsCode(self, indent=0):
        """Render a group of property statements"""
        count = 0
        ret = []
        for enumeration in self.enumerations:
            if enumeration.expression:
                cnt = enumeration.expression.renderAsCode()
            else:
                cnt = count
                count += 1
            ret.append("%s%s = %s" % (self.getIndent(indent),
                                      enumeration.identifier.element.text,
                                      cnt))

        return "%s# Enumeration '%s'\n%s\n" % (
                            self.getIndent(indent),
                            self.identifier,
                            "\n".join(ret),
                    )
    # -- end -- << VBEnum methods >>
# << Classes >> (69 of 74)
class VBEnumItem(VBCodeBlock):
    """Represents an enum item"""

    # << VBEnumItem methods >>
    def __init__(self, scope="Private"):
        """Initialize the Select"""
        super(VBEnumItem, self).__init__(scope)
        self.identifier = None
        self.expression = None
        #
        self.auto_class_handlers = {
                "identifier" : (VBConsumer, "identifier"),
                "expression" : (VBExpression, "expression"),
            }
    # -- end -- << VBEnumItem methods >>
# << Classes >> (70 of 74)
class VB2PYDirective(VBCodeBlock):
    """Handles a vb2py directive"""

    skip_handlers = [
            "vb2py_directive",
    ]

    would_end_docstring = 0

    # << VB2PYDirective methods >> (1 of 3)
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
    # << VB2PYDirective methods >> (2 of 3)
    def renderAsCode(self, indent=0):
        """We use the rendering to do our stuff"""
        if self.directive_type == "Set":
            Config.setLocalOveride(self.config_section, self.config_name, self.expression)
            log.info("Doing a set: %s" % str((self.config_section, self.config_name, self.expression)))
        elif self.directive_type == "Unset":
            Config.removeLocalOveride(self.config_section, self.config_name)
            log.info("Doing an uset: %s" % str((self.config_section, self.config_name)))
        elif self.directive_type in ("GlobalSet", "GlobalAdd"):
            pass # already handled this
        elif self.directive_type == "Add":
            Config.addLocalOveride(self.config_section, self.config_name, self.expression)
            log.info("Adding a setting: %s" % str((self.config_section, self.config_name, self.expression)))
        else:
            raise DirectiveError("Directive not understood: '%s'" % self.directive_type)
        return ""
    # << VB2PYDirective methods >> (3 of 3)
    def assignParent(self, *args, **kw):
        """Assign our parent

            We can use this opportunity to now determine if we are a global

            """
        super(VB2PYDirective, self).assignParent(*args, **kw)
        #
        # Check if we are a global level option - if se we set it now
        if self.directive_type == "GlobalSet":
            Config.setLocalOveride(self.config_section, self.config_name, self.expression)
        elif self.directive_type == "GlobalAdd":
            Config.addLocalOveride(self.config_section, self.config_name, self.expression)
    # -- end -- << VB2PYDirective methods >>
# << Classes >> (71 of 74)
class VBPass(VBCodeBlock):
    """Represents an empty statement"""

    def renderAsCode(self, indent=0):
        """Render it!"""
        return "%spass\n" % (self.getIndent(indent),)
# << Classes >> (72 of 74)
class VBRenderDirect(VBCodeBlock):
    """Represents a pre-rendered statement"""

    def __init__(self, text, indent=0, crlf=0):
        """Initialize"""
        super(VBRenderDirect, self).__init__()
        self.identifier = text
        self.indent = indent
        self.crlf = crlf

    def renderAsCode(self, indent=0):
        """Render it!"""
        s = ""
        if self.indent:
            s += self.getIndent(indent)
        s += self.identifier
        if self.crlf:
            s += "\n"
        return s

    def asString(self):
        """Return string representation"""
        return self.identifier
# << Classes >> (73 of 74)
class VBNothing(VBCodeBlock):
    """Represents a block which renders to nothing at all"""

    def renderAsCode(self, indent=0):
        """Render it!"""
        return ""
# << Classes >> (74 of 74)
class VBParserFailure(VBConsumer):
    """Represents a block which failed to parse"""

    def renderAsCode(self, indent=0):
        """Render it!"""
        fail_option = Config["General", "InsertIntoFailedCode"].lower()
        warn = self.getWarning("ParserError", self.element.text, indent, crlf=1) + \
               self.getWarning("ParserStop", "Conversion of VB code halted", indent, crlf=1)
        if fail_option == "exception":
            warn += "%sraise NotImplemented('VB2PY Code conversion failed at this point')" % self.getIndent(indent)
        elif fail_option == "warning":
            warn += "%simport warnings;warnings.warn('VB2PY Code conversion failed at this point')" % self.getIndent(indent)
        #
        return warn
# -- end -- << Classes >>

from vbparser import *

# Blocks which do not contain valid statements
# If a block contains only these then it needs a pass
# statement to be a valid Python suite
NonCodeBlocks = (VBComment, VBUnrendered, VB2PYDirective)
