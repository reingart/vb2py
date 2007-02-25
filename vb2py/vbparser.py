"""Main parsing and conversion routines for translating VB to Python code"""

# << Imports >>
#
# Configuration options
import config
Config = config.VB2PYConfig()

from pprint import pprint as pp
from simpleparse.common import chartypes
import sys
import os
import re
import utils

declaration = open(utils.relativePath("vbgrammar.txt"), "r").read()

from simpleparse.parser import Parser

import logger
log = logger.getLogger("VBParser")
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

# << Utility functions >> (1 of 10)
def convertToElements(details, txt):
    """Convert a parse tree to elements"""
    ret = []
    if details:
        for item in details:
            ret.append(VBElement(item, txt))
    return ret
# << Utility functions >> (2 of 10)
def buildParseTree(vbtext, starttoken="line", verbose=0, returnpartial=0, returnast=0):
    """Parse some VB"""

    # << Build Parser >>
    # Try to buid the parse - if this fails we probably have an early
    # version of Simpleparse
    try:
        parser = Parser(declaration, starttoken)
    except Exception, err:
        log.warn("Failed to build parse (%s) - trying case sensitive grammar" % err)
        parser = Parser(declaration.replace('c"', ' "'), starttoken)
        log.info("Downgraded to case sensitive grammar")
    # -- end -- << Build Parser >>

    txt = applyPlugins("preProcessVBText", vbtext)

    txt = makeSafeFromUnicode(txt)

    nodes = []
    while 1:
        success, tree, next = parser.parse(txt) 
        if not success:
            if txt.strip():
                # << Handle failure >>
                msg = "Parsing error: %d, '%s'" % (next, txt.split("\n")[0])
                if returnpartial:
                    log.error(msg)
                    nodes.append(VBFailedElement('parser_failure', msg))
                    break
                else:
                    raise VBParserError(msg)
                # -- end -- << Handle failure >>
            break
        if verbose:
            print success, next
            pp(tree)
            print "."
        if not returnast:
            nodes.extend(convertToElements(tree, txt))
        else:
            nodes.append(tree)
        txt = txt[next:]

    return nodes
# << Utility functions >> (3 of 10)
def makeSafeFromUnicode(text):
    """Return a safe version of the text without unicode characters

    We do some ugly hacks here to handle unicode since the SimpleParse library
    doesn't have an easy way to deal with it.

    """
    result = []
    letters = map(ord, text)
    marker1 = [ord('x'), ord('X')]
    marker2 = [ord('X'), ord('x')]    
    #
    # Replace all non asc characters with a marker
    for letter in letters:
        if letter < 128:
            result.append(letter)
        else:
            result.extend(marker1)
            result.extend(map(ord, str(letter)))
            result.extend(marker2)
    #
    return "".join(map(chr, result))
# << Utility functions >> (4 of 10)
def makeUnicodeFromSafe(text):
    """Recover the unicode text from a safe version of the text

    We do some ugly hacks here to handle unicode since the SimpleParse library
    doesn't have an easy way to deal with it.

    """
    def replacer(match):
        """Replace the safe unicode thingumy"""
        text = match.groups()[1]
        code = int(text)
        return chr(code)

    proper_text = re.sub('(xX)(\d+)(Xx)', replacer, text)

    return proper_text
# << Utility functions >> (5 of 10)
def parseVB(vbtext, container=None, starttoken="line", verbose=0, returnpartial=None):
    """Parse some VB"""

    if returnpartial is None:
        returnpartial = Config["General", "ReportPartialConversion"] == "Yes"

    nodes = buildParseTree(vbtext, starttoken, verbose, returnpartial)

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
# << Utility functions >> (6 of 10)
def getAST(vbtext, starttoken="line", returnpartial=None):
    """Parse some VB to produce an AST"""

    if returnpartial is None:
        returnpartial = Config["General", "ReportPartialConversion"] == "Yes"

    nodes = buildParseTree(vbtext, starttoken, 0, returnpartial, returnast=1)

    return nodes
# << Utility functions >> (7 of 10)
def renderCodeStructure(structure):
    """Render a code structure as Python

    We have this as a separate function so that we can apply the plugins

    """
    return applyPlugins("postProcessPythonText", structure.renderAsCode())
# << Utility functions >> (8 of 10)
def convertVBtoPython(vbtext, *args, **kw):
    """Convert some VB text to Python"""
    m = parseVB(vbtext, *args, **kw)
    return applyPlugins("postProcessPythonText", m.renderAsCode())
# << Utility functions >> (9 of 10)
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
# << Utility functions >> (10 of 10)
def parseVBFile(filename, text=None, parent=None, **kw):
    """Parse some VB from a file"""
    if not text:
        # << Get text >>
        f = open(filename, "r")
        try:
            text = f.read()
        finally:
            f.close()
        # -- end -- << Get text >>
    # << Choose appropriate container >>
    # Type of container to use for each extension type
    container_lookup = {
            ".bas" : VBCodeModule,
            ".cls" : VBClassModule,
            ".frm" : VBFormModule,
    }

    extension = os.path.splitext(filename)[1]
    try:
        container = container_lookup[extension.lower()]
    except KeyError:
        log.warn("File extension '%s' not recognized, using default container", extension)
        container = VBCodeModule
    # -- end -- << Choose appropriate container >>
    new_container=container()
    if parent:
        new_container.parent = parent
    code_structure = parseVB(text, container=new_container, **kw)
    return code_structure
# -- end -- << Utility functions >>

# The following imports must go at the end to avoid import errors
# caused by poor structuring of the package. This needs to be refactored!

# Plug-ins
import extensions
plugins = extensions.loadAllPlugins()

from parserclasses import *

if __name__ == "__main__":	
    from testparse import txt
    m = parseVB(txt)
