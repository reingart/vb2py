# Created by Leo from: C:\Development\Python23\Lib\site-packages\vb2py\vb2py.leo

# << Documentation >>
"""VB2Py - VB to Python + PythonCard conversion

This application converts VB projects to Python and PythonCard projects.
The form layouts are converted and you can optionally convert the VB
to Python, including translation of the VB events.

The VB conversion is very preliminary!

So is the layout...

"""
# -- end -- << Documentation >>
# << Declarations >>
import re	    # For text processing
import os     # For file processing
import pprint # For outputting dictionaries
import sys    # For getting Exec prefix
import getopt # For command line arguments
import imp    # For dynamic import of classes

from utils import rootPath
from config import VB2PYConfig
Config = VB2PYConfig()

import logger   # For logging output and debugging 
log = logger.getLogger("vb2Py")

import vbparser

#from controls import *
twips_per_pixel = 15

__app_name__ = "VB2Py"
__version__ = "0.2.2"

# Try to import ctypes module to read type libraries
try:
    import ctypes.com.tools.readtlb as readtlb
except ImportError:
    readtlb = None
# -- end -- << Declarations >>
# << Error classes >>
class VB2PyError(Exception): """An error occured converting a project"""
# -- end -- << Error classes >>

# << VBConverter >> (1 of 15)
class VBConverter(object):
    """Class to convert VB projects to Python Card projects"""

    # << class VBConverter methods >> (1 of 4)
    def __init__(self, resource, parser=None):
        """Initialize with a target resource"""
        self._target_resource = resource
        if parser:
            self.parser = parser
        else:
            self.parser = ProjectParser
    # << class VBConverter methods >> (2 of 4)
    def doConversion(self, filename, callback=None):
        """Convert the named VB project to a python project"""
        project_root, project_file = os.path.split(filename)
        self.logText("Parsing '%s'" % filename)
        project = self.parser(filename)
        project.doParse()
        self.logText("Processing project '%s'" % project.name)
        #
        self.resources = []
        #
        # TODO: Refactor here
        self.project_structure = vbparser.VBProject()
        #
        total = len(project.forms) + len(project.modules) + len(project.classes) + 1
        done = 0
        #
        # << Handle references >>
        if project.references:
            if not readtlb:
                self.logText("Unable to import ctypes - external references will not be converted")
            else:
                refs = self.handleReferences(project.references, self.project_structure, project_root)
                self.resources.append(ExternalRefParser(
                        os.path.join(project_root, "COM_Externals.py"), refs))
        # -- end -- << Handle references >>
        # << Handle modules >>
        self.modules = []
        for module in project.modules:
            module_name, module_filename = module.split(";")
            done +=1
            if callback:
                callback("Reading module '%s'" % module_name, 100.0*done/total)
            #
            self.logText("Reading module '%s'" % module_name)
            mod = ModuleParser(os.path.join(project_root, module_filename.strip()), module_name)
            mod.doParse(self.project_structure)
            self.modules.append(mod)
            self.resources.append(mod)
        # -- end -- << Handle modules >>
        # << Handle forms >>
        self.forms = []
        for form in project.forms:
            done +=1
            if callback:
                callback("Reading form '%s'" % form, 100.0*done/total)
            #
            self.logText("Reading form '%s'" % form)
            frm = FormParser(os.path.join(project_root, form))
            frm.doParse(self.project_structure)
            if frm.form:
                frm.resources = self._target_resource()
                frm.resources.updateFrom(frm.form)
                frm.resources.updateCode(frm)
                frm.resources.code_block = frm.code_block
                frm.resources.log = log
                self.forms.append(frm)
                self.resources.append(frm.resources)
            #
        # -- end -- << Handle forms >>    
        # << Handle classes >>
        self.classes = []
        for cls in project.classes:
            cls_name, cls_filename = cls.split(";")
            done +=1
            if callback:
                callback("Reading class '%s'" % cls_name, 100.0*done/total)
            #
            self.logText("Reading class module '%s'" % cls_name)
            class_mod = ClassParser(os.path.join(project_root, cls_filename.strip()), cls_name)
            class_mod.doParse(self.project_structure)
            self.classes.append(class_mod)
            self.resources.append(class_mod)
        # -- end -- << Handle classes >> 
        #
        if callback:
            callback("Done!", 100.0)
    # << class VBConverter methods >> (3 of 4)
    def logText(self, text):
        """Log text to show progress"""
        log.info("> %s" % text)
    # << class VBConverter methods >> (4 of 4)
    def handleReferences(self, references, project, root):
        """Handle the external references

        The references tell us the DLL names. We can then load these using
        ctypes readtlb. This then tells us the objects which are exposed.
        We can then try to use this to insert global references in the
        project which will be replaced by the appropriate calls whenever
        we need them.

        """
        global_names = {}
        # << Add to project >>
        #
        # Now stick these in a fake module
        externals = vbparser.VBCOMExternalModule(modulename="COM_Externals")
        externals.names = global_names
        # -- end -- << Add to project >>               
        for reference in references:
            # << Resolve references >>
            #
            # Get all the external objects that we can reference
            try:
                guid, ver, id, path, name = reference.split("#")
            except ValueError:
                self.logText("Unable to extract reference from '%s'" % reference)
            try:
                tlb = readtlb.TypeLibReader(os.path.join(root, path))
                self.logText("Found '%s' in DLL '%s'" % (tlb.name, path))
                for cls in tlb.coclasses.values():
                    self.logText(" - Member class %s" % cls.name)
                    global_names.setdefault(tlb.name, []).append(cls.name)
                    project.global_objects["%s.%s" % (tlb.name, cls.name)] = externals
            except Exception, err:
                self.logText("Failed while reading library '%s': %s" % (path, err))
            # -- end -- << Resolve references >>

        return externals
    # -- end -- << class VBConverter methods >>
# << VBConverter >> (2 of 15)
class BaseParser(object):
    """A base parser object"""

    # << class BaseParser methods >> (1 of 9)
    def __init__(self, filename, name=None):
        """Initialize the parser"""
        self.filename = filename
        self.text = open(filename.strip(), "r").read() # Use strip to remove \r 
        self.name = name or os.path.splitext(os.path.split(filename)[1])[0]
        self.basedir = os.path.split(filename)[0]
    # << class BaseParser methods >> (2 of 9)
    def doValidation(self):
        """Validate the data we parsed out of the file"""
        pass
    # << class BaseParser methods >> (3 of 9)
    def findMany(self, id):
        """Find a list of values in the file"""
        return self._getPattern(id).findall(self.text)
    # << class BaseParser methods >> (4 of 9)
    def findOne(self, id, default=None):
        """Find a value in the file"""
        match = self._getPattern(id).search(self.text)
        if match:
            return match.groups(1)[0]
        else:
            return default
    # << class BaseParser methods >> (5 of 9)
    def splitSectionByMarker(self, marker):
        """Split a block of text about a marker"""
        pattern = re.compile('^%s ' % marker, re.MULTILINE)
        match = pattern.search(self.text)
        if match:
            return (self.text[:match.start(0)], self.text[match.start(0):])
        else:
            return None
    # << class BaseParser methods >> (6 of 9)
    def _getPattern(self, id):
        """Create a search pattern"""
        return re.compile('^%s\s*=\s*"*(.*?)"*$' % id, re.MULTILINE)
    # << class BaseParser methods >> (7 of 9)
    def parseCode(self, project):
        """Parse the form code"""
        container = self.getContainer()
        #container.parent = project
        container.assignParent(project)
        try:
            self.code_structure = vbparser.parseVB(self.code_block, container=container)		
        except vbparser.VBParserError, err:
            log.error("Unable to parse '%s'(%s): %s" % (self.name, self.filename, err))
            self.code_structure = vbparser.VBMessage(
                        messagetype="ParsingError",
                        message="Failed to parse (%s)" % err)
    # << class BaseParser methods >> (8 of 9)
    def getContainer(self):
        """Return the container to use for code conversion"""
        return vbparser.VBModule()
    # << class BaseParser methods >> (9 of 9)
    def writeToFile(self, basedir, write_code=0):
        """Write this out to a file"""
        raise VB2PyError("Unable to write '%s' to a file" % self)
    # -- end -- << class BaseParser methods >>
# << VBConverter >> (3 of 15)
class ProjectParser(BaseParser):
    """A VB project parser object"""

    # << class ProjectParser methods >> (1 of 2)
    def doParse(self):
        """Parse the text"""
        self.forms = self.findMany("Form")
        self.startup = self.findOne("Startup")
        self.name = self.findOne("Name")
        self.modules = self.findMany("Module")
        self.classes = self.findMany("Class")
        self.references = self.findMany("Reference")
        #
        # Do sanity check
        self.doValidation()
    # << class ProjectParser methods >> (2 of 2)
    def doValidation(self):
        """Validate that the project was reasonable"""
        #if not self.forms:
        #	raise VB2PyError("No forms in the project! Nothing to convert")
    # -- end -- << class ProjectParser methods >>
# << VBConverter >> (4 of 15)
class FileParser(ProjectParser):
    """A parser for VB files which are not part of a project"""

    # << class FileParser methods >> (1 of 2)
    def doParse(self):
        """Parse the text"""
        #
        log.info("Using single file parser")
        #
        self.forms = []
        self.modules = []
        self.classes = []
        self.startup = None
        #
        extn = os.path.splitext(self.filename)[1].lower()
        #
        if extn == ".frm":
            self.name = self.findOne("Attribute VB_Name")
            self.forms = [self.filename]
            self.startup = self.findOne("Startup")
        elif extn == ".bas":
            self.name = self.findOne("Attribute VB_Name")
            self.modules = ["%s; %s" % (self.name, self.filename)]
        elif extn == ".cls":
            self.name = self.findOne("Attribute VB_Name")
            self.classes = ["%s; %s" % (self.name, self.filename)]
        else:
            raise VB2PyError("Unknown file extension: '%s'" % extn)
        #
        # Do sanity check
        self.doValidation()
    # << class FileParser methods >> (2 of 2)
    def doValidation(self):
        """Validate that the project was reasonable"""
    # -- end -- << class FileParser methods >>
# << VBConverter >> (5 of 15)
class FormParser(BaseParser):
    """A VB form parser object"""

    # << class FormParser methods >> (1 of 4)
    def doParse(self, project):
        """Parse the text"""
        # << Split off code section >>
        #	        We try to find the code section - this is delimeted by a series of 
#	        "Attribute" definitions. This is pretty hokey and there should be a
#	        better way to do this, but this method seems to work so far
        
        data = self.splitSectionByMarker("Attribute")
        if data:
            self.form_data, self.code_block = data
        else:
            self.form_data = self.code_block = None
        # -- end -- << Split off code section >>
        self.parseForm()
        if self.form:
            self.parseCode(project)
            # << Add controls to form namespace >>
            #	            All the controls that this form owns are in the local namespace
#	            so we need to add them to the VBFormModule so that they will be
#	            converted to the proper (self.) style
            
            distinct_names = {}
            for control in self.form._getControlsOfType():
                #
                # Add name to namespace
                name = control._realName()
                distinct_names[name] = 1
                #
                # Look for events for this control
                for event in control._getEvents():	
                    event_name, new_name = event.vbname, event.pyname
                    #
                    # Look for local definitions of methods which match the VB events for this object
                    for item in self.code_structure.locals:
                        if event_name % name == item.identifier:
                            # Add a name substitution to translate references to this name to the PythonCard version
                            self.code_structure.name_substitution[event_name % name] = "self." + new_name % name
                            # Change the definition
                            event.updateMethodDefinition(item, name)


            self.code_structure.local_names.extend(distinct_names.keys())

            #import pdb; pdb.set_trace()

            # Probably need to get self.form._getControlList()
            # then strip front of name (vbobj_txtName) and add to
            # code_structure.local_names
            # -- end -- << Add controls to form namespace >>
    # << class FormParser methods >> (2 of 4)
    #	    To parse the form we take advantage of the fact that the data is almost
#	    a readable python structure anyway. It looks like
#	    VERSION 5.00
#	    Begin VB.Form frmMain 
#	       Caption         =   "Form1"
#	       StartUpPosition =   3  'Windows Default
#	       Begin VB.CommandButton btnSecond 
#	          Caption         =   "Second Form"
#	          Height          =   375
#	       End
#	    End
#	    
#	    Begin = class
#	    name = super
    
    def parseForm(self):
        """Parse the form definition"""
        self.form_data = self.form_data.replace("\r\n", "\n") # For *nix
        # << Get name of form class >>
        #	        Need to get the name of the form class - this is not always the 
#	        same as the filename
        
        pattern = re.compile(r"^Begin\s+VB\.Form\s+(\w+)", re.MULTILINE)
        name_match = pattern.findall(self.form_data)

        if name_match:
            self.name = name_match[0]
        # -- end -- << Get name of form class >>
        # << Begin class conversion >>
        pattern = re.compile(r'^(\s*)Begin\s+(\w+)\.(.+?)\s+(.+?)\s*?$', re.MULTILINE)

        def sub_begin(match):
            if match.groups()[1] in ("VB", "ComctlLib"):
                return '%sclass vbobj_%s(resource.%s):' % (
                        match.groups()[0],
                        match.groups()[3],
                        resource.possible_controls.get(match.groups()[2], "VBUnknownControl"))
            else:
                log.warn('Unknown control %s.%s' % (match.groups()[1], match.groups()[2]))
                return '%sclass vbobj_%s(resource.VBUnknownControl):' % (
                        match.groups()[0], match.groups()[3])

        self.form_data = pattern.sub(sub_begin, self.form_data)
        # -- end -- << Begin class conversion >>
        # << Convert properties >>
        pattern = re.compile(r'^(\s*)BeginProperty\s+(\w+)(\(.*?\))?\s(.*?)$', re.MULTILINE)

        def sub_beginproperty(match):
            return '%sclass vbobj_%s(resource.%s): # %s %s' % (
                    match.groups()[0],
                    match.groups()[1],
                    resource.possible_controls.get(match.groups()[1], "VBUnknownControl"),
                    match.groups()[2],
                    match.groups()[3],
        )

        self.form_data = pattern.sub(sub_beginproperty, self.form_data)
        # -- end -- << Convert properties >>
        # << Menu shortcuts >>
        pattern = re.compile(r'^(\s*)Shortcut\s*=\s*(\S+)\s*$', re.MULTILINE)

        def sub_shortcut(match):
            return '%sshortcut = "%s"' % (
                    match.groups()[0],
                    match.groups()[1])

        self.form_data = pattern.sub(sub_shortcut, self.form_data)
        # -- end -- << Menu shortcuts >>
        # << Remove meaningless bits >>
        #
        # End
        pattern = re.compile("^\s*End$", re.MULTILINE)
        self.form_data = pattern.sub("", self.form_data)

        #
        # End Property
        pattern = re.compile("^\s*EndProperty$", re.MULTILINE)
        self.form_data = pattern.sub("", self.form_data)

        #
        # Version
        pattern = re.compile("^VERSION\s+.*?$", re.MULTILINE)
        self.form_data = pattern.sub("", self.form_data)

        #
        # Comments
        self.form_data = self.form_data.replace("'", "#")
        # -- end -- << Remove meaningless bits >>
        # << Remove references to frx file >>
        def sub_frx(match):
            s = '"%s.frx@%s"' % (os.path.join(self.basedir, match.groups()[0]), match.groups()[1])
            return s.replace("\\", "/")

        pattern = re.compile('\$?"(.*)\.frx":(\S+)', re.MULTILINE)
        self.form_data = pattern.sub(sub_frx, self.form_data)
        # -- end -- << Remove references to frx file >>
        # << Hex numbers >>
        #
        # Convert hex numbers - which are colours 
        # We will have problems with system colours (&H80 ... ) so we 
        # ultimately need a lookup table here

        pattern = re.compile("\&H([A-F0-9]{8})\&", re.MULTILINE)

        def sub_hex(match):
            txt = match.groups()[0]
            return "(%d, %d, %d)" % (int(txt[2:4], 16),
                                     int(txt[4:6], 16),
                                     int(txt[6:8], 16))

        self.form_data = pattern.sub(sub_hex, self.form_data)
        # -- end -- << Hex numbers >>
        # << Convert object references >>
        pattern = re.compile(r'^(\s*)Object\s*=\s*"(\S+)"\s*;\s*(.*?)$', re.MULTILINE)

        def sub_beginobject(match):
            return '%s# %s, %s' % (
                    match.groups()[0],
                    match.groups()[1],
                    match.groups()[2],
        )

        self.form_data = pattern.sub(sub_beginobject, self.form_data)
        # -- end -- << Convert object references >>
        if Config["General", "DumpFormData"] == "Yes":
            log.debug(self.form_data)
        self.namespace = {"resource" : resource, "Object" : NameSpace()}
        try:
            exec self.form_data.replace("\r", "") in self.namespace
        except Exception, err:
            log.error("Failed during conversion of '%s'" % self.name)
            self.form = None
        else:
            self.form = self.namespace["vbobj_%s" % self.name]
            self.form.name = self.name
            self.groupOptionButtons(self.form)
    # << class FormParser methods >> (3 of 4)
    #	    Option buttons in VB are just properties of their container. The 
#	    equivalent in PythonCard are option groups which are all in one place. 
#	    We have to go into the VB form and pull all the options together so we 
#	    start at the form level and then descend down each sub-container to 
#	    make sure that we grab them all as a group.
    
    def groupOptionButtons(self, cls):
        """Pull all the option buttons together for this class"""
        #
        # Get all options
        options = cls._getControlsOfType("RadioGroup")
        if not options:
            return
        #
        # Get start properties
        grp = options[0]
        #
        # Make a list of the captions and of the currently selected one
        captions = []
        selected = None
        for option in options:
            caption = option._get('Caption', 'Option')
            captions.append(caption)
            if option._get('Value', 0) == -1:
                selected = caption
                #
                # TODO: map names of options to pycard names
            #
            # Delete the group
            if option is not grp and hasattr(cls, option.__name__):
                delattr(cls, option.__name__)
        #
        # Now add a new option group
        grp.items = captions
        grp.selected = selected
        #
        # Make sure we also look in other containers on this form
        for container in cls._getContainerControls():
            self.groupOptionButtons(container)
    # << class FormParser methods >> (4 of 4)
    def getContainer(self):
        """Return the container to use for code conversion"""
        return vbparser.VBFormModule(modulename=self.name)
    # -- end -- << class FormParser methods >>
# << VBConverter >> (6 of 15)
class ModuleParser(BaseParser):
    """A VB module parser object"""

    # << class ModuleParser methods >> (1 of 3)
    def doParse(self, project):
        """Parse the text"""
        self.code_block = self.text
        self.parseCode(project)
    # << class ModuleParser methods >> (2 of 3)
    def getContainer(self):
        """Return the container to use for code conversion"""
        return vbparser.VBCodeModule(modulename=self.name)
    # << class ModuleParser methods >> (3 of 3)
    def writeToFile(self, basedir, write_code=0):
        """Write this out to a file"""
        fname = os.path.join(basedir, self.name) + ".py"
        fle = open(fname, "w")
        log.info("Writing: %s" % fname)
        try:
            fle.write(vbparser.renderCodeStructure(self.code_structure))
        finally:
            fle.close()
    # -- end -- << class ModuleParser methods >>
# << VBConverter >> (7 of 15)
class ClassParser(ModuleParser):
    """A VB class module parser object"""

    # << class ClassParser methods >>
    def getContainer(self):
        """Return the container to use for code conversion"""
        return vbparser.VBClassModule(modulename=self.name, classname=self.name)
    # -- end -- << class ClassParser methods >>
# << VBConverter >> (8 of 15)
class BaseResource(object):
    """A VB form resource object"""

    target_name = "Python"
    name = "baseResource"
    form_class_name = "FormClass"
    form_super_classes = []
    allow_new_style_class = 1

    # << class BaseResource methods >> (1 of 5)
    def __init__(self, basesourcefile=None):
        """Initialize the resource"""
        if basesourcefile is None:
            self.basesourcefile = os.path.join(
                        rootPath(), "targets", self.target_name, "basesource")
        else:
            self.basesourcefile = basesourcefile
        #
        # Apply default resource
        self._rsc = {}
        self._code = ""
        #
        log.debug("BaseResource init")
    # << class BaseResource methods >> (2 of 5)
    def updateFrom(self, form):
        """Update our resource from the form object"""
        # << Main properties >>
        #
        # The main properties of the form
        d = self._rsc['stack']['backgrounds'][0]
        self.name = form.name
        d['name'] = form.name
        d['title'] = form.Caption

        #
        # Add menu height to form height if it is needed
        if form._getControlsOfType("Menu"):
            height_modifier = form.HeightModifier + form.MenuHeight
        else:
            height_modifier = form.HeightModifier

        d['size'] = (form.ClientWidth/twips_per_pixel, form.ClientHeight/twips_per_pixel+height_modifier)
        d['position'] = (form.ClientLeft/twips_per_pixel, form.ClientTop/twips_per_pixel)
        # -- end -- << Main properties >>
        # << Components >>
        #
        # The components (controls) on the form
        c = self._rsc['stack']['backgrounds'][0]['components']

        for cmp in form._getControlList():
            obj = form._get(cmp)
            entry = obj._getControlEntry()
            if entry:
                c += entry
        # -- end -- << Components >>
        # << Menus >>
        #
        # The menus
        m = []
        self._rsc['stack']['backgrounds'][0]['menubar']['menus'] = m

        self.addMenus(form, m)
        # -- end -- << Menus >>
    # << class BaseResource methods >> (3 of 5)
    def updateCode(self, form):
        """Update our code blocks"""
        #
        # Make sure the code structure has the right context
        form.code_structure.classname = self.form_class_name
        form.code_structure.superclasses = self.form_super_classes
        form.code_structure.allow_new_style_class = self.allow_new_style_class
        #
        # Convert it to Python code
        self.code_structure = form.code_structure
    # << class BaseResource methods >> (4 of 5)
    def addMenus(self, obj, to_menu):
        """Add menus"""
        for mnu in obj._getControlsOfType("Menu"):
            d = mnu._pyCardMenuEntry()
            d["items"] = []
            to_menu.append(d)
            self.addMenus(mnu, d['items'])
            if not d['items']:
                del(d['items'])
                d['type'] = 'MenuItem'
    # << class BaseResource methods >> (5 of 5)
    def writeToFile(self, basedir, write_code=0):
        """Write ourselves out to a directory"""
    # -- end -- << class BaseResource methods >>
# << VBConverter >> (9 of 15)
class ExternalRefParser(ModuleParser):
    """Handlers writing out of external references"""

    # << class ExternalRefParser methods >>
    def __init__(self, filename, refs):
        """Initialize the parser"""
        self.filename = filename
        self.name = os.path.splitext(os.path.split(filename)[1])[0]
        self.basedir = os.path.split(filename)[0]
        self.code_structure = refs
    # -- end -- << class ExternalRefParser methods >>
# << VBConverter >> (10 of 15)
class NameSpace:
    """Namespace to store values"""
# << VBConverter >> (11 of 15)
def main():
    """Main application"""
    # << Parse options >>
    try:
        opts, args = getopt.getopt(sys.argv[1:], "dfchvst:", ["help", "code", "version", "supports"])
    except getopt.GetoptError, err:
        # print help information and exit:
        usage(error=err)
        sys.exit(2)

    do_code = 0
    target = "PythonCard"
    parser = ProjectParser

    for o, a in opts:
        if o in ("-h", "--help"):
            usage()
            sys.exit()
        if o in ("-s", "--supports"):
            supports()
            sys.exit()
        if o in ("-c", "--code"):
            do_code = 1
        if o in ("-v" , "--version"):
            print "%s v%s" % (__app_name__, __version__)
            sys.exit(2)
        if o in ("-t", ):
            target = a
        if o in ("-f", ):
            parser = FileParser
        if o in ("-d", ):
            Config.setLocalOveride("General", "DumpFormData", "Yes")

    if len(args) <> 2:
        usage()
        sys.exit(2)

    project_file, destination_dir = args
    # -- end -- << Parse options >>	
    # << Validate arguments >>
    if not os.path.isfile(project_file):
        print "First parameter must be a valid VB file"
        sys.exit(2)
    elif not os.path.isdir(destination_dir):
        print "Second argument must be a valid directory"
        sys.exit(2)
    # -- end -- << Validate arguments >>
    TargetResource = importTarget(target)
    conv = VBConverter(TargetResource, parser)
    conv.doConversion(project_file)
    renderTo(conv, destination_dir, do_code)
# << VBConverter >> (12 of 15)
#	We need to import classes which are appropriate for the given target using 
#	the imp module

def importTarget(target):
    """Import the target resource"""
    global event_translator, resource

    from targets.pythoncard	import resource
    TargetResource = resource.Resource

    try:
        event_translator = resource.event_translator
    except AttributeError:
        event_translator = {}

    return TargetResource
# << VBConverter >> (13 of 15)
def renderTo(conv, destination_dir, do_code=1):
    """Render the converted code to a localtion"""	
    for item in conv.resources:
        item.writeToFile(destination_dir, do_code)
# << VBConverter >> (14 of 15)
def usage(error=None):
    """Print usage statement"""
    if error:
        print "\n\nInvalid option! (%s)" % error
    print "\nconverter -chvs project.vpb destination\n\n" \
          "   project.vbp = VB project file\n" \
          "   desination  = Destination directory for files\n\n" \
          "   -tTarget = Target platform" \
          "   -c = Convert VB code also\n" \
          "   -v = Print version and exit\n" \
          "   -h = Print this message\n" \
          "   -f = Just process the given file\n" \
          "   -d = Dump out the form definition classes\n"
# << VBConverter >> (15 of 15)
def supports():
    """Show a list of controls supported by this converter"""
    print "This command line option is not currently available"
    return
    print "\nSupported controls\n"
    for control in possible_controls:
        ctrl = possible_controls[control]
        if ctrl <> VBUnknownControl:
            print "   - %s (as %s)" % (control, ctrl.pycard_name)
    print
# -- end -- << VBConverter >>

if __name__ == "__main__":
    main()
