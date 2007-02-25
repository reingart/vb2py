import vb2py.vbparser
from vb2py.config import VB2PYConfig
Config = VB2PYConfig()

from vb2py import logger   # For logging output and debugging 
log = logger.getLogger("PythonCardControls")

twips_per_pixel = 15

# << Events >>
# << EventSupport >>
class ControlEvent:
    """Represents a control event mapping from VB to PythonCard

    A control event (eg MouseClick) is defined in VB with a certain name
    and list of parameters. PythonCard has an analogus event with a
    different name and all the parameters are bound up in an event
    object.

    This class helps in the mapping of one to the other.

    """

    # << ControlEvent methods >> (1 of 2)
    def __init__(self, vbname, pyname, vbargs=None, pyargs=None):
        """Initialize the control event"""
        self.vbname = vbname
        self.pyname = pyname
        self.vbargs = vbargs or []
        self.pyargs = pyargs or []
    # << ControlEvent methods >> (2 of 2)
    def updateMethodDefinition(self, method, name):
        """Update the definition of a method based on this translation"""
        method.identifier = self.pyname % name
        # Overwrite the parameter definition
        method.parameters = [vb2py.vbparser.VBRenderDirect("*args")]
        method.scope = "Public"
        #
        # Map arguments
        if self.vbargs:
            mapping = "%s = vbGetEventArgs([%s], args)" % (
                        ", ".join(self.vbargs),
                        ", ".join(['"%s"' % arg for arg in self.pyargs]))
            method.block.blocks.insert(0, vb2py.vbparser.VBRenderDirect(mapping, indent=1, crlf=1))
    # -- end -- << ControlEvent methods >>
# -- end -- << EventSupport >>

#
evtClick = ControlEvent("%s_Click", "on_%s_mouseClick")
evtDblClick = ControlEvent("%s_DblClick", "on_%s_mouseDoubleClick")
evtClickAll = (evtClick, evtDblClick)

evtRefresh = ControlEvent("%s_Refresh", "on_%s_Refresh_NOTSUPPORTED") # TODO: what is the Pythoncard equivalent

evtChange = ControlEvent("%s_Change", "on_%s_textUpdate")

#
# Focus
evtGotFocus = ControlEvent("%s_GotFocus", "on_%s_gainFocus")
evtLostFocus = ControlEvent("%s_LostFocus", "on_%s_loseFocus")
evtSetFocus = ControlEvent("%s_SetFocus", "on_%s_setFocus_NOTSUPPORTED")  # TODO: what is the Pythoncard equivalent
evtFocusAll = (evtGotFocus, evtLostFocus, evtSetFocus)

#
# Mouse moving
evtMouseMove = ControlEvent("%s_MouseMove", "on_%s_mouseMove", 
                            ("Button", "Shift", "X", "Y"),
                            ("ButtonDown()", "ShiftDown()", "x", "y"))
evtMouseDown = ControlEvent("%s_MouseDown", "on_%s_mouseDown",
                            ("Button", "Shift", "X", "Y"),
                            ("ButtonDown()", "ShiftDown()", "x", "y"))
evtMouseUp = ControlEvent("%s_MouseUp", "on_%s_mouseUp",
                            ("Button", "Shift", "X", "Y"),
                            ("ButtonDown()", "ShiftDown()", "x", "y"))
evtMouseAll = (evtMouseMove, evtMouseDown, evtMouseUp)

#
# Pathologically some events have lower case X and Y in the VB version!
evtMouseMoveLC = ControlEvent("%s_MouseMove", "on_%s_mouseMove", 
                            ("Button", "Shift", "x", "y"),
                            ("ButtonDown()", "ShiftDown()", "x", "y"))
evtMouseDownLC = ControlEvent("%s_MouseDown", "on_%s_mouseDown",
                            ("Button", "Shift", "x", "y"),
                            ("ButtonDown()", "ShiftDown()", "x", "y"))
evtMouseUpLC = ControlEvent("%s_MouseUp", "on_%s_mouseUp",
                            ("Button", "Shift", "x", "y"),
                            ("ButtonDown()", "ShiftDown()", "x", "y"))
evtMouseAllLC = (evtMouseMoveLC, evtMouseDownLC, evtMouseUpLC)


#
# Keys
evtKeyUp = ControlEvent("%s_KeyUp", "on_%s_keyUp_NOTSUPPORTED") # TODO: what is the Pythoncard equivalent
evtKeyDown = ControlEvent("%s_KeyDown", "on_%s_keyDown_NOTSUPPORTED") # TODO: what is the Pythoncard equivalent
evtKeyPress = ControlEvent("%s_KeyPress", "on_%s_keyPress_NOTSUPPORTED") # TODO: what is the Pythoncard equivalent
evtKeyAll = (evtKeyUp, evtKeyDown, evtKeyPress)

#
# Dragging
evtDrag = ControlEvent("%s_Drag", "on_%s_Drag_NOTSUPPORTED") # TODO: what is the Pythoncard equivalent
evtDragDrop = ControlEvent("%s_DragDrop", "on_%s_DragDrop_NOTSUPPORTED") # TODO: what is the Pythoncard equivalent
evtDragOver = ControlEvent("%s_DragOver", "on_%s_DragOver_NOTSUPPORTED") # TODO: what is the Pythoncard equivalent
evtDragAll = (evtDrag, evtDragDrop, evtDragOver)
# -- end -- << Events >>
# << SupportedControls >> (1 of 2)
class VBControl:
    """Base class for all VB controls"""

    pycard_name = "VBControl"
    is_container = 0 # 1 for container classes like frames

    _attribute_translations = { 
                "Visible" : "visible",
                "BackColor" : "backgroundColor",
                "ForeColor" : "foregroundColor",
                "ToolTipText" : "toolTip",
                }

    #
    # Lookup table showing the VB event name and the Pythoncard event name
    _events = ()

    # Some standard attributes which can be absent
    Caption = "UnknownCaption"

    # << class VBControl methods >> (1 of 14)
    def _getPropertyList(cls):
        """Return a list of the properties of this control"""
        items = []
        for item in dir(cls):
            if not (item.startswith("_") or item.startswith("vbobj_")):
                items.append(item)
        return items

    _getPropertyList = classmethod(_getPropertyList)
    # << class VBControl methods >> (2 of 14)
    def _getControlList(cls):
        """Return a list of the controls contained in this control"""
        items = []
        for item in dir(cls):
            if not item.startswith("_") and item.startswith("vbobj_"):
                items.append(item)
        return items

    _getControlList = classmethod(_getControlList)
    # << class VBControl methods >> (3 of 14)
    def _getControlsOfType(cls, type_name=None):
        """Return a control with a given type"""
        lst = []
        for item in dir(cls):
            if not item.startswith("_") and item.startswith("vbobj_"):
                obj = cls._get(item)
                if obj.pycard_name == type_name or type_name is None:
                    lst.append(obj)
                if obj.is_container:
                    lst.extend(obj._getControlsOfType(type_name))
        return lst

    _getControlsOfType = classmethod(_getControlsOfType)
    # << class VBControl methods >> (4 of 14)
    def _getContainerControls(cls):
        """Return all container controls"""
        lst = []
        for item in dir(cls):
            if not item.startswith("_"):
                obj = cls._get(item)
                try:
                    is_container = obj.is_container
                except AttributeError:
                    pass
                else:
                    if is_container:
                        lst.append(obj)
        return lst

    _getContainerControls = classmethod(_getContainerControls)
    # << class VBControl methods >> (5 of 14)
    def _get(cls, name, default=None):
        """Get one of our items"""
        try:
            return getattr(cls, name)
        except AttributeError:
            if default is not None:
                return default
            else:
                raise

    _get = classmethod(_get)
    # << class VBControl methods >> (6 of 14)
    def _realName(cls):
        """Return our real name"""
        return cls.__name__[6:]

    _realName = classmethod(_realName)
    # << class VBControl methods >> (7 of 14)
    def _getControlEntry(cls):
        """Return the pycard representation of this object"""
        #
        # Get dictionary entries for this object
        d = {}
        ret = [d]
        d['name'] = cls._realName()
        try:
            d['position'] = (cls.Left/twips_per_pixel, cls.Top/twips_per_pixel)
        except AttributeError:
            pass
        #
        # Convert VB attributes to PythonCard attributes
        for attr, pycard_attr in cls._attribute_translations.iteritems():
            if hasattr(cls, attr):
                value = getattr(cls, attr)
                # Check for colours - these are bad news!
                if attr.endswith("Color"):
                    value = cls._getPyCardColours(value)
                d[pycard_attr] = value

        d.update(cls._getClassSpecificControlEntries())
        # Set the type - we do this here because occasionally the type will change after _getClassSpecificControlEntries
        if Config["General", "UseVBPythonCardControls"].find(cls.pycard_name) > -1:
            d['type'] = "VB%s" % cls.pycard_name
        else:
            d['type'] = cls.pycard_name
        #
        # Watch out for container objects - we have to recur down them
        if cls.is_container:
            cls._processChildObjects()
            for cmp in cls._getControlList():
                obj = cls._get(cmp)
                entry = obj._getControlEntry()
                if entry:
                    ret += entry
        return ret

    _getControlEntry = classmethod(_getControlEntry)
    # << class VBControl methods >> (8 of 14)
    def _getClassSpecificControlEntries(cls):
        """Return additional items for this entry

        This method is normally overriden in the subclass

        """
        return {}

    _getClassSpecificControlEntries = classmethod(_getClassSpecificControlEntries)
    # << class VBControl methods >> (9 of 14)
    def _mapNameReference(cls, match):
        """Map a reference to this object in code to something meaningful

        We have two issues, scope and attributes. The scope we need to map to self.components
        That was easy. But we also need to map attributes. We can't do that in the base class
        but a subclass can hopefully help us out via the _attributeTranslationClass

        """
        if match.groups()[0] is not None:
            return "self.components.%s.%s" % (cls._realName(), 
                                              cls._attributeTranslation(match.groups()[0]))
        else:
            return "self.components.%s" % (cls._realName(),)


    _mapNameReference = classmethod(_mapNameReference)
    # << class VBControl methods >> (10 of 14)
    def _attributeTranslation(cls, name):
        """Convert a VB attribute to a Python one"""
        try:
            return cls._attribute_translations[name]
        except KeyError:
            return VBControl._attribute_translations.get(name, name)

    _attributeTranslation = classmethod(_attributeTranslation)
    # << class VBControl methods >> (11 of 14)
    def _processChildObjects(cls):
        """Before we deal with our child object we get a chance to do some processing

        Sub-classed can use this to do special things, like frames adjusting the
        properties of our children

        """
        for container in cls._getContainerControls():
            container._processChildObject()

    _processChildObjects = classmethod(_processChildObjects)
    # << class VBControl methods >> (12 of 14)
    def _getEvents(cls):
        """Return a list of the events that this control has

        This is a list of tuples of the form
          (VBEventName, PythonCardEventName)


        """
        return cls._events

    _getEvents = classmethod(_getEvents)
    # << class VBControl methods >> (13 of 14)
    def _getAttribute(cls, name, default=None):
        """Return a property of this class with a default if it isn't there

        VB doesn't store attributes unless they differ from the base class value
        and so you have to be careful when getting attributes because they may
        not be there.

        """
        try:
            return getattr(cls, name)
        except AttributeError:
            return default

    _getAttribute = classmethod(_getAttribute)
    # << class VBControl methods >> (14 of 14)
    def _getPyCardColours(cls, vbcolour): 
            """Convert a VB colour to a PythonCard colour

            There are a number of issues here. The main one is that VB often
            uses the System colours which are not valid colour references.

            """
            log.debug("Converting colour '%s'" % str(vbcolour))
            if isinstance(vbcolour, int):
                if vbcolour < 0:
                    log.debug("Looks like a system colour - assume grey")
                    return "Grey"
                else:
                    log.debug("Looks like a normal colour")
                    return vbcolour
            else:
                return vbcolour

    _getPyCardColours = classmethod(_getPyCardColours)
    # -- end -- << class VBControl methods >>
# << SupportedControls >> (2 of 2)
# << Controls >> (1 of 16)
class Menu(VBControl):
    """Menu"""
    pycard_name = "Menu"

    def _getControlEntry(cls):
        """Return the pycard representation of this object"""
        return {}

    _getControlEntry = classmethod(_getControlEntry)

    def _pyCardMenuEntry(cls):		
        """Return the entry for this menu"""
        return {
                    "type" : "Menu",
                    "name" : cls._realName(),
                    "label" : cls.Caption,
             }

    _pyCardMenuEntry = classmethod(_pyCardMenuEntry)
# << Controls >> (2 of 16)
class CommandButton(VBControl):
    pycard_name = "Button"


    #
    # Lookup table showing the VB event name and the Pythoncard event name
    _events = evtClickAll + evtFocusAll + evtDragAll + evtMouseAll + evtKeyAll


    def _getClassSpecificControlEntries(cls):
        """Return additional items for this entry"""
        return {  "label" : cls.Caption,
                  "size" : (cls.Width/twips_per_pixel, cls.Height/twips_per_pixel),
}

    _getClassSpecificControlEntries = classmethod(_getClassSpecificControlEntries)
# << Controls >> (3 of 16)
class ComboBox(VBControl):
    pycard_name = "ComboBox"

    #
    # Lookup table showing the VB event name and the Pythoncard event name
    _events = (evtChange,) + evtClickAll + evtFocusAll + evtDragAll + evtMouseAll + evtKeyAll

    def _getEntriesFromFRX(cls, data):
        """Get list entries from FRX file"""
        file, offset = data.split("@")
        raw = open(file, "r").read()
        ptr = int(offset, 16)
        num = ord(raw[ptr])
        ptr = ptr + 4
        lst = []
        for i in range(num):
            length = ord(raw[ptr])
            lst.append(raw[ptr+2:ptr+2+length])
            ptr += 2+length
        return lst

    _getEntriesFromFRX = classmethod(_getEntriesFromFRX)			


    def _getClassSpecificControlEntries(cls):
        """Return additional items for this entry"""
        d = {"size" : (cls.Width/twips_per_pixel, cls.Height/twips_per_pixel)}
        if hasattr(cls, "List"):
            d["items"] =  cls._getEntriesFromFRX(cls.List)
        else:
            d["items"] = []
        return d

    _getClassSpecificControlEntries = classmethod(_getClassSpecificControlEntries)
# << Controls >> (4 of 16)
class ListBox(ComboBox):
    pycard_name = "List"


    #
    # Lookup table showing the VB event name and the Pythoncard event name
    _events = evtClickAll + evtFocusAll + evtDragAll + evtMouseAll + evtKeyAll
# << Controls >> (5 of 16)
class Label(VBControl):
    pycard_name = "StaticText"

    # Default properties
    Caption = "Label"

    def _getClassSpecificControlEntries(cls):
        """Return additional items for this entry"""
        return {  "text" : cls.Caption,
                }

    _getClassSpecificControlEntries = classmethod(_getClassSpecificControlEntries)
# << Controls >> (6 of 16)
class Image(VBControl):
    pycard_name = "BitmapCanvas"

    Stretch = 0

    #
    # Lookup table showing the VB event name and the Pythoncard event name
    _events = evtClickAll + evtFocusAll + evtDragAll + evtMouseAll + evtKeyAll

    def _getClassSpecificControlEntries(cls):
        """Return additional items for this entry"""
        d = {"size" : (cls.Width/twips_per_pixel, cls.Height/twips_per_pixel),
             "Stretch" : cls.Stretch}
        return d

    _getClassSpecificControlEntries = classmethod(_getClassSpecificControlEntries)
# << Controls >> (7 of 16)
class CheckBox(VBControl):
    pycard_name = "CheckBox"

    #
    # Lookup table showing the VB event name and the Pythoncard event name
    _events = evtClickAll + evtFocusAll + evtDragAll + evtMouseAll + evtKeyAll


    _attribute_translations = { 
                    "Value" : "checked",
                    }
    _attribute_translations.update(VBControl._attribute_translations)

    def _getClassSpecificControlEntries(cls):
        """Return additional items for this entry"""
        return {  "label" : cls.Caption,
                  "checked" : cls._get("Value", 0),
                }

    _getClassSpecificControlEntries = classmethod(_getClassSpecificControlEntries)
# << Controls >> (8 of 16)
class TextBox(VBControl):
    pycard_name = "TextField"

    # Default properties
    Text = ""

    #
    # Lookup table showing the VB event name and the Pythoncard event name
    _events = (evtClick, evtChange) + evtFocusAll + evtDragAll + evtMouseAll + evtKeyAll

    _attribute_translations = { 
                    "Text" : "text",
                    }
    _attribute_translations.update(VBControl._attribute_translations)

    def _getClassSpecificControlEntries(cls):
        """Return additional items for this entry"""
        # Check for multiline - if so then we need a TextArea instead
        if cls._getAttribute("MultiLine", 0):
            log.info("Changed TextField to TextArea for '%s'" % cls._realName())
            cls.pycard_name = "TextArea"
        # Check if our text is stored in the FRX file
        if cls.Text.lower().find(".frx@") > -1:
            log.info("Looking for text data in the FRX file")
            cls.Text = cls._getEntriesFromFRX(cls.Text)
        return {  "text" : cls.Text,
                  "size" : (cls.Width/twips_per_pixel, cls.Height/twips_per_pixel),
}

    _getClassSpecificControlEntries = classmethod(_getClassSpecificControlEntries)


    def _getEntriesFromFRX(cls, data):
        """Get text entries from FRX file"""
        file, offset = data.split("@")
        raw = open(file, "r").read()
        ptr = max(int(offset, 16)-1, 0)
        num = ord(raw[ptr])
        return raw[ptr+1:ptr+1+num-1]

    _getEntriesFromFRX = classmethod(_getEntriesFromFRX)
# << Controls >> (9 of 16)
class Form(VBControl):
    """Form"""

    HeightModifier = 20 # Used to account for form borders
    MenuHeight = 20 # Allows for menu
    Caption = "Form"
# << Controls >> (10 of 16)
class Font(VBControl):
    """A Font"""
# << Controls >> (11 of 16)
class Timer(VBControl):
    """A timer"""

    pycard_name = "Timer"
# << Controls >> (12 of 16)
class Frame(VBControl):
    """Frame"""
    pycard_name = "StaticBox"
    is_container = 1

    def _getClassSpecificControlEntries(cls):
        """Return additional items for this entry"""
        return {  "size" : (cls.Width/twips_per_pixel, cls.Height/twips_per_pixel),
                }

    _getClassSpecificControlEntries = classmethod(_getClassSpecificControlEntries)


    def _processChildObjects(cls):
        """Adjust the left and top properties of our children"""
        for item in cls._getControlList():
            obj = cls._get(item)
            log.debug("Offsetting %s, %s" % (obj, cls))
            if hasattr(obj, "Left"):
                obj.Left += cls.Left
            if hasattr(obj, "Top"):
                obj.Top += cls.Top
        #for container in cls._getContainerControls():
        #	container._processChildObjects()

    _processChildObjects = classmethod(_processChildObjects)
# << Controls >> (13 of 16)
class OptionButton(VBControl):
    pycard_name = "RadioGroup"

    def _getClassSpecificControlEntries(cls):
        """Return additional items for this entry"""
        return {  "items" : getattr(cls, "items", ""),
                  "selected" : getattr(cls, "selected", ""),
                }

    _getClassSpecificControlEntries = classmethod(_getClassSpecificControlEntries)
# << Controls >> (14 of 16)
class TreeView(VBControl):
    pycard_name = "TreeView"

    #
    # Lookup table showing the VB event name and the Pythoncard event name
    _events = (evtClick, evtChange) + evtFocusAll + evtDragAll + evtMouseAllLC + evtKeyAll

    def _getClassSpecificControlEntries(cls):
        """Return additional items for this entry"""
        d = { 
            "size" : (cls.Width/twips_per_pixel, cls.Height/twips_per_pixel),
        }
        return d            

    _getClassSpecificControlEntries = classmethod(_getClassSpecificControlEntries)
# << Controls >> (15 of 16)
class ImageList(VBControl):
    pycard_name = "ImageList"

    def _getClassSpecificControlEntries(cls):
        """Return additional items for this entry"""
        return { }

    _getClassSpecificControlEntries = classmethod(_getClassSpecificControlEntries)
# << Controls >> (16 of 16)
class VBUnknownControl(Label):
    """A control representing an unknown control

    We fake it out as a label

    """

    Caption = "Unknown control"
# -- end -- << Controls >>

possible_controls = { 
    "VBControl" : "VBControl",
    "Form" : "Form",
    "CommandButton" : "CommandButton",
    "OptionButton" : "OptionButton",
    "TextBox" : "TextBox",
    "Label" : "Label",
    "Menu" : "Menu",
    "ComboBox" : "ComboBox",
    "ListBox" : "ListBox",
    "CheckBox" : "CheckBox",
    "Frame" : "Frame",
    "Font" : "Font",
    "Image" : "Image",
    "ImageList" : "ImageList",
    "TreeView" : "TreeView",
    "Timer" : "Timer",

    "VBUnknownControl" : "VBUnknownControl",
    "FileListBox" : "VBUnknownControl",
    "OLE" : "VBUnknownControl",
    "Shape" : "VBUnknownControl",
}
# -- end -- << SupportedControls >>
