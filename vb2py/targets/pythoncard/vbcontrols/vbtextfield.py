from vb2py.targets.pythoncard.controlclasses import VBWrapped, VBWidget
import vb2py.logger
log = vb2py.logger.getLogger("VBTextField")

from PythonCardPrototype.components import textfield
from wxPython import wx
import sys
from PythonCardPrototype import binding, event, registry, widget


class VBTextField(VBWidget): 
    __metaclass__ = VBWrapped 

    _translations = { 
            "Text" : "text", 
            "Enabled" : "enabled", 
            "Visible" : "visible", 
        } 

    _indexed_translations = { 
            "Left" : ("position", 0), 
            "Top" : ("position", 1), 
            "Width" : ("size", 0), 
            "Height" : ("size", 1), 
        } 

    _proxy_for = textfield.TextField


log.debug("Registering VBTextField as '%s'" % sys.modules[__name__].VBTextField)
registry.getRegistry().register( sys.modules[__name__].VBTextField )
