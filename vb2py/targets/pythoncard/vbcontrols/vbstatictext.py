from vb2py.targets.pythoncard.controlclasses import VBWrapped, VBWidget
import vb2py.logger
log = vb2py.logger.getLogger("VBStaticText")

from PythonCardPrototype.components import statictext
from wxPython import wx
import sys
from PythonCardPrototype import binding, event, registry, widget


class VBStaticText(VBWidget): 
    __metaclass__ = VBWrapped 

    _translations = { 
            "Caption" : "text", 
            "Enabled" : "enabled", 
            "Visible" : "visible", 
        } 

    _indexed_translations = { 
            "Left" : ("position", 0), 
            "Top" : ("position", 1), 
            "Width" : ("size", 0), 
            "Height" : ("size", 1), 
        } 

    _proxy_for = statictext.StaticText




log.debug("Registering VBStaticText as '%s'" % sys.modules[__name__].VBStaticText)
registry.getRegistry().register( sys.modules[__name__].VBStaticText )
