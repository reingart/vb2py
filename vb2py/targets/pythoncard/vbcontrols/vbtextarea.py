# Created by Leo from: C:\Development\Python23\Lib\site-packages\vb2py\vb2py.leo

from vb2py.targets.pythoncard.controlclasses import VBWrapped, VBWidget
import vb2py.logger
log = vb2py.logger.getLogger("VBTextArea")

from PythonCardPrototype.components import textarea
from wxPython import wx
import sys
from PythonCardPrototype import binding, event, registry, widget


class VBTextArea(VBWidget): 
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

    _proxy_for = textarea.TextArea


log.debug("Registering VBTextArea as '%s'" % sys.modules[__name__].VBTextArea)
registry.getRegistry().register( sys.modules[__name__].VBTextArea )
