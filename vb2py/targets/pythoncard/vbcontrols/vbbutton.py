# Created by Leo from: C:\Development\Python23\Lib\site-packages\vb2py\vb2py.leo

from vb2py.targets.pythoncard.controlclasses import VBWrapped, VBWidget
import vb2py.logger
log = vb2py.logger.getLogger("VBButton")

from PythonCardPrototype.components import button
from wxPython import wx
import sys
from PythonCardPrototype import binding, event, registry, widget


class VBButton(VBWidget): 
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

    _proxy_for = button.Button 


log.debug("Registering VBButton as '%s'" % sys.modules[__name__].VBButton)
registry.getRegistry().register( sys.modules[__name__].VBButton )
