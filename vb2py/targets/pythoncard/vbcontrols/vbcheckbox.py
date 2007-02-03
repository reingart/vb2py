# Created by Leo from: C:\Development\Python23\Lib\site-packages\vb2py\vb2py.leo

from vb2py.targets.pythoncard.controlclasses import VBWrapped, VBWidget
import vb2py.logger
log = vb2py.logger.getLogger("VBCheckBox")

from PythonCardPrototype.components import checkbox
from wxPython import wx
import sys
from PythonCardPrototype import binding, event, registry, widget


class VBCheckBox(VBWidget): 
    __metaclass__ = VBWrapped 

    _translations = { 
            "Text" : "text", 
            "Enabled" : "enabled", 
            "Visible" : "visible", 
            "Value" : "checked",
            "Caption" : "label",
        } 

    _indexed_translations = { 
            "Left" : ("position", 0), 
            "Top" : ("position", 1), 
            "Width" : ("size", 0), 
            "Height" : ("size", 1), 
        } 

    _proxy_for = checkbox.CheckBox


log.debug("Registering VBCheckBox as '%s'" % sys.modules[__name__].VBCheckBox)
registry.getRegistry().register( sys.modules[__name__].VBCheckBox )
