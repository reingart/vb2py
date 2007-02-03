# Created by Leo from: C:\Development\Python23\Lib\site-packages\vb2py\vb2py.leo

from vb2py.targets.pythoncard.controlclasses import VBWrapped, VBWidget
import vb2py.logger
log = vb2py.logger.getLogger("VBImageList")

from PythonCardPrototype.components import statictext
from wxPython import wx
import sys
from PythonCardPrototype import binding, event, registry, widget


class VBImageList(VBWidget): 
    __metaclass__ = VBWrapped 

    _translations = { 
            "ListImages" : "items",
    } 

    _name_to_method_translations = {
            "ListCount" : "getNumber",
            "ListIndex" : "getSelectionIndex",
    }

    _indexed_translations = { 
    } 

    _method_translations = {			
    }

    _proxy_for = statictext.StaticText # Not a PythonCard object at all but this at least works!

    # << VBImageList methods >>
    pass
    # -- end -- << VBImageList methods >>   

log.debug("Registering VBImageList as '%s'" % sys.modules[__name__].VBImageList)
registry.getRegistry().register( sys.modules[__name__].VBImageList )
