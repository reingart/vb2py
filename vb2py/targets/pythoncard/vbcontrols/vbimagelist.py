from vb2py.targets.pythoncard.controlclasses import VBWrapped, VBWidget
from vb2py.targets.pythoncard import Register
import vb2py.logger
log = vb2py.logger.getLogger("VBImageList")

from PythonCard.components import statictext
import wx
import sys
from PythonCard import event, registry, widget


class VBImageList(VBWidget): 
    __metaclass__ = VBWrapped 

    _translations = { 
            "ListImages" : "items",
    } 

    _name_to_method_translations = {
            "ListCount" : ("getNumber", None),
            "ListIndex" : ("getSelectionIndex", None),
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
Register(VBImageList)
