from vb2py.targets.pythoncard.controlclasses import VBWrapped, VBWidget
import vb2py.logger
log = vb2py.logger.getLogger("VBBitmapCanvas")

from PythonCardPrototype.components import bitmapcanvas
from wxPython import wx
import sys
from PythonCardPrototype import binding, event, registry, widget


class VBBitmapCanvas(VBWidget): 
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

    _proxy_for = bitmapcanvas.BitmapCanvas

    _bitmap = None # Our bitmap
    Stretch = 0

    # << VBBitmapCanvas Methods >>
    def _setPicture(self, bitmap):
        """Setting the picture property"""
        if self.Stretch:
            try:
                bitmap.setSize(self.size)
            except NotImplementedError:
                log.error("PythonCard bitmap resize not implemented, Stretch mode will not work")
        else:
            self.size = bitmap.getSize()
        self.drawBitmap(bitmap)
        self.__dict__["_bitmap"] = bitmap

    def _getPicture(self):
        """Get the bitmap property"""
        return self._bitmap

    Picture = property(fget=_getPicture, fset=_setPicture)
    # -- end -- << VBBitmapCanvas Methods >>  

VBBitmapCanvas._setters["Picture"] = VBBitmapCanvas._setPicture	

log.debug("Registering VBBitmapCanvas as '%s'" % sys.modules[__name__].VBBitmapCanvas)
registry.getRegistry().register( sys.modules[__name__].VBBitmapCanvas )
