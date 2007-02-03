# Created by Leo from: C:\Development\Python23\Lib\site-packages\vb2py\vb2py.leo

from vb2py.targets.pythoncard.controlclasses import VBWrapped, VBWidget
import vb2py.logger
log = vb2py.logger.getLogger("VBTimer")

from PythonCardPrototype.components import statictext
from wxPython import wx
import sys
from PythonCardPrototype import binding, event, registry, widget


class VBTimer(VBWidget): 
    __metaclass__ = VBWrapped 

    _name_to_method_translations = {
            "Enabled" : ("_getEnabled", "_setEnabled"),
    }

    def _setEnabled(self, value):
        """Setting enabled state"""
        # Call the event handler
        getattr(self.GetGrandParent(), "%s_Timer" % self.name)()

    def _getEnabled(self):
        """Getting enabled state"""
        print "Inside getenabled"
        return True

    _proxy_for = statictext.StaticText

log.debug("Registering VBTimer as '%s'" % sys.modules[__name__].VBTimer)
registry.getRegistry().register( sys.modules[__name__].VBTimer )
