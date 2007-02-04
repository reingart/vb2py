#@+leo-ver=4-thin
#@+node:pap.20070203003353:@thin vb2pyOptions.py

#@<< vb2pyOptions declarations >>
#@+node:pap.20070203003353.1:<< vb2pyOptions declarations >>
#!/usr/bin/python

__version__ = "0.1"

from wxPython import wx
from PythonCard import model, dialog
from vb2py import converter, vbparser, utils
import os
import time

#@-node:pap.20070203003353.1:<< vb2pyOptions declarations >>
#@nl
#@+others
#@+node:pap.20070203003353.2:class vb2pyOptions
class vb2pyOptions(model.CustomDialog):
    """GUI for the vb2Py converter"""
    #@	@+others
    #@+node:pap.20070203003353.3:__init__
    def __init__(self, logger, *args, **kw):
        """Load the options file"""
        model.CustomDialog.__init__(self, *args, **kw)
        self.log = logger
        self.log.info("Opening INI file")
        fle = open(utils.relativePath("vb2py.ini"), "r")
        try:
            text = fle.read()
            self.components.optionText.text = text
        finally:
            fle.close()
    #@-node:pap.20070203003353.3:__init__
    #@+node:pap.20070203003353.4:on_btnOK_mouseClick
    def on_btnOK_mouseClick(self, event):
        """Pressed OK"""
        self.saveINI()
        self.Close()
    #@-node:pap.20070203003353.4:on_btnOK_mouseClick
    #@+node:pap.20070203003353.5:saveINI
    def saveINI(self):
        """Save the INI file"""
        self.log.info("Saving INI file")
        fle = open(utils.relativePath("vb2py.ini"), "w")
        try:
            fle.write(self.components.optionText.text)
            self.log.info("Succeeded!")
        finally:
            fle.close()
    #@-node:pap.20070203003353.5:saveINI
    #@+node:pap.20070203003353.6:on_btnApply_mouseClick
    def on_btnApply_mouseClick(self, event):
        """User clicked the apply button"""
        self.saveINI()
        self.parent.rereadOptions()
        self.parent.updateView()
    #@-node:pap.20070203003353.6:on_btnApply_mouseClick
    #@-others
#@-node:pap.20070203003353.2:class vb2pyOptions
#@-others
        
if __name__ == '__main__':
    app = model.PythonCardApp(vb2pyOptions)
    app.MainLoop()
#@-node:pap.20070203003353:@thin vb2pyOptions.py
#@-leo
