#!/usr/bin/python

__version__ = "0.1"

from wxPython import wx
from PythonCardPrototype import model, dialog
from vb2py import converter, vbparser, utils
import os
import time

class vb2pyOptions(model.CustomDialog):
    """GUI for the vb2Py converter"""
    
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
    
    def on_btnOK_mouseClick(self, event):
        """Pressed OK"""
        self.saveINI()
        self.Close()

    def saveINI(self):
        """Save the INI file"""
        self.log.info("Saving INI file")
        fle = open(utils.relativePath("vb2py.ini"), "w")
        try:
            fle.write(self.components.optionText.text)
            self.log.info("Succeeded!")
        finally:
            fle.close()
        
    def on_btnApply_mouseClick(self, event):
        """User clicked the apply button"""
        self.saveINI()
        self.parent.rereadOptions()
        self.parent.updateView()
        
if __name__ == '__main__':
    app = model.PythonCardApp(vb2pyOptions)
    app.MainLoop()
