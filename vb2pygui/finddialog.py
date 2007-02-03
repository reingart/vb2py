#!/usr/bin/python

__version__ = "0.1"

from wxPython import wx
from PythonCardPrototype import model, dialog
from vb2py import converter, vbparser, utils
import os
import time

class FindDialog(model.CustomDialog):
    """Find dialog"""

    options = ["Python", "VB"]
    
    def __init__(self, *args, **kw):
        """Load the options file"""
        model.CustomDialog.__init__(self, *args, **kw)
        self.components.optSearchIn.SetSelection(self.options.index(self.parent.find_language))
        self.components.txtFind.text = self.parent.find_text
        self.components.txtFind.SetFocus()
        self.components.txtFind.SetSelection(0, len(self.parent.find_text))
    
    def on_btnFind_mouseClick(self, event):
        """Click the find button"""
        self.parent.find_text = self.components.txtFind.text
        self.parent.find_language = self.options[self.components.optSearchIn.GetSelection()]
        self.parent.findText(self.parent.find_text, self.parent.find_language)
        
    def on_btnFindNext_mouseClick(self, event):
        """Click the find next button"""
        self.parent.find_text = self.components.txtFind.text
        self.parent.find_language = self.options[self.components.optSearchIn.GetSelection()]
        self.parent.findText(self.parent.find_text, self.parent.find_language, next=1)


       
if __name__ == '__main__':
    app = model.PythonCardApp(FindDialog)
    app.MainLoop()
