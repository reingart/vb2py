#@+leo-ver=4-thin
#@+node:pap.20070203003236:@thin finddialog.py
#@<< finddialog declarations >>
#@+node:pap.20070203003236.1:<< finddialog declarations >>
#!/usr/bin/python

__version__ = "0.1"

from wxPython import wx
from PythonCard import model, dialog
from vb2py import converter, vbparser, utils
import os
import time

#@-node:pap.20070203003236.1:<< finddialog declarations >>
#@nl
#@+others
#@+node:pap.20070203003236.2:class FindDialog
class FindDialog(model.CustomDialog):
    """Find dialog"""
    #@	<< class FindDialog declarations >>
    #@+node:pap.20070203003236.3:<< class FindDialog declarations >>
    options = ["Python", "VB"]
    
    #@-node:pap.20070203003236.3:<< class FindDialog declarations >>
    #@nl
    #@	@+others
    #@+node:pap.20070203003236.4:__init__
    def __init__(self, *args, **kw):
        """Load the options file"""
        model.CustomDialog.__init__(self, *args, **kw)
        self.components.optSearchIn.SetSelection(self.options.index(self.parent.find_language))
        self.components.txtFind.text = self.parent.find_text
        self.components.txtFind.SetFocus()
        self.components.txtFind.SetSelection(0, len(self.parent.find_text))
    #@-node:pap.20070203003236.4:__init__
    #@+node:pap.20070203003236.5:on_btnFind_mouseClick
    def on_btnFind_mouseClick(self, event):
        """Click the find button"""
        self.parent.find_text = self.components.txtFind.text
        self.parent.find_language = self.options[self.components.optSearchIn.GetSelection()]
        self.parent.findText(self.parent.find_text, self.parent.find_language)
    #@-node:pap.20070203003236.5:on_btnFind_mouseClick
    #@+node:pap.20070203003236.6:on_btnFindNext_mouseClick
    def on_btnFindNext_mouseClick(self, event):
        """Click the find next button"""
        self.parent.find_text = self.components.txtFind.text
        self.parent.find_language = self.options[self.components.optSearchIn.GetSelection()]
        self.parent.findText(self.parent.find_text, self.parent.find_language, next=1)
    #@-node:pap.20070203003236.6:on_btnFindNext_mouseClick
    #@-others
#@-node:pap.20070203003236.2:class FindDialog
#@-others


       
if __name__ == '__main__':
    app = model.PythonCardApp(FindDialog)
    app.MainLoop()
#@-node:pap.20070203003236:@thin finddialog.py
#@-leo
