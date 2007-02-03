#!/usr/bin/python

__version__ = "0.2.1"

from wxPython import wx
from PythonCardPrototype import model, dialog
from vb2py import converter, vbparser, config, utils
import os
import time
import vb2pyOptions, finddialog
import webbrowser
import threading
import sys


ConversionContexts = {
    "menuCodeModuleContext" : vbparser.VBCodeModule,
    "menuClassModuleContext" : vbparser.VBClassModule,
    "menuFormModuleContext" : vbparser.VBFormModule,
}


class VB2PyIDE(model.Background):
    """GUI for the vb2Py converter"""
    
    def parseProject(self):
        """Populates the tree view with the parse results.
        """
        #
        # Select an appropriate parser
        if not self.projectFilename.lower().endswith(".vbp"):
            SelectedParser = converter.FileParser
            self.logText("Using File Parser")
        else:
            self.logText("Using Project Parser")
            SelectedParser = converter.ProjectParser
        try:
            self.project = SelectedParser(self.projectFilename)
            self.project.doParse()
        except Exception, err:
            self.logText("Error: %s (%s)" % ("Parsing %s failed:" % (self.projectFilename,), str(err)))
            return
        else:
            tree = self.components.parseTree
            # First tree level is the VB project file itself:
            rootNode = tree.AddRoot(os.path.basename(self.projectFilename))
            tree.SetPyData(rootNode, self.project)
            self.results = {} # Place to store the converted code
            #
            try:
                TargetResource = converter.importTarget("PythonCard")
                self.converter = converter.VBConverter(TargetResource, SelectedParser)
                self.converter.doConversion(self.projectFilename, callback=self.conversionProgress)
            except Exception, err:
                self.logText("Error: %s (%s)" % ("Parsing %s failed:" % (self.projectFilename,), str(err)))
                return
            #
            # TODO - take this out - just a quick test for PyFilling
            global conv
            conv = self.converter
            #
            # Second tree level is all the modules (forms):
            for resource in self.converter.resources:
                child = tree.AppendItem(rootNode, str(resource.name))
                self.results[resource.name] = resource
            self.tree = tree
        self.tree.Expand(rootNode)            
        self.tree.Show(True)
        self.conversionProgress("", 0)

    def conversionProgress(self, text, amount):
        """Report on progress"""
        self.components.prgProgress.value = amount
        self.components.txtStatus.text = text
        self.components.prgProgress.visible = amount > 0
        self.components.txtStatus.visible = amount > 0
        
    def on_openBackground(self, event):
        self.log = converter.log = vbparser.log = LogInterceptor(self.logText) # Redirect logs to our window
        self.conversion_context = vbparser.VBCodeModule
        self.current_resource = None
        self.find_text = ""
        self.find_language = "VB"
        self.setSize((1000, 800))
        #
        self.find = None
        self.testing = None
        self.options = None
        
    def on_vb2pyGUI_size(self, event):
        """Resize the window"""
        try:
            width, height = event.size
            self.panel.SetSize((width,height))
            frame = 2; middle = 6; topping = 38 # TODO: Calculate these?
            height = height - topping - frame - frame - self.components.logWindow.size[1] - self.components.prgProgress.size[1] - 14
            width = ((width - middle)//4 - frame)
            self.components.parseTree.position = (frame, 0)
            self.components.parseTree.size = (width, height)
            height = height // 2
            self.components.vbText.position = (width + middle , 0)
            self.components.vbText.size = (width * 3, height)
            self.components.pythonText.position = (width + middle , height)
            self.components.pythonText.size = (width * 3, height)
            self.components.logWindow.size = (event.size[0]-4*frame, self.components.logWindow.size[1])
            self.components.logWindow.position = (0, self.components.parseTree.size[1])
            self.components.prgProgress.size = (self.components.logWindow.size[0], self.components.prgProgress.size[1])
            self.components.prgProgress.position = (self.components.prgProgress.position[0], self.components.parseTree.size[1]+self.components.logWindow.size[1])
            self.components.txtStatus.position = (frame, self.components.prgProgress.position[1])
        except Exception, err:
            self.logText("Error resizing: '%s'" % err)
        
        
    def on_menuFileOpen_select(self, event):
        """Opens a VB project file (.vbp) and extracts the paths of files to analyze."""
        reply = dialog.openFileDialog(title="Select a VB project file", 
                                      wildcard="VB Project Files (*.vbp)|*.vbp|" \
                                               "Form Files (*.frm)|*.frm|" \
                                               "Class Modules (*.cls)|*.cls|" \
                                               "Code Modules (*.bas)|*.bas|" \
                                               "All Files (*.*)|*.*")
        if reply["accepted"]:
            self.components.parseTree.DeleteAllItems()
            self.projectFilename = reply["paths"][0]
            #
            self.components.vbText.text = ""
            self.components.pythonText.text = ""
            #
            # If on Windows it looks like we can do this in the background
            # but it doesn't work on Linux
            if sys.platform == "win32":
                thread = threading.Thread(target=self.parseProject)
                thread.start()
            else:
                self.parseProject()

    def on_menuViewStructure_select(self, event):
        """View the structure of the converted code using PyFilling"""
        import wx.py.PyFilling, wx.py.introspect, inspect    
        _getAttributeNames = wx.py.introspect.getAttributeNames 
        def getGoodChildren(obj):
            """Only get things that we are interested in"""
            names = _getAttributeNames(obj)
            goodnames = []
            for name in goodnames:
                try:
                    if (not name.startswith("_")) and (not inspect.ismethod(getattr(obj, name))) \
                            and (not inspect.isfunction(getattr(obj, name))):
                        goodnames.append(name)
                    else:
                        print "Threw out name: '%s'" % name
                except:
                    pass
            return goodnames
        wx.py.introspect.getAttributeNames = getGoodChildren
        wx.py.PyFilling.main()
        
    def on_menuFileSave_select(self, event):
        """Save the converted code"""
        try:
            default_dir = os.path.split(self.projectFilename)[0]
        except AttributeError:
            default_dir = None
        reply = dialog.directoryDialog(None, "Select a folder for the results:",
                                       default_dir)
        if reply["accepted"]:
            self.outFolder = reply["path"]
            converter.renderTo(self.converter, self.outFolder, do_code=1)
            
    def on_menuFileExit_select(self, event):
        """Close the application"""
        if self.find: self.find.Destroy()
        if self.options: self.options.Destroy()
        self.Destroy()

    def on_parseTree_selectionChanged(self, event):
        """Change the view"""
        name = self.tree.GetItemText(event.GetItem())
        try:
            resource = self.results[name]
        except KeyError:
            return
        self.current_resource = resource
        self.updateView()

    def updateView(self):
        """Update the current view"""
        if self.current_resource == "Text":
            self.on_menuConvert_select(None)
        elif self.current_resource:
            self.components.pythonText.text = vbparser.renderCodeStructure(self.current_resource.code_structure)
            self.components.vbText.text = self.current_resource.code_block

    def on_menuAbout_select(self, event):
        """User clicked on the help .. about menu"""
        reply = dialog.alertDialog(None, 
                                   "A Visual Basic to Python conversion toolkit\nVersion %s (GUI v%s)" % (
                                       converter.__version__, __version__),
                                   "About vb2Py")
    def logText(self, text):
        """Log some text to the log window"""
        self.components.logWindow.AppendText("%s\n" % text)

    def on_menuOptions_select(self, event):
        """User clicked on the view ... options menu"""
        if not self.options:
            self.options = vb2pyOptions.vb2pyOptions(self.log, self)
        self.options.Show()
        self.rereadOptions()
        self.updateView()

    def rereadOptions(self):
        """Re-read the options"""
        self.logText("Re-reading options now")
        converter.Config.initConfig()
        vbparser.Config.initConfig()
        self.logText("Succeeded!")

    def on_menuConvert_select(self, event):
        """User clicked on the convert ... convert menu"""
        self.logText("Converting active VB window")
        text = self.components.vbText.text
        self.convertText(text)

    def convertText(self, text):
        """Convert some text to Python"""
        try:
            py = vbparser.parseVB(text, container=self.conversion_context())
            py_text = py.renderAsCode()
            self.components.pythonText.text = py_text
        except Exception, err:
            err_msg = "Unable to parse: '%s'" % err
            self.logText(err_msg)
            self.components.pythonText.text = err_msg
        else:
            self.logText("Succeeded!")
        self.current_resource = "Text"

    def on_menuConvertSelection_select(self, event):
        """User clicked on the convert ... selection menu"""
        self.logText("Converting selection")
        start, finish = self.components.vbText.GetSelection()
        text = self.components.vbText.text[start:finish]
        self.convertText(text)
        
    def on_menuClassModuleContext_select(self, event):
        """Set the context of the conversion"""
        self.updateContext("menuClassModuleContext")
    def on_menuCodeModuleContext_select(self, event):
        """Set the context of the conversion"""
        self.updateContext("menuCodeModuleContext")
    def on_menuFormModuleContext_select(self, event):
        """Set the context of the conversion"""
        self.updateContext("menuFormModuleContext")

    def on_menuHelp_select(self, event):
        """User clicked on the help menu item"""
        webbrowser.open(utils.relativePath("doc/index.html"))

    def updateContext(self, newcontext):
        """Clear all checks from context menus"""
        for menu in ConversionContexts:
            if menu == newcontext:
                self.menuBar.setChecked(menu, 1)
                self.conversion_context = ConversionContexts[menu]
            else:
                self.menuBar.setChecked(menu, 0)
            
    def on_menuHelpGUI_select(self, event):
        """User clicked on the help menu item"""
        webbrowser.open(utils.relativePath("doc/index.html"))

    def on_menuFind_select(self, event):
        """User click on Edit ... find"""
        if not self.find:
            self.find = finddialog.FindDialog(self)
        self.find.Show()

    def on_menuFindNext_select(self, event):
        """User click on Edit ... find"""
        self.findText(self.find_text, self.find_language, next=1)

    def findText(self, search_text, language, next=0):
        """Find some text"""
        if language == "VB":
            control = self.components.vbText
        else:
            control = self.components.pythonText
        text = control.text.lower() # Case insensitive
        if next:
            current, end = control.GetSelection()
            posn = text[end:].find(search_text.lower())
        else:
            posn = text.find(search_text.lower())
            end = 0
        if posn == -1:
            self.logText("Search text not found")
        else:
            self.logText("Found at %d" % posn)
            control.SetFocus()
            control.SetSelection(posn+end, posn+end+len(search_text))
        
    def on_menuTestView_select(self, event):
        """The user clicked on the test window menu item"""
        launchApp("testingview.py")
        
        
class LogInterceptor:
    """Intercept logging calls and send them to the log window"""

    def __init__(self, callback):
        """Initialize"""
        self.callback = callback
        callback("Started logging (%s)" % time.ctime())

    def __getattr__(self, name):
        """Catch attr gets"""
        def logTo(*args):
            self.callback("%s: %s" % (name, args[-1]))
        return logTo
    

def launchApp(filename):
    """Launch another app - code taken from the Pythoncard samples"""
    import sys, os, wx
    args = []
    python = sys.executable
    if ' ' in python:
        pythonQuoted = '"' + python + '"'
    else:
        pythonQuoted = python
    if sys.platform.startswith('win'):
        os.spawnv(os.P_NOWAIT, python, [pythonQuoted, filename] + args)
    elif wx.wxPlatform == '__WXMAC__':
        # this is a bad hack to deal with the user starting
        # samples.py from the Finder
        if sys.executable == '/':
            python = '/Applications/Python.app/Contents/MacOS/python'
        os.spawnv(os.P_NOWAIT, python, [pythonQuoted, filename] + args)
    else:
        os.spawnv(os.P_NOWAIT, python, [pythonQuoted, filename] + args)
    
if __name__ == '__main__':
    app = model.PythonCardApp(VB2PyIDE)
    app.MainLoop()
