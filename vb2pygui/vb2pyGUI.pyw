#@+leo-ver=4-thin
#@+node:pap.20070203003154.1:@thin vb2pyGUI.pyw

#@<< vb2pyGUI declarations >>
#@+node:pap.20070203003154.2:<< vb2pyGUI declarations >>
#!/usr/bin/python

__version__ = "0.2.2"

from wxPython import wx
from PythonCard import model, dialog
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


#@-node:pap.20070203003154.2:<< vb2pyGUI declarations >>
#@nl
#@+others
#@+node:pap.20070203003154.3:class VB2PyIDE
class VB2PyIDE(model.Background):
    """GUI for the vb2Py converter"""
    #@	@+others
    #@+node:pap.20070203003154.6:on_initialize
    def on_initialize(self, event):
        self.log = converter.log = vbparser.log = LogInterceptor(self.logText) # Redirect logs to our window
        self.logText("Created form")
        self.conversion_context = vbparser.VBCodeModule
        self.current_resource = None
        self.find_text = ""
        self.find_language = "VB"
        self.SetSize((1000, 800))
        self.resizeWindow(1000, 800)
        #
        self.find = None
        self.testing = None
        self.options = None
    #@-node:pap.20070203003154.6:on_initialize
    #@+node:pap.20070203003154.4:parseProject
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
    #@-node:pap.20070203003154.4:parseProject
    #@+node:pap.20070203003154.5:conversionProgress
    def conversionProgress(self, text, amount):
        """Report on progress"""
        self.components.prgProgress.value = amount
        self.components.txtStatus.text = text
        self.components.prgProgress.visible = amount > 0
        self.components.txtStatus.visible = amount > 0
    #@-node:pap.20070203003154.5:conversionProgress
    #@+node:pap.20070203003154.7:on_vb2pyGUI_size
    def on_vb2pyGUI_size(self, event):
        """Resize the window"""
        self.resizeWindow(*event.size)
    
    #@-node:pap.20070203003154.7:on_vb2pyGUI_size
    #@+node:pap.20070203103015:resizeWindow
    def resizeWindow(self, width, height):
        """Resize the main window and its child controls"""
        try:
            self.panel.SetSize((width, height))
            frame = 2; middle = 6; topping = 38 # TODO: Calculate these?
            h = height - topping - frame - frame - self.components.logWindow.size[1] - self.components.prgProgress.size[1] - 14
            w = ((width - middle)//4 - frame)
            self.components.parseTree.position = (frame, 0)
            self.components.parseTree.size = (w, h)
            h = h // 2
            self.components.vbText.position = (w + middle , 0)
            self.components.vbText.size = (w * 3, h)
            self.components.pythonText.position = (w + middle , h)
            self.components.pythonText.size = (w * 3, h)
            self.components.logWindow.size = (width-4*frame, self.components.logWindow.size[1])
            self.components.logWindow.position = (0, self.components.parseTree.size[1])
            self.components.prgProgress.size = (self.components.logWindow.size[0], self.components.prgProgress.size[1])
            self.components.prgProgress.position = (self.components.prgProgress.position[0], self.components.parseTree.size[1]+self.components.logWindow.size[1])
            self.components.txtStatus.position = (frame, self.components.prgProgress.position[1])
        except Exception, err:
            self.logText("Error resizing: '%s'" % err)
    #@-node:pap.20070203103015:resizeWindow
    #@+node:pap.20070203003154.8:on_menuFileOpen_select
    def on_menuFileOpen_select(self, event):
        """Opens a VB project file (.vbp) and extracts the paths of files to analyze."""
        reply = dialog.openFileDialog(title="Select a VB project file", 
                                      wildcard="VB Project Files (*.vbp)|*.vbp|" \
                                               "Form Files (*.frm)|*.frm|" \
                                               "Class Modules (*.cls)|*.cls|" \
                                               "Code Modules (*.bas)|*.bas|" \
                                               "All Files (*.*)|*.*")
        if reply.accepted:
            self.components.parseTree.DeleteAllItems()
            self.projectFilename = reply.paths[0]
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
    #@-node:pap.20070203003154.8:on_menuFileOpen_select
    #@+node:pap.20070203003154.9:on_menuViewStructure_select
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
    #@-node:pap.20070203003154.9:on_menuViewStructure_select
    #@+node:pap.20070203003154.10:on_menuFileSave_select
    def on_menuFileSave_select(self, event):
        """Save the converted code"""
        try:
            default_dir = os.path.split(self.projectFilename)[0]
        except AttributeError:
            default_dir = None
        reply = dialog.directoryDialog(None, "Select a folder for the results:",
                                       default_dir)
        if reply.accepted:
            self.outFolder = reply.path
            converter.renderTo(self.converter, self.outFolder, do_code=1)
    #@-node:pap.20070203003154.10:on_menuFileSave_select
    #@+node:pap.20070203003154.11:on_menuFileExit_select
    def on_menuFileExit_select(self, event):
        """Close the application"""
        if self.find: self.find.Destroy()
        if self.options: self.options.Destroy()
        self.Destroy()
    #@-node:pap.20070203003154.11:on_menuFileExit_select
    #@+node:pap.20070203003154.12:on_parseTree_selectionChanged
    def on_parseTree_selectionChanged(self, event):
        """Change the view"""
        name = self.tree.GetItemText(event.GetItem())
        try:
            resource = self.results[name]
        except KeyError:
            return
        self.current_resource = resource
        self.updateView()
    #@-node:pap.20070203003154.12:on_parseTree_selectionChanged
    #@+node:pap.20070203003154.13:updateView
    def updateView(self):
        """Update the current view"""
        if self.current_resource == "Text":
            self.on_menuConvert_select(None)
        elif self.current_resource:
            self.components.pythonText.text = vbparser.renderCodeStructure(self.current_resource.code_structure)
            self.components.vbText.text = self.current_resource.code_block
    #@-node:pap.20070203003154.13:updateView
    #@+node:pap.20070203003154.14:on_menuAbout_select
    def on_menuAbout_select(self, event):
        """User clicked on the help .. about menu"""
        reply = dialog.alertDialog(None, 
                                   "A Visual Basic to Python conversion toolkit\nVersion %s (GUI v%s)" % (
                                       converter.__version__, __version__),
                                   "About vb2Py")
    #@-node:pap.20070203003154.14:on_menuAbout_select
    #@+node:pap.20070203003154.15:logText
    def logText(self, text):
        """Log some text to the log window"""
        self.components.logWindow.AppendText("%s\n" % text)
    #@-node:pap.20070203003154.15:logText
    #@+node:pap.20070203003154.16:on_menuOptions_select
    def on_menuOptions_select(self, event):
        """User clicked on the view ... options menu"""
        if not self.options:
            self.options = vb2pyOptions.vb2pyOptions(self.log, self)
        self.options.Show()
        self.rereadOptions()
        self.updateView()
    #@-node:pap.20070203003154.16:on_menuOptions_select
    #@+node:pap.20070203003154.17:rereadOptions
    def rereadOptions(self):
        """Re-read the options"""
        self.logText("Re-reading options now")
        converter.Config.initConfig()
        vbparser.Config.initConfig()
        self.logText("Succeeded!")
    #@-node:pap.20070203003154.17:rereadOptions
    #@+node:pap.20070203003154.18:on_menuConvert_select
    def on_menuConvert_select(self, event):
        """User clicked on the convert ... convert menu"""
        self.logText("Converting active VB window")
        text = self.components.vbText.text
        self.convertText(text)
    #@-node:pap.20070203003154.18:on_menuConvert_select
    #@+node:pap.20070203003154.19:convertText
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
    #@-node:pap.20070203003154.19:convertText
    #@+node:pap.20070203003154.20:on_menuConvertSelection_select
    def on_menuConvertSelection_select(self, event):
        """User clicked on the convert ... selection menu"""
        self.logText("Converting selection")
        start, finish = self.components.vbText.GetSelection()
        text = self.components.vbText.text[start:finish]
        self.convertText(text)
    #@-node:pap.20070203003154.20:on_menuConvertSelection_select
    #@+node:pap.20070203003154.21:on_menuClassModuleContext_select
    def on_menuClassModuleContext_select(self, event):
        """Set the context of the conversion"""
        self.updateContext("menuClassModuleContext")
    #@-node:pap.20070203003154.21:on_menuClassModuleContext_select
    #@+node:pap.20070203003154.22:on_menuCodeModuleContext_select
    def on_menuCodeModuleContext_select(self, event):
        """Set the context of the conversion"""
        self.updateContext("menuCodeModuleContext")
    #@-node:pap.20070203003154.22:on_menuCodeModuleContext_select
    #@+node:pap.20070203003154.23:on_menuFormModuleContext_select
    def on_menuFormModuleContext_select(self, event):
        """Set the context of the conversion"""
        self.updateContext("menuFormModuleContext")
    #@-node:pap.20070203003154.23:on_menuFormModuleContext_select
    #@+node:pap.20070203003154.24:on_menuHelp_select
    def on_menuHelp_select(self, event):
        """User clicked on the help menu item"""
        webbrowser.open(utils.relativePath("doc/index.html"))
    #@-node:pap.20070203003154.24:on_menuHelp_select
    #@+node:pap.20070203003154.25:updateContext
    def updateContext(self, newcontext):
        """Clear all checks from context menus"""
        for menu in ConversionContexts:
            if menu == newcontext:
                self.menuBar.setChecked(menu, 1)
                self.conversion_context = ConversionContexts[menu]
            else:
                self.menuBar.setChecked(menu, 0)
    #@-node:pap.20070203003154.25:updateContext
    #@+node:pap.20070203003154.26:on_menuHelpGUI_select
    def on_menuHelpGUI_select(self, event):
        """User clicked on the help menu item"""
        webbrowser.open(utils.relativePath("doc/index.html"))
    #@-node:pap.20070203003154.26:on_menuHelpGUI_select
    #@+node:pap.20070203003154.27:on_menuFind_select
    def on_menuFind_select(self, event):
        """User click on Edit ... find"""
        if not self.find:
            self.find = finddialog.FindDialog(self)
        self.find.Show()
    #@-node:pap.20070203003154.27:on_menuFind_select
    #@+node:pap.20070203003154.28:on_menuFindNext_select
    def on_menuFindNext_select(self, event):
        """User click on Edit ... find"""
        self.findText(self.find_text, self.find_language, next=1)
    #@-node:pap.20070203003154.28:on_menuFindNext_select
    #@+node:pap.20070203003154.29:findText
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
    #@-node:pap.20070203003154.29:findText
    #@+node:pap.20070203003154.30:on_menuTestView_select
    def on_menuTestView_select(self, event):
        """The user clicked on the test window menu item"""
        launchApp("testingview.py")
    #@-node:pap.20070203003154.30:on_menuTestView_select
    #@-others
#@-node:pap.20070203003154.3:class VB2PyIDE
#@+node:pap.20070203003154.31:class LogInterceptor
        
        
class LogInterceptor:
    """Intercept logging calls and send them to the log window"""
    #@	@+others
    #@+node:pap.20070203003154.32:__init__
    def __init__(self, callback):
        """Initialize"""
        self.callback = callback
        callback("Started logging (%s)" % time.ctime())
    #@-node:pap.20070203003154.32:__init__
    #@+node:pap.20070203003154.33:__getattr__
    def __getattr__(self, name):
        """Catch attr gets"""
        def logTo(*args):
            self.callback("%s: %s" % (name, args[-1]))
        return logTo
    #@-node:pap.20070203003154.33:__getattr__
    #@-others
#@-node:pap.20070203003154.31:class LogInterceptor
#@+node:pap.20070203003154.34:launchApp
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
#@-node:pap.20070203003154.34:launchApp
#@-others
    
if __name__ == '__main__':
    app = model.Application(VB2PyIDE)
    app.MainLoop()
#@-node:pap.20070203003154.1:@thin vb2pyGUI.pyw
#@-leo
