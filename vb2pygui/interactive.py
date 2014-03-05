#@+leo-ver=4-thin
#@+node:pap.20070203172243.1:@thin interactive.py

#@<< interactive declarations >>
#@+node:pap.20070203172243.2:<< interactive declarations >>
#!/usr/bin/python

from vb2py import converter, vbparser, config, utils
from vb2pyGUI import LogInterceptor

from PythonCard import model


#@-node:pap.20070203172243.2:<< interactive declarations >>
#@nl
#@+others
#@+node:pap.20070203172243.3:class Interactive
class Interactive(model.Background):
    #@	@+others
    #@+node:pap.20070203172243.4:on_initialize
    def on_initialize(self, event):
        self.log = converter.log = vbparser.log = LogInterceptor(self.logText) # Redirect logs to our window
        self.components.VB.SetLexerLanguage("vb")
        self.components.Python.SetLexerLanguage("python")
        self.components.VBPane.Raise()
    
    #@-node:pap.20070203172243.4:on_initialize
    #@+node:pap.20070203175538:on_Convert_mouseClick
    def on_Convert_mouseClick(self, event):
        """Convert the code"""
        self.components.VBPane.Raise()
        text = self.getSelectedText()
        self.logText("Converting code")
        print self.components.CodeContext.stringSelection
        try:
            py = vbparser.parseVB(text, container=self.getConversionContext())
            py_text = py.renderAsCode()
            self.components.Python.text = py_text
        except Exception, err:
            err_msg = "Unable to parse: '%s'" % err
            self.logText(err_msg)
            self.components.Python.text = err_msg
        else:
            self.logText("Succeeded!")
        
    #@-node:pap.20070203175538:on_Convert_mouseClick
    #@+node:pap.20070203185350:on_HelpConvertAs_mouseClick
    def on_HelpConvertAs_mouseClick(self, event):
        """Help on converting as"""
        print "hello"
    #@nonl
    #@-node:pap.20070203185350:on_HelpConvertAs_mouseClick
    #@+node:pap.20070203185511:on_HelpConversionStyle_mouseClick
    def on_HelpConversionStyle_mouseClick(self, event):
        """Help on converting as"""
        print "hello"
    #@nonl
    #@-node:pap.20070203185511:on_HelpConversionStyle_mouseClick
    #@+node:pap.20070203185706:on_CodeContext_mouseEnter
    def on_CodeContext_mouseEnter(self, event):
        """Mouse over the code context"""
        print "Over!"
    #@nonl
    #@-node:pap.20070203185706:on_CodeContext_mouseEnter
    #@+node:pap.20070203175538.1:getSelectedText
    def getSelectedText(self):
        """Return the highlighted text or the whole thing if not selected"""
        start, finish = self.components.VB.GetSelection()
        if start < finish:
            text = self.components.VB.text[start:finish]
        else:
            text = self.components.VB.text
        return text
    #@-node:pap.20070203175538.1:getSelectedText
    #@+node:pap.20070203191301:getConversionContext
    def getConversionContext(self):
        """Return the current conversion context"""
        contexts = {
            "Code module" : vbparser.VBCodeModule,
            "Class module" : vbparser.VBClassModule,
            "Form" : vbparser.VBFormModule,
        }
        return contexts[self.components.CodeContext.stringSelection]()
    #@nonl
    #@-node:pap.20070203191301:getConversionContext
    #@+node:pap.20070203175538.2:logText
    def logText(self, text):
        """Log some text"""
        self.components.LogWindow.AppendText("%s\n" % text)
    #@nonl
    #@-node:pap.20070203175538.2:logText
    #@-others
#@-node:pap.20070203172243.3:class Interactive
#@-others


if __name__ == '__main__':
    app = model.Application(Interactive)
    app.MainLoop()
#@-node:pap.20070203172243.1:@thin interactive.py
#@-leo
