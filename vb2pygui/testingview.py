#!/usr/bin/python

"""
__version__ = "$Revision: 1.1 $"
__date__ = "$Date: 2004/01/26 07:21:13 $"
"""

from PythonCardPrototype import model
from vb2py.test.scripttest import testCode
from vb2py.vbparser import convertVBtoPython, VBCodeModule

class TestingView(model.Background):

    left_pc = 33
    mid_pc = 33
    panel_top = 25
    panel_height_border = 80
    panel_width_border = 20
    label_top = 5
    label_left = 5
    
    def on_openBackground(self, event):
        # if you have any initialization
        # including sizer setup, do it here
        self._positionWidgets()

    def _positionWidgets(self, event=None):
        """Put widgets in the right place depending on the screen size"""
        if event:
            width, height = event.size
        else:
            width, height = self.panel.GetSize()
            height += 50 # Why do we have to do this?
            width += 5
        #
        # Get relative sizes
        width1 = width*self.left_pc/100.0
        width2 = width*(self.left_pc + self.mid_pc)/100.0
        #
        # Main panel size
        self.panel.SetSize((width,height))
        #
        # Label positions
        self.components.VBLabel.position = (self.label_left, self.label_top)
        self.components.PythonLabel.position = (self.label_left+width1, self.label_top)
        self.components.ResultLabel.position = (self.label_left+width2, self.label_top)
        #
        # Panels
        self.components.VBCodeEditor.position = (self.label_left, self.panel_top)
        self.components.VBCodeEditor.size = (width1, height-self.panel_height_border)
        #
        self.components.PythonCodeEditor.position = (self.label_left+width1, self.panel_top)
        self.components.PythonCodeEditor.size = (width2-width1, height-self.panel_height_border)
        #
        self.components.ResultsView.position = (self.label_left+width2, self.panel_top)
        self.components.ResultsView.size = (width-width2-self.panel_width_border, height-self.panel_height_border)

    def on_TestingView_size(self, event):
        """Resize event"""
        self._positionWidgets(event)
        
    def on_menuFileExit_select(self, event):
        """Quit this window"""
        self.Close()

    def on_menuTestCode_select(self, event):
        """Test the code"""
        vb = self.components.VBCodeEditor.text
        self.components.PythonCodeEditor.text = convertVBtoPython(vb, container=VBCodeModule())
        try:
            output = testCode(vb, verbose=2)
        except Exception, err:
            self.components.ResultsView.text = str(err)
        else:
            self.components.ResultsView.text = "\n".join(output)

if __name__ == '__main__':
    app = model.PythonCardApp(TestingView)
    app.MainLoop()
