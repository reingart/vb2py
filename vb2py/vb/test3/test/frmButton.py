# Created by Leo from: C:\Development\Python23\Lib\site-packages\vb2py\vb2py.leo

"""The main form for the application"""

from PythonCardPrototype import model

# Allow importing of our custom controls
import PythonCardPrototype.res
PythonCardPrototype.res.APP_COMPONENTS_PACKAGE = "vb2py.targets.pythoncard.vbcontrols"

class Background(model.Background):

    def __getattr__(self, name):
        """If a name was not found then look for it in components"""
        return getattr(self.components, name)


    def __init__(self, *args, **kw):
        """Initialize the form"""
        model.Background.__init__(self, *args, **kw)
        # Call the VB Form_Load
        # TODO: This is brittle - depends on how the private indicator is set
        if hasattr(self, "_MAINFORM__Form_Load"):
            self._MAINFORM__Form_Load()
        elif hasattr(self, "Form_Load"):
            self.Form_Load()


from vb2py.vbfunctions import *
from vb2py.vbdebug import *
import Globals

class MAINFORM(Background):


    def on_Command1_mouseClick(self, *args):
        Globals.Log('Click')

    def on_Command1_gainFocus(self, *args):
        Globals.Log('Got focus')

    def on_Command1_loseFocus(self, *args):
        Globals.Log('LostFocus')

    def on_Command1_mouseDown(self, *args):
        Button, Shift, X, Y = vbGetEventArgs(["ButtonDown()", "ShiftDown()", "x", "y"], args)
        Globals.Log('MouseDown' + Str(Button) + ', ' + Str(Shift) + ', ' + Str(X) + ', ' + Str(Y))

    def on_Command1_mouseMove(self, *args):
        Button, Shift, X, Y = vbGetEventArgs(["ButtonDown()", "ShiftDown()", "x", "y"], args)
        Globals.Log('MouseMove' + Str(Button) + ', ' + Str(Shift) + ', ' + Str(X) + ', ' + Str(Y))

    def on_Command1_mouseUp(self, *args):
        Button, Shift, X, Y = vbGetEventArgs(["ButtonDown()", "ShiftDown()", "x", "y"], args)
        Globals.Log('MouseUp' + Str(Button) + ', ' + Str(Shift) + ', ' + Str(X) + ', ' + Str(Y))

    def on_Command10_mouseClick(self, *args):
        Globals.Log('Hit cancel')

    def on_Command11_mouseClick(self, *args):
        self.on_Command1_mouseMove(10, 20, 30, 40)
        self.on_Command1_mouseUp(10, 20, 30, 40)
        self.on_Command1_mouseDown(10, 20, 30, 40)

    def on_Command4_mouseClick(self, *args):
        self.Command1.Visible = not self.Command1.Visible

    def on_Command5_mouseClick(self, *args):
        self.Command5.Top = self.Command5.Top + 30
        self.Command5.Left = self.Command5.Left + 30

    def on_Command6_mouseClick(self, *args):
        self.Command6.Height = self.Command6.Height + 30
        self.Command6.Width = self.Command6.Width + 30

    def on_Command7_mouseClick(self, *args):
        self.Command1.Enabled = not self.Command1.Enabled

    def on_Command9_mouseClick(self, *args):
        Globals.Log('Hit default')

    # VB2PY (UntranslatedCode) Attribute VB_Name = "frmButton"
    # VB2PY (UntranslatedCode) Attribute VB_GlobalNameSpace = False
    # VB2PY (UntranslatedCode) Attribute VB_Creatable = False
    # VB2PY (UntranslatedCode) Attribute VB_PredeclaredId = True
    # VB2PY (UntranslatedCode) Attribute VB_Exposed = False



if __name__ == '__main__':
    app = model.PythonCardApp(MAINFORM)
    app.MainLoop()
