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


    def on_Command4_mouseClick(self, *args):
        self.Text3.Visible = not self.Text3.Visible

    def on_Command5_mouseClick(self, *args):
        self.Text3.Top = self.Text3.Top + 50
        self.Text3.Left = self.Text3.Left + 50

    def on_Command6_mouseClick(self, *args):
        self.Text3.Width = self.Text3.Width + 50
        self.Text3.Height = self.Text3.Height + 50

    def on_Command7_mouseClick(self, *args):
        self.Text3.Enabled = not self.Text3.Enabled

    def on_Text1_textUpdate(self, *args):
        Globals.Log('Change, \'' + self.Text1.Text + '\'')

    def on_Text1_mouseClick(self, *args):
        Globals.Log('Click')

    def __Text1_DblClick(self):
        Globals.Log('DblClick')

    def on_Text1_gainFocus(self, *args):
        Globals.Log('GotFocus')

    def on_Text1_keyDown_NOTSUPPORTED(self, *args):
        Globals.Log('Keydown' + ', ' + Str(KeyCode) + ', ' + Str(Shift))

    def on_Text1_keyPress_NOTSUPPORTED(self, *args):
        Globals.Log('KeyPress' + ', ' + Str(KeyCode) + ', ' + Str(Shift) + ', ' + self.Text1.Text)

    def on_Text1_keyUp_NOTSUPPORTED(self, *args):
        Globals.Log('KeyUp' + ', ' + Str(KeyCode) + ', ' + Str(Shift))

    def on_Text1_loseFocus(self, *args):
        Globals.Log('LostFocus')

    def on_Text1_mouseDown(self, *args):
        Button, Shift, X, Y = vbGetEventArgs(["ButtonDown()", "ShiftDown()", "x", "y"], args)
        Globals.Log('MouseDown' + Str(Button) + ', ' + Str(Shift) + ', ' + Str(X) + ', ' + Str(Y))

    def on_Text1_mouseMove(self, *args):
        Button, Shift, X, Y = vbGetEventArgs(["ButtonDown()", "ShiftDown()", "x", "y"], args)
        Globals.Log('MouseMove' + Str(Button) + ', ' + Str(Shift) + ', ' + Str(X) + ', ' + Str(Y))

    def on_Text1_mouseUp(self, *args):
        Button, Shift, X, Y = vbGetEventArgs(["ButtonDown()", "ShiftDown()", "x", "y"], args)
        Globals.Log('MouseUp' + Str(Button) + ', ' + Str(Shift) + ', ' + Str(X) + ', ' + Str(Y))

    def on_Text4_textUpdate(self, *args):
        Globals.Log('Change, \'' + self.Text4.Text + '\'')

    # VB2PY (UntranslatedCode) Attribute VB_Name = "frmTextBox"
    # VB2PY (UntranslatedCode) Attribute VB_GlobalNameSpace = False
    # VB2PY (UntranslatedCode) Attribute VB_Creatable = False
    # VB2PY (UntranslatedCode) Attribute VB_PredeclaredId = True
    # VB2PY (UntranslatedCode) Attribute VB_Exposed = False



if __name__ == '__main__':
    app = model.PythonCardApp(MAINFORM)
    app.MainLoop()
