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
        self.Combo1.Visible = not self.Combo1.Visible

    def on_Command5_mouseClick(self, *args):
        self.Combo1.Left = self.Combo1.Left + 50
        self.Combo1.Top = self.Combo1.Top + 50

    def on_Command6_mouseClick(self, *args):
        self.Combo1.Width = self.Combo1.Width + 50

    def on_Command7_mouseClick(self, *args):
        self.Combo1.Enabled = not self.Combo1.Enabled

    def on_Combo1_textUpdate(self, *args):
        Globals.Log('Change, \'' + self.Combo1.Text + '\'')

    def on_Combo1_mouseClick(self, *args):
        Globals.Log('Click')

    def on_Combo1_mouseDoubleClick(self, *args):
        Globals.Log('DblClick')

    def on_Combo1_gainFocus(self, *args):
        Globals.Log('GotFocus')

    def on_Combo1_keyDown_NOTSUPPORTED(self, *args):
        Globals.Log('Keydown' + ', ' + Str(KeyCode) + ', ' + Str(Shift))

    def on_Combo1_keyPress_NOTSUPPORTED(self, *args):
        Globals.Log('KeyPress' + ', ' + Str(KeyCode) + ', ' + Str(Shift) + ', ' + self.Combo1.Text)

    def on_Combo1_keyUp_NOTSUPPORTED(self, *args):
        Globals.Log('KeyUp' + ', ' + Str(KeyCode) + ', ' + Str(Shift))

    def on_Combo1_loseFocus(self, *args):
        Globals.Log('LostFocus')

    def on_Combo1_mouseDown(self, *args):
        Button, Shift, X, Y = vbGetEventArgs(["ButtonDown()", "ShiftDown()", "x", "y"], args)
        Globals.Log('MouseDown' + Str(Button) + ', ' + Str(Shift) + ', ' + Str(X) + ', ' + Str(Y))

    def on_Combo1_mouseMove(self, *args):
        Button, Shift, X, Y = vbGetEventArgs(["ButtonDown()", "ShiftDown()", "x", "y"], args)
        Globals.Log('MouseMove' + Str(Button) + ', ' + Str(Shift) + ', ' + Str(X) + ', ' + Str(Y))

    def on_Combo1_mouseUp(self, *args):
        Button, Shift, X, Y = vbGetEventArgs(["ButtonDown()", "ShiftDown()", "x", "y"], args)
        Globals.Log('MouseUp' + Str(Button) + ', ' + Str(Shift) + ', ' + Str(X) + ', ' + Str(Y))

    def on_cmdAdd_mouseClick(self, *args):
        self.Combo1.AddItem('Item ' + Str(self.Combo1.ListCount + 1))

    def on_cmdAddFirst_mouseClick(self, *args):
        self.Combo1.AddItem('First ' + Str(self.Combo1.ListCount), 0)

    def on_cmdClear_mouseClick(self, *args):
        self.Combo1.Clear()

    def on_cmdDump_mouseClick(self, *args):
        for i in vbForRange(0, self.Combo1.ListCount - 1):
            Debug.Print(i, self.Combo1.List(i))

    def on_Delete_mouseClick(self, *args):
        self.Combo1.RemoveItem(self.Combo1.ListIndex)

    # VB2PY (UntranslatedCode) Attribute VB_Name = "frmComboBox"
    # VB2PY (UntranslatedCode) Attribute VB_GlobalNameSpace = False
    # VB2PY (UntranslatedCode) Attribute VB_Creatable = False
    # VB2PY (UntranslatedCode) Attribute VB_PredeclaredId = True
    # VB2PY (UntranslatedCode) Attribute VB_Exposed = False



if __name__ == '__main__':
    app = model.PythonCardApp(MAINFORM)
    app.MainLoop()
