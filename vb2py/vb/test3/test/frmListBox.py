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


    def on_cmdAdd_mouseClick(self, *args):
        self.List1.AddItem('Item ' + Str(self.List1.ListCount + 1))

    def on_cmdAddFirst_mouseClick(self, *args):
        self.List1.AddItem('First ' + Str(self.List1.ListCount), 0)

    def on_cmdClear_mouseClick(self, *args):
        self.List1.Clear()

    def on_cmdDump_mouseClick(self, *args):
        for i in vbForRange(0, self.List1.ListCount - 1):
            Debug.Print(i, self.List1.List(i))

    def on_Command4_mouseClick(self, *args):
        self.List2.Visible = not self.List2.Visible

    def on_Command5_mouseClick(self, *args):
        self.List2.Left = self.List2.Left + 20
        self.List2.Top = self.List2.Top + 20

    def on_Command6_mouseClick(self, *args):
        self.List2.Width = self.List2.Width + 20
        self.List2.Height = self.List2.Height + 50

    def on_Command7_mouseClick(self, *args):
        self.List2.Enabled = not self.List2.Enabled

    def on_Delete_mouseClick(self, *args):
        self.List1.RemoveItem(self.List1.ListIndex)

    def on_List1_mouseClick(self, *args):
        Globals.Log('Click: ' + Str(self.List1.ListIndex))

    def on_List1_mouseDoubleClick(self, *args):
        Globals.Log('DblClick: ' + Str(self.List1.ListIndex))

    def on_List1_gainFocus(self, *args):
        Globals.Log('Got Focus')

    def __List1_ItemCheck(self, Item):
        Globals.Log('Item check' + Str(Item))

    def on_List1_loseFocus(self, *args):
        Globals.Log('Lost Focus')

    def on_List1_mouseDown(self, *args):
        Button, Shift, X, Y = vbGetEventArgs(["ButtonDown()", "ShiftDown()", "x", "y"], args)
        Globals.Log('MouseDown' + Str(Button) + ', ' + Str(Shift) + ', ' + Str(X) + ', ' + Str(Y))

    def on_List1_mouseMove(self, *args):
        Button, Shift, X, Y = vbGetEventArgs(["ButtonDown()", "ShiftDown()", "x", "y"], args)
        Globals.Log('MouseMove' + Str(Button) + ', ' + Str(Shift) + ', ' + Str(X) + ', ' + Str(Y))

    def on_List1_mouseUp(self, *args):
        Button, Shift, X, Y = vbGetEventArgs(["ButtonDown()", "ShiftDown()", "x", "y"], args)
        Globals.Log('MouseUp' + Str(Button) + ', ' + Str(Shift) + ', ' + Str(X) + ', ' + Str(Y))

    def __List1_Scroll(self):
        Globals.Log('Scroll')

    # VB2PY (UntranslatedCode) Attribute VB_Name = "frmListBox"
    # VB2PY (UntranslatedCode) Attribute VB_GlobalNameSpace = False
    # VB2PY (UntranslatedCode) Attribute VB_Creatable = False
    # VB2PY (UntranslatedCode) Attribute VB_PredeclaredId = True
    # VB2PY (UntranslatedCode) Attribute VB_Exposed = False



if __name__ == '__main__':
    app = model.PythonCardApp(MAINFORM)
    app.MainLoop()
