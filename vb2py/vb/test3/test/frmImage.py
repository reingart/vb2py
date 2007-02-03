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

    __LastPicture = String()
    __RootDir = String()

    def on_Command1_mouseClick(self, *args):
        self.Image1.Stretch = not self.Image1.Stretch
        self.Label1.Caption = 'Strech = ' + Str(self.Image1.Stretch)

    def on_Command2_mouseClick(self, *args):
        if self.__LastPicture == 'vb2py.gif':
            self.Image1.Picture = LoadPicture(self.__RootDir + '/vb2pylogosm.jpg')
            self.__LastPicture = 'vb2pylogosm.jpg'
        else:
            self.Image1.Picture = LoadPicture(self.__RootDir + '/vb2py.gif')
            self.__LastPicture = 'vb2py.gif'
        self.Label2.Caption = 'Showing ' + self.__LastPicture

    def Form_Load(self):
        self.__RootDir = App.Path

    def on_Image1_mouseDown(self, *args):
        Button, Shift, X, Y = vbGetEventArgs(["ButtonDown()", "ShiftDown()", "x", "y"], args)
        Globals.Log('MouseDown' + Str(Button) + ', ' + Str(Shift) + ', ' + Str(X) + ', ' + Str(Y))

    def on_Image1_mouseMove(self, *args):
        Button, Shift, X, Y = vbGetEventArgs(["ButtonDown()", "ShiftDown()", "x", "y"], args)
        Globals.Log('MouseMove' + Str(Button) + ', ' + Str(Shift) + ', ' + Str(X) + ', ' + Str(Y))

    def on_Image1_mouseUp(self, *args):
        Button, Shift, X, Y = vbGetEventArgs(["ButtonDown()", "ShiftDown()", "x", "y"], args)
        Globals.Log('MouseUp' + Str(Button) + ', ' + Str(Shift) + ', ' + Str(X) + ', ' + Str(Y))

    def on_Command4_mouseClick(self, *args):
        self.Image1.Visible = not self.Image1.Visible

    def on_Command5_mouseClick(self, *args):
        self.Image1.Left = self.Image1.Left + 20
        self.Image1.Top = self.Image1.Top + 20

    def on_Command6_mouseClick(self, *args):
        self.Image1.Width = self.Image1.Width + 20
        self.Image1.Height = self.Image1.Height + 50

    def on_Command7_mouseClick(self, *args):
        self.Image1.Enabled = not self.Image1.Enabled

    # VB2PY (UntranslatedCode) Attribute VB_Name = "frmImage"
    # VB2PY (UntranslatedCode) Attribute VB_GlobalNameSpace = False
    # VB2PY (UntranslatedCode) Attribute VB_Creatable = False
    # VB2PY (UntranslatedCode) Attribute VB_PredeclaredId = True
    # VB2PY (UntranslatedCode) Attribute VB_Exposed = False



if __name__ == '__main__':
    app = model.PythonCardApp(MAINFORM)
    app.MainLoop()
