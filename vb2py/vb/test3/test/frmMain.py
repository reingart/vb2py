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
        frmButton.Show()

    def on_Command10_mouseClick(self, *args):
        frmSettings.Show()

    def on_Command11_mouseClick(self, *args):
        Globals.EraseTest()

    def on_Command2_mouseClick(self, *args):
        frmTextBox.Show()

    def on_Command3_mouseClick(self, *args):
        frmComboBox.Show()

    def on_Command4_mouseClick(self, *args):
        frmListBox.Show()

    def on_Command5_mouseClick(self, *args):
        frmCheckBox.Show()

    def on_Command6_mouseClick(self, *args):
        frmImage.Show()

    def on_Command7_mouseClick(self, *args):
        Globals.test()

    def on_Command8_mouseClick(self, *args):
        frmFiles.Show()

    def on_Command9_mouseClick(self, *args):
        frmRandom.Show()

    # VB2PY (UntranslatedCode) Attribute VB_Name = "frmMain"
    # VB2PY (UntranslatedCode) Attribute VB_GlobalNameSpace = False
    # VB2PY (UntranslatedCode) Attribute VB_Creatable = False
    # VB2PY (UntranslatedCode) Attribute VB_PredeclaredId = True
    # VB2PY (UntranslatedCode) Attribute VB_Exposed = False



if __name__ == '__main__':
    app = model.PythonCardApp(MAINFORM)
    app.MainLoop()
