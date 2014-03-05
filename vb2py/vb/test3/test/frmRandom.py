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

class MAINFORM(Background):


    def on_cmdGetRandom_mouseClick(self, *args):
        for i in vbForRange(1, 10):
            self.txtResults.Text = self.txtResults.Text + CStr(i) + ' : ' + CStr(Rnd()) + vbCrLf

    def on_cmdNegCall_mouseClick(self, *args):
        a = Rnd(- 1)

    def on_cmdRandomize_mouseClick(self, *args):
        Randomize()

    def on_cmdRandomizeWithSeed_mouseClick(self, *args):
        Randomize(Int(self.txtSeed.Text))

    # VB2PY (UntranslatedCode) Attribute VB_Name = "frmRandom"
    # VB2PY (UntranslatedCode) Attribute VB_GlobalNameSpace = False
    # VB2PY (UntranslatedCode) Attribute VB_Creatable = False
    # VB2PY (UntranslatedCode) Attribute VB_PredeclaredId = True
    # VB2PY (UntranslatedCode) Attribute VB_Exposed = False



if __name__ == '__main__':
    app = model.PythonCardApp(MAINFORM)
    app.MainLoop()
