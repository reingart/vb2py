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


    def on_cmdDelete_mouseClick(self, *args):
        DeleteSetting('Testing', self.txtSection.Text, self.txtName.Text)

    def on_cmdGet_mouseClick(self, *args):
        self.txtValue.Text = GetSetting('Testing', self.txtSection.Text, self.txtName.Text, '<default>')

    def on_cmdGetAll_mouseClick(self, *args):
        Setting = Variant()

        Settings = Variant()
        self.lstSettings.Clear()
        Settings = GetAllSettings('Testing', self.txtSection.Text)
        for Setting in vbForRange(0, UBound(Settings)):
            self.lstSettings.AddItem(Settings(Setting, 0) + ' = ' + Settings(Setting, 1))

    def on_cmdSet_mouseClick(self, *args):
        SaveSetting('Testing', self.txtSection.Text, self.txtName.Text, self.txtValue.Text)

    # VB2PY (UntranslatedCode) Attribute VB_Name = "frmSettings"
    # VB2PY (UntranslatedCode) Attribute VB_GlobalNameSpace = False
    # VB2PY (UntranslatedCode) Attribute VB_Creatable = False
    # VB2PY (UntranslatedCode) Attribute VB_PredeclaredId = True
    # VB2PY (UntranslatedCode) Attribute VB_Exposed = False



if __name__ == '__main__':
    app = model.PythonCardApp(MAINFORM)
    app.MainLoop()
