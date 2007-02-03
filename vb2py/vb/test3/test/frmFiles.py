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


    def on_cmdCheck_mouseClick(self, *args):
        chn = Integer()
        chn = FreeFile()
        VBFiles.openFile(chn, 'test.txt', 'r') 
        Globals.test = VBFiles.getInput(chn, 1)
        if Globals.test == 'Ok!':
            MsgBox('It worked')
        else:
            MsgBox('It didn\'t work')
        VBFiles.closeFile(chn)

    def on_cmdDelete_mouseClick(self, *args):
        Kill('test.txt')

    def on_cmdListFiles_mouseClick(self, *args):
        Name = String()
        self.lstFiles.Clear()
        Name = Dir(self.txtDir.Text + '\\*')
        while Name <> '':
            self.lstFiles.AddItem(Name)
            Name = Dir()

    def on_cmdMakeDir_mouseClick(self, *args):
        MkDir(self.txtDir.Text)

    def on_cmdMakeTestFile_mouseClick(self, *args):
        chn = Integer()
        chn = FreeFile()
        VBFiles.openFile(chn, 'test.txt', 'w') 
        VBFiles.writeText(chn, 'Ok!', '\n')
        VBFiles.closeFile(chn)

    def on_cmdSet_mouseClick(self, *args):
        ChDir(self.txtDir.Text)

    def on_Command1_mouseClick(self, *args):
        RmDir(self.txtDir.Text)

    # VB2PY (UntranslatedCode) Attribute VB_Name = "frmFiles"
    # VB2PY (UntranslatedCode) Attribute VB_GlobalNameSpace = False
    # VB2PY (UntranslatedCode) Attribute VB_Creatable = False
    # VB2PY (UntranslatedCode) Attribute VB_PredeclaredId = True
    # VB2PY (UntranslatedCode) Attribute VB_Exposed = False



if __name__ == '__main__':
    app = model.PythonCardApp(MAINFORM)
    app.MainLoop()
