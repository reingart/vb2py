"""The main form for the application"""

from PythonCard import model

# Allow importing of our custom controls
import PythonCard.resource
PythonCard.resource.APP_COMPONENTS_PACKAGE = "vb2py.targets.pythoncard.vbcontrols"

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
import Module1

class MAINFORM(Background):

    __a = 10
    __b = 'hello'
    __c = 43

    def on_btnChange_mouseClick(self, *args):
        if self.chkAdd.Value:
            self.txtValue.Text = CStr(self.__doAnAdd(CInt(self.txtValue.Text), 1))
        elif self.chkSub.Value:
            self.txtValue.Text = CStr(self.__doAnAdd(CInt(self.txtValue.Text), -1))

    def on_btnDirectHide_mouseClick(self, *args):
        self.btnDirectHide.Visible = False

    def on_btnDoIt_mouseClick(self, *args):
        self.lblLabel.Caption = self.txtName.Text + self.txtSecond.Text

    def on_btnHideMe_mouseClick(self, *args):
        #HideSomething (btnHideMe)
        pass

    def on_btnSecond_mouseClick(self, *args):
        frmSecond.Show()

    def __doAnAdd(self, Value, Adding):
        _ret = None
        _ret = Value + Adding
        return _ret

    def __HideSomething(self, btn):
        btn.Visible = False

    def on_btnZeroIt_mouseClick(self, *args):
        self.txtValue.Text = '0'

    def on_cmdFactorial_mouseClick(self, *args):
        MsgBox('Factorial 6 is ' + Module1.Factorial(6))

    def on_Command1_mouseClick(self, *args):
        frmRadio.Show()

    def __Form_Load(self):
        a = Integer()
        VBFiles.closeFile()
        #b (1,2), f(3)
        _select0 = a
#a = 10
        bb = MakeDate(1, 10, 2000)

    # VB2PY (UntranslatedCode) Attribute VB_Name = "frmMain"
    # VB2PY (UntranslatedCode) Attribute VB_GlobalNameSpace = False
    # VB2PY (UntranslatedCode) Attribute VB_Creatable = False
    # VB2PY (UntranslatedCode) Attribute VB_PredeclaredId = True
    # VB2PY (UntranslatedCode) Attribute VB_Exposed = False



if __name__ == '__main__':
    app = model.Application(MAINFORM)
    app.MainLoop()
