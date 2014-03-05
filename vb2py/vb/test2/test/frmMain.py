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
import Utils

class MAINFORM(Background):

    __MyExcel = Object()

    def on_cmbFunction_textUpdate(self, *args):
        self.on_txtValue_textUpdate()

    def on_cmdAttach_mouseClick(self, *args):
        self.__MyExcel = CreateObject('Excel.Application')
        self.__MyExcel.Workbooks.Add()
        self.lblAttached.Caption = self.__MyExcel.Name

    def on_cmdClear_mouseClick(self, *args):
        self.txtRawString.Text = ''

    def on_cmdClose_mouseClick(self, *args):
        sys.exit(0)

    def on_cmdDefault_mouseClick(self, *args):
        self.on_cmdClear_mouseClick()
        self.txtRawString.Text = 'The cat sat on the MAT'
        self.on_cmdDoIt_mouseClick()

    def on_cmdDoIt_mouseClick(self, *args):
        #
        # Do the lower case
        self.txtLower.Text = LCase(self.txtRawString.Text)
        #
        # Do the reversing
        self.txtReverse.Text = ''
        for i in vbForRange(Len(self.txtRawString.Text), 1, -1):
            self.txtReverse.Text = self.txtReverse.Text + Mid(self.txtRawString.Text, i, 1)
        #
        # Do the splitting
        self.lstSplit.Clear()
        for Word in Utils.SplitOnWord(self.txtRawString.Text, ' '):
            self.lstSplit.AddItem(Word)
        #

    def on_cmdGetUserName_mouseClick(self, *args):
        self.lblUserName.Caption = GetUserName

    def on_cmdGetValue_mouseClick(self, *args):
        self.lblCellValue.Caption = self.__MyExcel.Workbooks(1).Sheets(1).Range(self.txtCell.Text).Value

    def on_cmdSetValue_mouseClick(self, *args):
        self.__MyExcel.Sheets[1].Range[self.txtCell.Text].Value = self.txtNewValue.Text

    def on_txtValue_textUpdate(self, *args):
        if IsNumeric(self.txtValue.Text):
            self.__doCalcs()

    def __doCalcs(self):
        _select0 = self.cmbFunction.Text
        if (_select0 == 'Sin'):
            self.lblResult.Caption = CStr(Sin(self.txtValue.Text))
        elif (_select0 == 'Cos'):
            self.lblResult.Caption = CStr(Cos(self.txtValue.Text))
        elif (_select0 == 'Sqrt'):
            self.lblResult.Caption = CStr(Sqr(self.txtValue.Text))
        elif (_select0 == 'Factorial'):
            self.lblResult.Caption = CStr(Utils.Factorial(CSng(self.txtValue.Text)))

    # VB2PY (UntranslatedCode) Attribute VB_Name = "frmMain"
    # VB2PY (UntranslatedCode) Attribute VB_GlobalNameSpace = False
    # VB2PY (UntranslatedCode) Attribute VB_Creatable = False
    # VB2PY (UntranslatedCode) Attribute VB_PredeclaredId = True
    # VB2PY (UntranslatedCode) Attribute VB_Exposed = False



if __name__ == '__main__':
    app = model.Application(MAINFORM)
    app.MainLoop()
