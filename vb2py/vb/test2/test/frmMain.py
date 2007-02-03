# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

"""The main form for the application"""

from PythonCardPrototype import model

class Background(model.Background):

	def __getattr__(self, name):
		"""If a name was not found then look for it in components"""
		return getattr(self.components, name)


from vb2py.vbfunctions import *
import Utils

class MAINFORM(Background):

    __MyExcel = Object()

    def on_cmbFunction_textUpdate(self, event=None):
        self.on_txtValue_textUpdate()

    def on_cmdAttach_mouseClick(self, event=None):
        self.__MyExcel = CreateObject('Excel.Application')
        self.__MyExcel.Workbooks.Add()
        self.lblAttached.text = self.__MyExcel.Name

    def on_cmdClear_mouseClick(self, event=None):
        self.txtRawString.text = ''

    def on_cmdClose_mouseClick(self, event=None):
        sys.exit(0)

    def on_cmdDefault_mouseClick(self, event=None):
        self.on_cmdClear_mouseClick()
        self.txtRawString.text = 'The cat sat on the MAT'
        self.on_cmdDoIt_mouseClick()

    def on_cmdDoIt_mouseClick(self, event=None):
        #
        # Do the lower case
        self.txtLower.text = LCase(self.txtRawString.text)
        #
        # Do the reversing
        self.txtReverse.text = ''
        for i in vbForRange(Len(self.txtRawString.text), 1, -1):
            self.txtReverse.text = self.txtReverse.text + Mid(self.txtRawString.text, i, 1)
        #
        # Do the splitting
        self.lstSplit.Clear()
        for Word in Utils.SplitOnWord(self.txtRawString.text, ' '):
            self.lstSplit.append(Word)
        #

    def on_cmdGetUserName_mouseClick(self, event=None):
        self.lblUserName.text = GetUserName

    def on_cmdGetValue_mouseClick(self, event=None):
        self.lblCellValue.text = self.__MyExcel.Workbooks(1).Sheets(1).Range(self.txtCell.text).Value

    def on_cmdSetValue_mouseClick(self, event=None):
        self.__MyExcel.Sheets[1].Range[self.txtCell.text].Value = self.txtNewValue.text

    def on_txtValue_textUpdate(self, event=None):
        if IsNumeric(self.txtValue.text):
            self.__doCalcs()

    def __doCalcs(self):
        _select0 = self.cmbFunction.text
        if (_select0 == 'Sin'):
            self.lblResult.text = CStr(Sin(self.txtValue.text))
        elif (_select0 == 'Cos'):
            self.lblResult.text = CStr(Cos(self.txtValue.text))
        elif (_select0 == 'Sqrt'):
            self.lblResult.text = CStr(Sqr(self.txtValue.text))
        elif (_select0 == 'Factorial'):
            self.lblResult.text = CStr(Utils.Factorial(CSng(self.txtValue.text)))

    # VB2PY (UntranslatedCode) Attribute VB_Name = "frmMain"
    # VB2PY (UntranslatedCode) Attribute VB_GlobalNameSpace = False
    # VB2PY (UntranslatedCode) Attribute VB_Creatable = False
    # VB2PY (UntranslatedCode) Attribute VB_PredeclaredId = True
    # VB2PY (UntranslatedCode) Attribute VB_Exposed = False



if __name__ == '__main__':
	app = model.PythonCardApp(MAINFORM)
	app.MainLoop()
