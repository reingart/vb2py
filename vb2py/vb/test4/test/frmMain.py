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
import vb2py.custom.comctllib

class MAINFORM(Background):
    """ Make sure we import the common controls for Python"""


    def on_chkEnableAutoEdit_mouseClick(self, *args):
        if self.chkEnableAutoEdit.Value:
            self.tvTree.LabelEdit = tvwAutomatic
        else:
            self.tvTree.LabelEdit = vb2py.custom.comctllib.tvwManual

    def on_cmdAdd_mouseClick(self, *args):
        if self.tvTree.SelectedItem is Nothing:
            self.tvTree.Nodes.Add(VBGetMissingArgument(self.tvTree.Nodes.Add, 0), VBGetMissingArgument(self.tvTree.Nodes.Add, 1), self.txtName.Text, self.txtName.Text)
        else:
            self.tvTree.Nodes.Add(self.tvTree.SelectedItem.Key, vb2py.custom.comctllib.tvwChild, self.txtName.Text, self.txtName.Text)

    def on_cmdAddTree_mouseClick(self, *args):
        self.tvTree.Nodes.Clear()
        self.__setTree(self.txtTree.Text)

    def on_cmdClear_mouseClick(self, *args):
        self.tvTree.Nodes.Clear()

    def on_cmdEnable_mouseClick(self, *args):
        self.tvTree.Enabled = not self.tvTree.Enabled

    def on_cmdExpand_mouseClick(self, *args):
        for Node in self.tvTree.Nodes:
            vb2py.custom.comctllib.Node.Expanded = True

    def on_cmdCollapse_mouseClick(self, *args):
        for Node in self.tvTree.Nodes:
            vb2py.custom.comctllib.Node.Expanded = False

    def on_cmdLoadPicture_mouseClick(self, *args):
        self.ilDynamic.ListImages.Add(VBGetMissingArgument(self.ilDynamic.ListImages.Add, 0), 'closed', LoadPicture(App.Path + '\\closedicon.ico'))
        self.ilDynamic.ListImages.Add(VBGetMissingArgument(self.ilDynamic.ListImages.Add, 0), 'open', LoadPicture(App.Path + '\\openicon.ico'))

    def on_cmdMove_mouseClick(self, *args):
        self.tvTree.Left = self.tvTree.Left + 10
        self.tvTree.Top = self.tvTree.Top + 10

    def on_cmdRemove_mouseClick(self, *args):
        if self.tvTree.SelectedItem is Nothing:
            MsgBox('No selection')
        else:
            self.tvTree.Nodes.Remove(self.tvTree.SelectedItem.Key)

    def on_cmdSetAsDynamic_mouseClick(self, *args):
        self.tvTree.ImageList = self.ilDynamic

    def on_cmdSetAsPreload_mouseClick(self, *args):
        self.tvTree.ImageList = self.imPreload

    def on_cmdSetPictures_mouseClick(self, *args):
        Nde = vb2py.custom.comctllib.Node()
        for Nde in self.tvTree.Nodes:
            Nde.Image = 'closed'
            Nde.ExpandedImage = 'open'

    def on_cmdSize_mouseClick(self, *args):
        self.tvTree.Width = self.tvTree.Width + 10
        self.tvTree.Height = self.tvTree.Height + 10

    def on_cmdTestNode_mouseClick(self, *args):
        This = vb2py.custom.comctllib.Node()
        #
        This = self.tvTree.Nodes(self.txtNodeName.Text)
        This.Selected = True
        self.txtResults.Text = 'text:' + This.Text + vbCrLf + 'tag:' + This.Tag + vbCrLf
        self.txtResults.Text = self.txtResults.Text + 'visible:' + This.Visible + vbCrLf + 'children:' + This.Children + vbCrLf
        if This.Children > 0:
            self.txtResults.Text = self.txtResults.Text + 'childtext:' + This.Child.Text + vbCrLf
        self.txtResults.Text = self.txtResults.Text + 'firstsib:' + This.FirstSibling.Text + vbCrLf + 'lastsib:' + This.LastSibling.Text + vbCrLf
        self.txtResults.Text = self.txtResults.Text + 'path:' + This.FullPath + vbCrLf + 'next:' + This.Next.Text + vbCrLf
        self.txtResults.Text = self.txtResults.Text + 'parent:' + This.Parent.Text + vbCrLf + 'previous:' + This.Previous.Text + vbCrLf
        self.txtResults.Text = self.txtResults.Text + 'root:' + This.Root.Text + vbCrLf
        This.EnsureVisible()
        This.Selected = True
        #

    def on_cmdVisible_mouseClick(self, *args):
        self.tvTree.Visible = not self.tvTree.Visible

    def __Form_Load(self):
        self.txtTree.Text = 'A=ROOT' + vbCrLf + 'A1=A' + vbCrLf + 'A2=A' + vbCrLf + 'A3=A' + vbCrLf + 'A3A=A3' + vbCrLf + 'A4=A' + vbCrLf + 'B=ROOT' + vbCrLf + 'B1=B'

    def __setTree(self, Text):
        Name = String()

        Last = vb2py.custom.comctllib.Node()

        Remainder = String()
        # Set the tree up
        # Get name
        #
        while Text <> '':
            #
            posn = InStr(Text, vbCrLf)
            if posn <> 0:
                parts = self.__strSplit(Text, vbCrLf)
                Name = parts(0)
                Text = parts(1)
            else:
                Name = Text
                Text = ''
            #
            parts = self.__strSplit(Name, '=')
            nodename = parts(0)
            parentname = parts(1)
            #
            if parentname == 'ROOT':
                self.tvTree.Nodes.Add(VBGetMissingArgument(self.tvTree.Nodes.Add, 0), VBGetMissingArgument(self.tvTree.Nodes.Add, 1), nodename, nodename)
            else:
                self.tvTree.Nodes.Add(parentname, vb2py.custom.comctllib.tvwChild, nodename, nodename)
            #
        #

    def __strSplit(self, Text, Delim):
        _ret = None
        parts = vbObjectInitialize((1,), Variant)
        posn = InStr(Text, Delim)
        parts[0] = Left(Text, posn - 1)
        parts[1] = Right(Text, Len(Text) - posn - Len(Delim) + 1)
        _ret = parts
        return _ret

    def __tvTree_AfterLabelEdit(self, Cancel, NewString):
        Debug.Print('After label edit on ' + self.tvTree.SelectedItem.Text + ' new name is ' + NewString)
        if NewString == 'CCC':
            Debug.Print('Cancelled')
            Cancel = 1
        else:
            Debug.Print('OK')
            Cancel = 0

    def __tvTree_BeforeLabelEdit(self, Cancel):
        Debug.Print('Before label edit on ' + self.tvTree.SelectedItem.Text)
        if self.chkAllowEdits.Value:
            Cancel = 0
        else:
            Cancel = 1

    def on_tvTree_mouseClick(self, *args):
        Debug.Print('Tree view click')

    def __tvTree_Collapse(self, Node):
        Debug.Print('Tree view collapse on ' + Node.Text)

    def __tvTree_DblClick(self):
        Debug.Print('Tree view double click')

    def on_tvTree_DragDrop_NOTSUPPORTED(self, *args):
        Debug.Print('Tree view drag drop')

    def on_tvTree_DragOver_NOTSUPPORTED(self, *args):
        Debug.Print('Tree view drag over')

    def __tvTree_Expand(self, Node):
        Debug.Print('Tree expand on ' + Node.Text)

    def on_tvTree_gainFocus(self, *args):
        Debug.Print('Tree view got focus')

    def on_tvTree_keyDown_NOTSUPPORTED(self, *args):
        Debug.Print('Tree view keydown (code, shift) ' + CStr(KeyCode) + ', ' + CStr(Shift))

    def on_tvTree_keyPress_NOTSUPPORTED(self, *args):
        Debug.Print('Tree view keypress (code) ' + CStr(KeyAscii))

    def on_tvTree_keyUp_NOTSUPPORTED(self, *args):
        Debug.Print('Tree view keyup (code, shift) ' + CStr(KeyCode) + ', ' + CStr(Shift))

    def on_tvTree_loseFocus(self, *args):
        Debug.Print('Tree view lost focus')

    def on_tvTree_mouseDown(self, *args):
        Button, Shift, x, y = vbGetEventArgs(["ButtonDown()", "ShiftDown()", "x", "y"], args)
        Debug.Print('Tree view mouse down (button, shift, x, y) ' + CStr(Button) + ', ' + CStr(Shift) + ', ' + CStr(x) + ', ' + CStr(y))

    def on_tvTree_mouseMove(self, *args):
        Button, Shift, x, y = vbGetEventArgs(["ButtonDown()", "ShiftDown()", "x", "y"], args)
        Debug.Print('Tree view mouse move (button, shift, x, y) ' + CStr(Button) + ', ' + CStr(Shift) + ', ' + CStr(x) + ', ' + CStr(y))

    def on_tvTree_mouseUp(self, *args):
        Button, Shift, x, y = vbGetEventArgs(["ButtonDown()", "ShiftDown()", "x", "y"], args)
        Debug.Print('Tree view mouse up (button, shift, x, y) ' + CStr(Button) + ', ' + CStr(Shift) + ', ' + CStr(x) + ', ' + CStr(y))

    def __tvTree_NodeClick(self, Node):
        Debug.Print('Tree node click ' + Node.Text)

    # VB2PY (UntranslatedCode) Attribute VB_Name = "frmMain"
    # VB2PY (UntranslatedCode) Attribute VB_GlobalNameSpace = False
    # VB2PY (UntranslatedCode) Attribute VB_Creatable = False
    # VB2PY (UntranslatedCode) Attribute VB_PredeclaredId = True
    # VB2PY (UntranslatedCode) Attribute VB_Exposed = False



if __name__ == '__main__':
    app = model.PythonCardApp(MAINFORM)
    app.MainLoop()
