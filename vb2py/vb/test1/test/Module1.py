from vb2py.vbfunctions import *
from vb2py.vbdebug import *


MyGlobal = Integer()
MyInt = Integer()
MyReal = Single()
MyVar = Variant()
x = Variant()
Suby = Variant()
Functioni = Variant()
a = Variant()
b = String()
c = vbObjectInitialize(objtype=Variant)
d = Variant()
zz = Variant()
dddd = 12
eeee = 'hello there'
g = Variant()
e = vbObjectInitialize((10,), Variant)
f = vbObjectInitialize((12, 10, 5,), String)

def btnChange_Click():
    if chkAdd.Value:
        txtValue.Text = CStr(doAnAdd(CInt(txtValue.Text), 1))
    elif chkSub.Value:
        txtValue.Text = CStr(doAnAdd(CInt(txtValue.Text), -1))

def btnDirectHide_Click():
    btnDirectHide.Visible = False

def btnDoIt_Click():
    lblLabel.Caption = txtName.Text + txtSecond.Text

def btnHideMe_Click():
    #HideSomething (btnHideMe)
    pass

def btnSecond_Click():
    frmSecond.Show()

def doAnAdd(Value, Adding):
    _ret = None
    _ret = Value + Adding
    return _ret

def HideSomething(btn):
    btn.Visible = False

def btnZeroIt_Click():
    txtValue.Text = '0'

def Command1_Click():
    frmRadio.Show()

def Factorial(n):
    _ret = None
    if n == 0:
        _ret = 1
    else:
        _ret = n * Factorial(n - 1)
    return _ret

def test():
    a = FixedString(20)
    Debug.Print(a, '!')

def TestCollection():
    c = Collection()
    for i in vbForRange(1, 10):
        if i <> 5:
            c.Add('txt' + i)
    c.Add('txt5', before=5)
    for i in c:
        Debug.Print(i)

# VB2PY (UntranslatedCode) Attribute VB_Name = "Module1"
