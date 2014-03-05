Attribute VB_Name = "Module1"
Dim MyGlobal As Integer
Private MyInt As Integer
Private MyReal As Single
Public MyVar As Variant
Dim x
Dim Suby
Dim Functioni
Dim a, b As String, c(), d As Variant
Private zz
Private Const dddd = 12
Private Const eeee = "hello there"
Public g, e(10), f(12, 10, 5) As String

Private Sub btnChange_Click()
If chkAdd.Value Then
    txtValue.Text = CStr(doAnAdd(CInt(txtValue.Text), 1))
ElseIf chkSub.Value Then
    txtValue.Text = CStr(doAnAdd(CInt(txtValue.Text), -1))
End If
End Sub

Private Sub btnDirectHide_Click()
btnDirectHide.Visible = False
End Sub

Private Sub btnDoIt_Click()
lblLabel.Caption = txtName.Text & txtSecond.Text
End Sub

Private Sub btnHideMe_Click()
'HideSomething (btnHideMe)
End Sub

Private Sub btnSecond_Click()
frmSecond.Show
End Sub

Function doAnAdd(Value, Adding)
doAnAdd = Value + Adding
End Function

Sub HideSomething(btn)
btn.Visible = False
End Sub

Private Sub btnZeroIt_Click()
txtValue.Text = "0"
End Sub

Private Sub Command1_Click()
frmRadio.Show
End Sub


Public Function Factorial(n)
If n = 0 Then
    Factorial = 1
Else
    Factorial = n * Factorial(n - 1)
End If
End Function


Public Sub test()
Dim a As String * 20
Debug.Print a, "!"
End Sub



Public Sub TestCollection()
Dim c As New Collection
For i = 1 To 10
    If i <> 5 Then c.Add "txt" & i
Next i
c.Add "txt5", before:=5
For Each i In c
    Debug.Print i
Next i
End Sub
