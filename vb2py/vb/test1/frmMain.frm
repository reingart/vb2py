VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Test form in VB"
   ClientHeight    =   6075
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFactorial 
      Caption         =   "Factorial"
      Height          =   375
      Left            =   8520
      TabIndex        =   22
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show radio"
      Height          =   375
      Left            =   6960
      TabIndex        =   21
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton btnDirectHide 
      Caption         =   "Hide me direct"
      Height          =   375
      Left            =   3240
      TabIndex        =   20
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton btnZeroIt 
      Caption         =   "Zero it!"
      Height          =   375
      Left            =   480
      TabIndex        =   19
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton btnHideMe 
      Caption         =   "Hide me (late bound)"
      Height          =   375
      Left            =   4320
      TabIndex        =   18
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CheckBox chkSub 
      Caption         =   "Sub"
      Height          =   255
      Left            =   1560
      TabIndex        =   17
      Top             =   4080
      Width           =   855
   End
   Begin VB.CheckBox chkAdd 
      Caption         =   "Add"
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   3720
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.TextBox txtValue 
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Text            =   "0"
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton btnChange 
      Caption         =   "Change"
      Height          =   375
      Left            =   480
      TabIndex        =   14
      Top             =   3720
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Top             =   3720
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   3360
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   6480
      TabIndex        =   10
      Top             =   2520
      Width           =   3495
   End
   Begin VB.ListBox List1 
      Height          =   1425
      ItemData        =   "frmMain.frx":0000
      Left            =   6480
      List            =   "frmMain.frx":0019
      TabIndex        =   9
      Top             =   720
      Width           =   3495
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   360
      TabIndex        =   8
      Text            =   "Combo3"
      Top             =   3000
      Width           =   2655
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmMain.frx":0047
      Left            =   3600
      List            =   "frmMain.frx":004E
      TabIndex        =   7
      Text            =   "Combo2"
      Top             =   2640
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmMain.frx":0059
      Left            =   3600
      List            =   "frmMain.frx":0072
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   2160
      Width           =   2535
   End
   Begin VB.ComboBox cmbChoose 
      Height          =   315
      ItemData        =   "frmMain.frx":00B4
      Left            =   3600
      List            =   "frmMain.frx":00C7
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtSecond 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton btnSecond 
      Caption         =   "Second Form"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton btnDoIt 
      Caption         =   "Do it"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "More static text"
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblLabel 
      Caption         =   "My text"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const a As Integer = 10, b As String = "hello", c As String * 10 = 43


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

Private Sub cmdFactorial_Click()
MsgBox "Factorial 6 is " & Factorial(6)
End Sub

Private Sub Command1_Click()
frmRadio.Show
End Sub

Private Sub Form_Load()
Close
'b (1,2), f(3)
Dim a As Integer
Select Case a
'a = 10
End Select
bb = #1/10/2000#
End Sub
