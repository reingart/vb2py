VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "vb2py Test Form"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Windows API"
      Height          =   735
      Left            =   120
      TabIndex        =   27
      Top             =   7920
      Width           =   9255
      Begin VB.CommandButton cmdGetUserName 
         Caption         =   "Get User Name"
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblUserName 
         Caption         =   "Press the button to check"
         Height          =   255
         Left            =   2280
         TabIndex        =   29
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Create Object"
      Height          =   2535
      Left            =   120
      TabIndex        =   18
      Top             =   5280
      Width           =   9255
      Begin VB.CommandButton cmdSetValue 
         Caption         =   "Set Value"
         Height          =   375
         Left            =   1800
         TabIndex        =   25
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtNewValue 
         Height          =   375
         Left            =   720
         TabIndex        =   24
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton cmdGetValue 
         Caption         =   "Get Value"
         Height          =   375
         Left            =   1800
         TabIndex        =   22
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtCell 
         Height          =   375
         Left            =   720
         TabIndex        =   21
         Text            =   "A1"
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton cmdAttach 
         Caption         =   "Attach to Excel"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblAttached 
         Caption         =   "Not attached"
         Height          =   255
         Left            =   2400
         TabIndex        =   26
         Top             =   360
         Width           =   6495
      End
      Begin VB.Label lblCellValue 
         Caption         =   "="
         Height          =   375
         Left            =   3360
         TabIndex        =   23
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Cell"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Numeric"
      Height          =   1215
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Width           =   9255
      Begin VB.ComboBox cmbFunction 
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   3480
         List            =   "frmMain.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox txtValue 
         Height          =   375
         Left            =   960
         TabIndex        =   14
         Text            =   "1"
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblResult 
         Caption         =   "Answer"
         Height          =   255
         Left            =   6600
         TabIndex        =   17
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Number"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdDoIt 
      Caption         =   "Process Text"
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Text"
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   7920
      TabIndex        =   9
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "String Processing"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.TextBox txtReverse 
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1200
         Width           =   6615
      End
      Begin VB.TextBox txtLower 
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   720
         Width           =   6615
      End
      Begin VB.ListBox lstSplit 
         Height          =   1230
         Left            =   2040
         TabIndex        =   5
         Top             =   1680
         Width           =   6615
      End
      Begin VB.TextBox txtRawString 
         Height          =   405
         Left            =   2040
         TabIndex        =   2
         Text            =   "This TEXT contains a few words"
         Top             =   240
         Width           =   6615
      End
      Begin VB.Label Label4 
         Caption         =   "Split into words"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Reversed"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Lower case"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Enter your string here"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private MyExcel As Object

Private Sub cmbFunction_Change()
txtValue_Change
End Sub

Private Sub cmdAttach_Click()
Set MyExcel = CreateObject("Excel.Application")
MyExcel.Workbooks.Add
lblAttached.Caption = MyExcel.Name
End Sub

Private Sub cmdClear_Click()
txtRawString.Text = ""
End Sub

Private Sub cmdClose_Click()
End
End Sub

Private Sub cmdDefault_Click()
cmdClear_Click
txtRawString.Text = "The cat sat on the MAT"
cmdDoIt_Click
End Sub

Private Sub cmdDoIt_Click()
'
' Do the lower case
txtLower.Text = LCase(txtRawString.Text)
'
' Do the reversing
txtReverse.Text = ""
For i = Len(txtRawString.Text) To 1 Step -1
    txtReverse.Text = txtReverse.Text & Mid(txtRawString.Text, i, 1)
Next i
'
' Do the splitting
lstSplit.Clear
For Each Word In SplitOnWord(txtRawString.Text, " ")
    lstSplit.AddItem Word
Next Word
'
End Sub



Private Sub cmdGetUserName_Click()
lblUserName.Caption = GetUserName
End Sub

Private Sub cmdGetValue_Click()
lblCellValue.Caption = MyExcel.Workbooks(1).Sheets(1).Range(txtCell.Text).Value
End Sub

Private Sub cmdSetValue_Click()
MyExcel.Sheets(1).Range(txtCell.Text).Value = txtNewValue.Text
End Sub

Private Sub txtValue_Change()
If IsNumeric(txtValue.Text) Then doCalcs
End Sub

Private Sub doCalcs()
Select Case cmbFunction.Text
    Case "Sin"
        lblResult.Caption = CStr(Sin(txtValue.Text))
    Case "Cos"
        lblResult.Caption = CStr(Cos(txtValue.Text))
    Case "Sqrt"
        lblResult.Caption = CStr(Sqr(txtValue.Text))
    Case "Factorial"
        lblResult.Caption = CStr(Factorial(CSng(txtValue.Text)))
End Select
End Sub
