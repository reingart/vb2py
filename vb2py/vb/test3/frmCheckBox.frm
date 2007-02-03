VERSION 5.00
Begin VB.Form frmCheckBox 
   Caption         =   "Check Boxes"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Change Text"
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Enable"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Size "
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Move "
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Make Visible"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CheckBox Check6 
      Alignment       =   1  'Right Justify
      Caption         =   "Three"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CheckBox Check5 
      Alignment       =   1  'Right Justify
      Caption         =   "Two"
      Height          =   615
      Left            =   3480
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.CheckBox Check4 
      Alignment       =   1  'Right Justify
      Caption         =   "One"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Three"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Two"
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "One"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Log "Click " & Str(Check1.Value)
End Sub

Private Sub Check1_GotFocus()
Log "Got focus"
End Sub


Private Sub Check1_LostFocus()
Log "Lost focus"
End Sub

Private Sub Check1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Log "MouseDown" & Str(Button) & ", " & Str(Shift) & ", " & Str(X) & ", " & Str(Y)
End Sub

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Log "MouseMove" & Str(Button) & ", " & Str(Shift) & ", " & Str(X) & ", " & Str(Y)
End Sub

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Log "MouseUp" & Str(Button) & ", " & Str(Shift) & ", " & Str(X) & ", " & Str(Y)
End Sub


Private Sub Command1_Click()
Check1.Caption = "Value " & Str(Check1.Value)
End Sub

Private Sub Command4_Click()
Check1.Visible = Not Check1.Visible
End Sub

Private Sub Command5_Click()
Check1.Left = Check1.Left + 20
Check1.Top = Check1.Top + 20
End Sub

Private Sub Command6_Click()
Check1.Width = Check1.Width + 20
Check1.Height = Check1.Height + 50
End Sub

Private Sub Command7_Click()
Check1.Enabled = Not Check1.Enabled
End Sub

