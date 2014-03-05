VERSION 5.00
Begin VB.Form frmTextBox 
   Caption         =   "TextBox"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14295
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   14295
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmTextBox.frx":0000
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      Text            =   "Movable one"
      Top             =   600
      Width           =   5775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Make Visible"
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Move "
      Height          =   375
      Left            =   8520
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Size "
      Height          =   375
      Left            =   9960
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Enable"
      Height          =   375
      Left            =   11400
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Text            =   "Coloured"
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Single"
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
Text3.Visible = Not Text3.Visible
End Sub

Private Sub Command5_Click()
Text3.Top = Text3.Top + 50
Text3.Left = Text3.Left + 50
End Sub

Private Sub Command6_Click()
Text3.Width = Text3.Width + 50
Text3.Height = Text3.Height + 50
End Sub

Private Sub Command7_Click()
Text3.Enabled = Not Text3.Enabled
End Sub

Private Sub Text1_Change()
Log "Change, '" & Text1.Text & "'"
End Sub

Private Sub Text1_Click()
Log "Click"
End Sub

Private Sub Text1_DblClick()
Log "DblClick"
End Sub

Private Sub Text1_GotFocus()
Log "GotFocus"

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Log "Keydown" + ", " + Str(KeyCode) + ", " + Str(Shift)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Log "KeyPress" + ", " + Str(KeyCode) + ", " + Str(Shift) + ", " + Text1.Text
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
Log "KeyUp" + ", " + Str(KeyCode) + ", " + Str(Shift)
End Sub

Private Sub Text1_LostFocus()
Log "LostFocus"
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Log "MouseDown" & Str(Button) & ", " & Str(Shift) & ", " & Str(X) & ", " & Str(Y)
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Log "MouseMove" & Str(Button) & ", " & Str(Shift) & ", " & Str(X) & ", " & Str(Y)
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Log "MouseUp" & Str(Button) & ", " & Str(Shift) & ", " & Str(X) & ", " & Str(Y)
End Sub

Private Sub Text4_Change()
Log "Change, '" & Text4.Text & "'"
End Sub
