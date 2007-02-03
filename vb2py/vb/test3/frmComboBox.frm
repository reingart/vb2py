VERSION 5.00
Begin VB.Form frmComboBox 
   Caption         =   "ComboBox"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13500
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   9480
      TabIndex        =   12
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Delete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   10680
      TabIndex        =   11
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   12000
      TabIndex        =   10
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdAddFirst 
      Caption         =   "Add as first"
      Height          =   375
      Left            =   8160
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdDump 
      Caption         =   "Dump out"
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "frmComboBox.frx":0000
      Left            =   240
      List            =   "frmComboBox.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1200
      Width           =   3015
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      ItemData        =   "frmComboBox.frx":001E
      Left            =   3480
      List            =   "frmComboBox.frx":0028
      TabIndex        =   6
      Text            =   "Combo2"
      Top             =   480
      Width           =   3015
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmComboBox.frx":003D
      Left            =   240
      List            =   "frmComboBox.frx":0047
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   480
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmComboBox.frx":005A
      Left            =   7080
      List            =   "frmComboBox.frx":0064
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   960
      Width           =   5775
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Enable"
      Height          =   375
      Left            =   11400
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Size "
      Height          =   375
      Left            =   9960
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Move "
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Make Visible"
      Height          =   375
      Left            =   7080
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
Combo1.Visible = Not Combo1.Visible
End Sub

Private Sub Command5_Click()
Combo1.Left = Combo1.Left + 50
Combo1.Top = Combo1.Top + 50
End Sub

Private Sub Command6_Click()
Combo1.Width = Combo1.Width + 50
End Sub

Private Sub Command7_Click()
Combo1.Enabled = Not Combo1.Enabled
End Sub


Private Sub Combo1_Change()
Log "Change, '" & Combo1.Text & "'"
End Sub

Private Sub Combo1_Click()
Log "Click"
End Sub

Private Sub Combo1_DblClick()
Log "DblClick"
End Sub

Private Sub Combo1_GotFocus()
Log "GotFocus"

End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
Log "Keydown" + ", " + Str(KeyCode) + ", " + Str(Shift)
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
Log "KeyPress" + ", " + Str(KeyCode) + ", " + Str(Shift) + ", " + Combo1.Text
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
Log "KeyUp" + ", " + Str(KeyCode) + ", " + Str(Shift)
End Sub

Private Sub Combo1_LostFocus()
Log "LostFocus"
End Sub

Private Sub Combo1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Log "MouseDown" & Str(Button) & ", " & Str(Shift) & ", " & Str(X) & ", " & Str(Y)
End Sub

Private Sub Combo1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Log "MouseMove" & Str(Button) & ", " & Str(Shift) & ", " & Str(X) & ", " & Str(Y)
End Sub

Private Sub Combo1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Log "MouseUp" & Str(Button) & ", " & Str(Shift) & ", " & Str(X) & ", " & Str(Y)
End Sub

Private Sub cmdAdd_Click()
Combo1.AddItem "Item " & Str(Combo1.ListCount + 1)
End Sub

Private Sub cmdAddFirst_Click()
Combo1.AddItem "First " & Str(Combo1.ListCount), 0
End Sub

Private Sub cmdClear_Click()
Combo1.Clear
End Sub

Private Sub cmdDump_Click()
For i = 0 To Combo1.ListCount - 1
    Debug.Print i, Combo1.List(i)
Next i
End Sub

Private Sub Delete_Click()
Combo1.RemoveItem Combo1.ListIndex
End Sub

