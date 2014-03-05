VERSION 5.00
Begin VB.Form frmButton 
   Caption         =   "Buttons"
   ClientHeight    =   1620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   1620
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      Caption         =   "Mimic Mouse Move"
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command10 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Default"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "No Tab Stop"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Enable"
      Height          =   375
      Left            =   9840
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Size Button"
      Height          =   375
      Left            =   8400
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Move Button"
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Make Visible"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Mouse Cursor"
      Height          =   375
      Left            =   3840
      MousePointer    =   2  'Cross
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "Back Colour + Font"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simple"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Log "Click"
End Sub

Private Sub Command1_GotFocus()
Log "Got focus"
End Sub

Private Sub Command1_LostFocus()
Log "LostFocus"
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Log "MouseDown" & Str(Button) & ", " & Str(Shift) & ", " & Str(X) & ", " & Str(Y)
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Log "MouseMove" & Str(Button) & ", " & Str(Shift) & ", " & Str(X) & ", " & Str(Y)
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Log "MouseUp" & Str(Button) & ", " & Str(Shift) & ", " & Str(X) & ", " & Str(Y)
End Sub



Private Sub Command10_Click()
Log "Hit cancel"
End Sub

Private Sub Command11_Click()
Command1_MouseMove 10, 20, 30, 40
Command1_MouseUp 10, 20, 30, 40
Command1_MouseDown 10, 20, 30, 40
End Sub

Private Sub Command4_Click()
Command1.Visible = Not Command1.Visible
End Sub

Private Sub Command5_Click()
Command5.Top = Command5.Top + 30
Command5.Left = Command5.Left + 30
End Sub

Private Sub Command6_Click()
Command6.Height = Command6.Height + 30
Command6.Width = Command6.Width + 30
End Sub

Private Sub Command7_Click()
Command1.Enabled = Not Command1.Enabled
End Sub

Private Sub Command9_Click()
Log "Hit default"
End Sub

