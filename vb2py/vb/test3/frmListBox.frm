VERSION 5.00
Begin VB.Form frmListBox 
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   14460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDump 
      Caption         =   "Dump out"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddFirst 
      Caption         =   "Add as first"
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   7200
      TabIndex        =   8
      Top             =   720
      Width           =   5775
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Delete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   6015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Make Visible"
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Move "
      Height          =   375
      Left            =   8640
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Size "
      Height          =   375
      Left            =   10080
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Enable"
      Height          =   375
      Left            =   11520
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
List1.AddItem "Item " & Str(List1.ListCount + 1)
End Sub

Private Sub cmdAddFirst_Click()
List1.AddItem "First " & Str(List1.ListCount), 0
End Sub

Private Sub cmdClear_Click()
List1.Clear
End Sub

Private Sub cmdDump_Click()
For i = 0 To List1.ListCount - 1
    Debug.Print i, List1.List(i)
Next i
End Sub

Private Sub Command4_Click()
List2.Visible = Not List2.Visible
End Sub

Private Sub Command5_Click()
List2.Left = List2.Left + 20
List2.Top = List2.Top + 20
End Sub

Private Sub Command6_Click()
List2.Width = List2.Width + 20
List2.Height = List2.Height + 50
End Sub

Private Sub Command7_Click()
List2.Enabled = Not List2.Enabled
End Sub

Private Sub Delete_Click()
List1.RemoveItem List1.ListIndex
End Sub

Private Sub List1_Click()
Log "Click: " + Str(List1.ListIndex)
End Sub

Private Sub List1_DblClick()
Log "DblClick: " + Str(List1.ListIndex)
End Sub

Private Sub List1_GotFocus()
Log "Got Focus"
End Sub

Private Sub List1_ItemCheck(Item As Integer)
Log "Item check" + Str(Item)
End Sub

Private Sub List1_LostFocus()
Log "Lost Focus"
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Log "MouseDown" & Str(Button) & ", " & Str(Shift) & ", " & Str(X) & ", " & Str(Y)
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Log "MouseMove" & Str(Button) & ", " & Str(Shift) & ", " & Str(X) & ", " & Str(Y)
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Log "MouseUp" & Str(Button) & ", " & Str(Shift) & ", " & Str(X) & ", " & Str(Y)
End Sub



Private Sub List1_Scroll()
Log "Scroll"
End Sub
