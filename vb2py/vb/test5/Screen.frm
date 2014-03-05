VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Load form 2"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   720
      Top             =   3120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start timer"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'
Timer1.Interval = 500
Timer1.Enabled = True
'
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Timer1_Timer()
'
If Screen.ActiveControl Is Nothing Then
    Debug.Print "No active control"
Else
    Debug.Print "Active control:", Screen.ActiveControl.Name
End If
Debug.Print "Active form:", Screen.ActiveForm.Name
Debug.Print "Font count", Screen.FontCount
Debug.Print "Font 5:", Screen.Fonts(5)
Debug.Print "Height:", Screen.Height
Debug.Print "Mouse icon Height:", Screen.MouseIcon.Height
Debug.Print "Mouse pointer:", Screen.MousePointer
Debug.Print "Twips per pixel x and y:", Screen.TwipsPerPixelX, Screen.TwipsPerPixelY
Debug.Print "Width:", Screen.Width
'
End Sub
