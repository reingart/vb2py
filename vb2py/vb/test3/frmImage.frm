VERSION 5.00
Begin VB.Form frmImage 
   Caption         =   "Image"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   11550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Load New"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Strech"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Enable"
      Height          =   375
      Left            =   8040
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Size "
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Move "
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Make Visible"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1380
      Left            =   240
      Picture         =   "frmImage.frx":0000
      Top             =   240
      Width           =   2370
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastPicture As String
Dim RootDir As String

Private Sub Command1_Click()
Image1.Stretch = Not Image1.Stretch
Label1.Caption = "Strech = " & Str(Image1.Stretch)
End Sub

Private Sub Command2_Click()
If LastPicture = "vb2py.gif" Then
    Image1.Picture = LoadPicture(RootDir & "/vb2pylogosm.jpg")
    LastPicture = "vb2pylogosm.jpg"
Else
    Image1.Picture = LoadPicture(RootDir & "/vb2py.gif")
    LastPicture = "vb2py.gif"
End If
Label2.Caption = "Showing " & LastPicture
End Sub

Public Sub Form_Load()
RootDir = App.Path
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Log "MouseDown" & Str(Button) & ", " & Str(Shift) & ", " & Str(X) & ", " & Str(Y)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Log "MouseMove" & Str(Button) & ", " & Str(Shift) & ", " & Str(X) & ", " & Str(Y)
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Log "MouseUp" & Str(Button) & ", " & Str(Shift) & ", " & Str(X) & ", " & Str(Y)
End Sub

Private Sub Command4_Click()
Image1.Visible = Not Image1.Visible
End Sub

Private Sub Command5_Click()
Image1.Left = Image1.Left + 20
Image1.Top = Image1.Top + 20
End Sub

Private Sub Command6_Click()
Image1.Width = Image1.Width + 20
Image1.Height = Image1.Height + 50
End Sub

Private Sub Command7_Click()
Image1.Enabled = Not Image1.Enabled
End Sub

