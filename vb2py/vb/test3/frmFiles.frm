VERSION 5.00
Begin VB.Form frmFiles 
   Caption         =   "File Testing"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Remove directory"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdMakeDir 
      Caption         =   "Make directory"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "CheckTest file"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdMakeTestFile 
      Caption         =   "Make Test file"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Test file"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdListFiles 
      Caption         =   "List Files"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtDir 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "c:\temp"
      Top             =   120
      Width           =   6615
   End
   Begin VB.ListBox lstFiles 
      Height          =   3960
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "frmFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCheck_Click()
Dim chn As Integer
chn = FreeFile
Open "test.txt" For Input As #chn
Input #chn, test
If test = "Ok!" Then
    MsgBox "It worked"
Else
    MsgBox "It didn't work"
End If
Close #chn
End Sub

Private Sub cmdDelete_Click()
Kill "test.txt"
End Sub

Private Sub cmdListFiles_Click()
lstFiles.Clear
Dim Name As String
Name = Dir(txtDir.Text & "\*")
Do While Name <> ""
    lstFiles.AddItem Name
    Name = Dir()
Loop
End Sub

Private Sub cmdMakeDir_Click()
MkDir txtDir.Text
End Sub

Private Sub cmdMakeTestFile_Click()
Dim chn As Integer
chn = FreeFile
Open "test.txt" For Output As #chn
Print #chn, "Ok!"
Close #chn
End Sub

Private Sub cmdSet_Click()
ChDir txtDir.Text
End Sub


Private Sub Command1_Click()
RmDir txtDir.Text
End Sub
