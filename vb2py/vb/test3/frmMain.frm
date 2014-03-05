VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main Navigation Form"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      Caption         =   "Erase Test"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Settings"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Random"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Files"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Test"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Image"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Check"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ListBox"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ComboBox"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TextBox"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Button"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmButton.Show
End Sub

Private Sub Command10_Click()
frmSettings.Show
End Sub

Private Sub Command11_Click()
EraseTest
End Sub

Private Sub Command2_Click()
frmTextBox.Show
End Sub

Private Sub Command3_Click()
frmComboBox.Show
End Sub

Private Sub Command4_Click()
frmListBox.Show
End Sub

Private Sub Command5_Click()
frmCheckBox.Show
End Sub

Private Sub Command6_Click()
frmImage.Show
End Sub

Private Sub Command7_Click()
test
End Sub

Private Sub Command8_Click()
frmFiles.Show
End Sub

Private Sub Command9_Click()
frmRandom.Show
End Sub
