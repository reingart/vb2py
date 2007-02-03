VERSION 5.00
Begin VB.Form frmSettings 
   Caption         =   "Settings"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSection 
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Text            =   "Main"
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdGetAll 
      Caption         =   "Get All"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ListBox lstSettings 
      Height          =   2985
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtValue 
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "Get"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Text            =   "MySetting"
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
DeleteSetting "Testing", txtSection.Text, txtName.Text
End Sub

Private Sub cmdGet_Click()
txtValue.Text = GetSetting("Testing", txtSection.Text, txtName.Text, "<default>")
End Sub

Private Sub cmdGetAll_Click()
Dim Setting, Settings
lstSettings.Clear
Settings = GetAllSettings("Testing", txtSection.Text)
For Setting = 0 To UBound(Settings)
    lstSettings.AddItem Settings(Setting, 0) & " = " & Settings(Setting, 1)
Next Setting
End Sub

Private Sub cmdSet_Click()
SaveSetting "Testing", txtSection.Text, txtName.Text, txtValue.Text
End Sub
