VERSION 5.00
Begin VB.Form frmForm 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCheckLike 
      Caption         =   "Like ?"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtTwo 
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   840
      Width           =   4095
   End
   Begin VB.TextBox txtOne 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label lblLike 
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
End
Attribute VB_Name = "frmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' A form to test the Like function

Const a = 10, b = 2

Private Sub btnCheckLike_Click()
If txtOne.Text Like txtTwo.Text Then
    lblLike.Caption = "Yes"
Else
    lblLike.Caption = "No"
End If
End Sub
