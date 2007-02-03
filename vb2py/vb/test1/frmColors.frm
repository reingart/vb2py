VERSION 5.00
Begin VB.Form frmColors 
   BackColor       =   &H0000FF00&
   Caption         =   "Colorful form"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmColors.frx":0000
      Left            =   840
      List            =   "frmColors.frx":0013
      TabIndex        =   3
      Text            =   "Combo1"
      ToolTipText     =   "So should this"
      Top             =   2040
      Width           =   1935
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0C000&
      ForeColor       =   &H000080FF&
      Height          =   2010
      ItemData        =   "frmColors.frx":0033
      Left            =   3480
      List            =   "frmColors.frx":0046
      TabIndex        =   2
      ToolTipText     =   "Do you see the tip"
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000000FF&
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Text            =   "Text1"
      ToolTipText     =   "This should have a tip"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Command1"
      Height          =   855
      Left            =   720
      TabIndex        =   0
      ToolTipText     =   "Another tip"
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Should see color + tooltips"
      Height          =   615
      Left            =   720
      TabIndex        =   5
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Can you hear me now?"
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      ToolTipText     =   "Am I tipped?"
      Top             =   2760
      Width           =   2655
   End
End
Attribute VB_Name = "frmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' A form to test if colours are working correctly
