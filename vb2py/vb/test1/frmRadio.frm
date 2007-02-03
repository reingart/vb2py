VERSION 5.00
Begin VB.Form frmRadio 
   Caption         =   "Radio buttons"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3375
      Left            =   3120
      TabIndex        =   5
      Top             =   360
      Width           =   3975
      Begin VB.Frame Frame2 
         Caption         =   "Frame1"
         Height          =   1335
         Left            =   360
         TabIndex        =   8
         Top             =   1440
         Width           =   2655
         Begin VB.OptionButton Option8 
            Caption         =   "Real Inner 1"
            Height          =   255
            Left            =   720
            TabIndex        =   10
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Real Inner 2"
            Height          =   375
            Left            =   600
            TabIndex        =   9
            Top             =   840
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Inner 2"
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Inner 1"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   3960
      Width           =   1575
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4"
      Height          =   400
      Left            =   1000
      TabIndex        =   3
      Top             =   2640
      Width           =   2000
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   400
      Left            =   1000
      TabIndex        =   2
      Top             =   1920
      Width           =   2000
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   400
      Left            =   1000
      TabIndex        =   1
      Top             =   1320
      Value           =   -1  'True
      Width           =   2000
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   400
      Left            =   1000
      TabIndex        =   0
      Top             =   600
      Width           =   2000
   End
End
Attribute VB_Name = "frmRadio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This form has radio buttons on it ... these are tougher than they look
