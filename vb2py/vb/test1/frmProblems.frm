VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmProblems 
   Caption         =   "Form with problems"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ListView ListView1 
      Height          =   975
      Left            =   2280
      TabIndex        =   3
      Top             =   3960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1720
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   1815
      Left            =   4920
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   3201
      _Version        =   327682
      Style           =   7
      Appearance      =   1
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   2280
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   1200
   End
   Begin VB.OLE OLE1 
      Class           =   "Paint.Picture"
      Height          =   495
      Left            =   2280
      OleObjectBlob   =   "frmProblems.frx":0000
      TabIndex        =   9
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "List view"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "OLE Container"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Square"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "File list box"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Tree view"
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   360
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   $"frmProblems.frx":48018
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmProblems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This form has controls which don't render correctly
