VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Converter"
   ClientHeight    =   6765
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10245
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Python"
      Height          =   6495
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   4935
      Begin VB.TextBox Text1 
         Height          =   5655
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   4695
      End
      Begin VB.CommandButton cmdToFile 
         Caption         =   "To file ..."
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   6000
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "VB"
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton cmdFromFile 
         Caption         =   "From file ..."
         Height          =   375
         Left            =   3600
         TabIndex        =   2
         Top             =   6000
         Width           =   1215
      End
      Begin VB.TextBox txtVB 
         Height          =   5655
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Proper&ties"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuShowHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFromFile_Click()
t = 0
For i = 0 To 100
    t = t + i
Next i
End Sub
