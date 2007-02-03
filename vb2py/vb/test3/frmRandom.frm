VERSION 5.00
Begin VB.Form frmRandom 
   Caption         =   "Random"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNegCall 
      Caption         =   "Call rnd with -ve"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdGetRandom 
      Caption         =   "Get Random"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox txtResults 
      Height          =   3255
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmRandom.frx":0000
      Top             =   1200
      Width           =   5535
   End
   Begin VB.CommandButton cmdRandomize 
      Caption         =   "Randomize no seed"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdRandomizeWithSeed 
      Caption         =   "Randomize with seed"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtSeed 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "1234"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Seed"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmRandom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGetRandom_Click()
For i = 1 To 10
    txtResults.Text = txtResults.Text & CStr(i) & " : " & CStr(Rnd) & vbCrLf
Next i
End Sub

Private Sub cmdNegCall_Click()
a = Rnd(-1)
End Sub

Private Sub cmdRandomize_Click()
Randomize
End Sub

Private Sub cmdRandomizeWithSeed_Click()
Randomize Int(txtSeed.Text)
End Sub
