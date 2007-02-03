VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Tree view test form"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkEnableAutoEdit 
      Caption         =   "Enable auto edit"
      Height          =   255
      Left            =   6840
      TabIndex        =   23
      Top             =   4080
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CheckBox chkAllowEdits 
      Caption         =   "Allow label edits"
      Height          =   255
      Left            =   6840
      TabIndex        =   22
      Top             =   3720
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Caption         =   "Test node methods"
      Height          =   2775
      Left            =   6840
      TabIndex        =   18
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtResults 
         Height          =   1455
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   21
         Text            =   "tree.frx":0000
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtNodeName 
         Height          =   375
         Left            =   1200
         TabIndex        =   20
         Text            =   "A2"
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdTestNode 
         Caption         =   "Test"
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pictures / Image list"
      Height          =   1815
      Left            =   240
      TabIndex        =   13
      Top             =   6600
      Width           =   5535
      Begin VB.CommandButton cmdSetAsDynamic 
         Caption         =   "Set as dynamic"
         Height          =   495
         Left            =   2040
         TabIndex        =   17
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdSetAsPreload 
         Caption         =   "Set as preload"
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdLoadPicture 
         Caption         =   "Load pictures"
         Height          =   495
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdSetPictures 
         Caption         =   "Set pictures"
         Height          =   495
         Left            =   2040
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCollapse 
      Caption         =   "Collapse All"
      Height          =   495
      Left            =   2640
      TabIndex        =   12
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdExpand 
      Caption         =   "Expand All"
      Height          =   495
      Left            =   960
      TabIndex        =   11
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdAddTree 
      Caption         =   "Add Tree"
      Height          =   495
      Left            =   5040
      TabIndex        =   10
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox txtTree 
      Height          =   2295
      Left            =   6840
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "tree.frx":000A
      Top             =   4920
      Width           =   2895
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Text            =   "Next node"
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   5040
      TabIndex        =   7
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   495
      Left            =   5040
      TabIndex        =   6
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdSize 
      Caption         =   "Size"
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move"
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdEnable 
      Caption         =   "Enable"
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdVisible 
      Caption         =   "Visible"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin ComctlLib.TreeView tvTree 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   9975
      _Version        =   327680
      Style           =   7
      Appearance      =   1
   End
   Begin ComctlLib.ImageList imPreload 
      Left            =   6000
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "tree.frx":0012
            Key             =   "closed"
            Object.Tag             =   "closedicon"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "tree.frx":01EC
            Key             =   "open"
            Object.Tag             =   "openicon"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ilDynamic 
      Left            =   6000
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327680
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Make sure we import the common controls for Python
' VB2PY-GlobalAdd: CustomIncludes.ComctlLib = comctllib


Private Sub chkEnableAutoEdit_Click()
If chkEnableAutoEdit.Value Then
    tvTree.LabelEdit = tvwAutomatic
Else
    tvTree.LabelEdit = tvwManual
End If
End Sub

Private Sub cmdAdd_Click()
If tvTree.SelectedItem Is Nothing Then
    tvTree.Nodes.Add , , txtName.Text, txtName.Text
Else
    tvTree.Nodes.Add tvTree.SelectedItem.Key, tvwChild, txtName.Text, txtName.Text
End If
End Sub

Private Sub cmdAddTree_Click()
tvTree.Nodes.Clear
setTree txtTree.Text
End Sub

Private Sub cmdClear_Click()
tvTree.Nodes.Clear
End Sub


Private Sub cmdEnable_Click()
tvTree.Enabled = Not tvTree.Enabled
End Sub

Private Sub cmdExpand_Click()
For Each Node In tvTree.Nodes
    Node.Expanded = True
Next Node
End Sub

Private Sub cmdCollapse_Click()
For Each Node In tvTree.Nodes
    Node.Expanded = False
Next Node
End Sub


Private Sub cmdLoadPicture_Click()
ilDynamic.ListImages.Add , "closed", LoadPicture(App.Path & "\closedicon.ico")
ilDynamic.ListImages.Add , "open", LoadPicture(App.Path & "\openicon.ico")
End Sub

Private Sub cmdMove_Click()
tvTree.Left = tvTree.Left + 10
tvTree.Top = tvTree.Top + 10
End Sub

Private Sub cmdRemove_Click()
If tvTree.SelectedItem Is Nothing Then
    MsgBox "No selection"
Else
    tvTree.Nodes.Remove tvTree.SelectedItem.Key
End If
End Sub

Private Sub cmdSetAsDynamic_Click()
Set tvTree.ImageList = ilDynamic
End Sub

Private Sub cmdSetAsPreload_Click()
Set tvTree.ImageList = imPreload
End Sub

Private Sub cmdSetPictures_Click()
Dim Nde As Node
For Each Nde In tvTree.Nodes
    Nde.Image = "closed"
    Nde.ExpandedImage = "open"
Next Nde
End Sub

Private Sub cmdSize_Click()
tvTree.Width = tvTree.Width + 10
tvTree.Height = tvTree.Height + 10
End Sub

Private Sub cmdTestNode_Click()
'
Dim This As Node
Set This = tvTree.Nodes(txtNodeName.Text)
This.Selected = True
txtResults.Text = "text:" & This.Text & vbCrLf & "tag:" & This.Tag & vbCrLf
txtResults.Text = txtResults.Text & "visible:" & This.Visible & vbCrLf & "children:" & This.Children & vbCrLf
If This.Children > 0 Then
    txtResults.Text = txtResults.Text & "childtext:" & This.Child.Text & vbCrLf
End If
txtResults.Text = txtResults.Text & "firstsib:" & This.FirstSibling.Text & vbCrLf & "lastsib:" & This.LastSibling.Text & vbCrLf
txtResults.Text = txtResults.Text & "path:" & This.FullPath & vbCrLf & "next:" & This.Next.Text & vbCrLf
txtResults.Text = txtResults.Text & "parent:" & This.Parent.Text & vbCrLf & "previous:" & This.Previous.Text & vbCrLf
txtResults.Text = txtResults.Text & "root:" & This.Root.Text & vbCrLf
This.EnsureVisible
This.Selected = True
'
End Sub

Private Sub cmdVisible_Click()
tvTree.Visible = Not tvTree.Visible
End Sub

Private Sub Form_Load()
txtTree.Text = "A=ROOT" & vbCrLf & _
               "A1=A" & vbCrLf & _
               "A2=A" & vbCrLf & _
               "A3=A" & vbCrLf & _
               "A3A=A3" & vbCrLf & _
               "A4=A" & vbCrLf & _
               "B=ROOT" & vbCrLf & _
               "B1=B"

End Sub

Sub setTree(Text As String)
' Set the tree up
' Get name
Dim Name As String, Last As Node, Remainder As String
'
Do While Text <> ""
    '
    posn = InStr(Text, vbCrLf)
    If posn <> 0 Then
        parts = strSplit(Text, vbCrLf)
        Name = parts(0) 'Left$(Text, posn - 1)
        Text = parts(1) ' Right$(Text, Len(Text) - posn - 1)
    Else
        Name = Text
        Text = ""
    End If
    '
    parts = strSplit(Name, "=")
    nodename = parts(0)
    parentname = parts(1)
    '
    If parentname = "ROOT" Then
        tvTree.Nodes.Add , , nodename, nodename
    Else
        tvTree.Nodes.Add parentname, tvwChild, nodename, nodename
    End If
    '
Loop
'
End Sub


Function strSplit(Text, Delim) As Variant
posn = InStr(Text, Delim)
Dim parts(1)
parts(0) = Left$(Text, posn - 1)
parts(1) = Right$(Text, Len(Text) - posn - Len(Delim) + 1)
strSplit = parts
End Function





Private Sub tvTree_AfterLabelEdit(Cancel As Integer, NewString As String)
Debug.Print "After label edit on " & tvTree.SelectedItem.Text & " new name is " & NewString
If NewString = "CCC" Then
    Debug.Print "Cancelled"
    Cancel = 1
Else
    Debug.Print "OK"
    Cancel = 0
End If
End Sub

Private Sub tvTree_BeforeLabelEdit(Cancel As Integer)
Debug.Print "Before label edit on " & tvTree.SelectedItem.Text
If chkAllowEdits.Value Then
    Cancel = 0
Else
    Cancel = 1
End If
End Sub

Private Sub tvTree_Click()
Debug.Print "Tree view click"
End Sub

Private Sub tvTree_Collapse(ByVal Node As ComctlLib.Node)
Debug.Print "Tree view collapse on " & Node.Text
End Sub

Private Sub tvTree_DblClick()
Debug.Print "Tree view double click"
End Sub

Private Sub tvTree_DragDrop(Source As Control, x As Single, y As Single)
Debug.Print "Tree view drag drop"
End Sub

Private Sub tvTree_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Debug.Print "Tree view drag over"
End Sub

Private Sub tvTree_Expand(ByVal Node As ComctlLib.Node)
Debug.Print "Tree expand on " & Node.Text
End Sub

Private Sub tvTree_GotFocus()
Debug.Print "Tree view got focus"
End Sub

Private Sub tvTree_KeyDown(KeyCode As Integer, Shift As Integer)
Debug.Print "Tree view keydown (code, shift) " & CStr(KeyCode) & ", " & CStr(Shift)
End Sub

Private Sub tvTree_KeyPress(KeyAscii As Integer)
Debug.Print "Tree view keypress (code) " & CStr(KeyAscii)
End Sub

Private Sub tvTree_KeyUp(KeyCode As Integer, Shift As Integer)
Debug.Print "Tree view keyup (code, shift) " & CStr(KeyCode) & ", " & CStr(Shift)
End Sub

Private Sub tvTree_LostFocus()
Debug.Print "Tree view lost focus"
End Sub


Private Sub tvTree_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Debug.Print "Tree view mouse down (button, shift, x, y) " & CStr(Button) & ", " & CStr(Shift) & ", " & CStr(x) & ", " & CStr(y)
End Sub

Private Sub tvTree_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Debug.Print "Tree view mouse move (button, shift, x, y) " & CStr(Button) & ", " & CStr(Shift) & ", " & CStr(x) & ", " & CStr(y)
End Sub


Private Sub tvTree_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Debug.Print "Tree view mouse up (button, shift, x, y) " & CStr(Button) & ", " & CStr(Shift) & ", " & CStr(x) & ", " & CStr(y)
End Sub

Private Sub tvTree_NodeClick(ByVal Node As ComctlLib.Node)
Debug.Print "Tree node click " & Node.Text
End Sub


