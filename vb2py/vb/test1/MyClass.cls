VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Started As Boolean
Public Name As String
Private Size As Integer


Private Sub Class_Initialize()
'
Started = False
Name = "Default"
Size = 10
'
End Sub

Public Sub startWork(SizeIncrement As Integer, Optional NewName = "Working Object")
Started = True
Size = Size + SizeIncrement
Name = NewName
End Sub


Public Sub stopWork()
Started = False
End Sub

Public Function isWorking()
isWorking = Started
End Function

