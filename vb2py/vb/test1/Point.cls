VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public x As Single
Public y As Single
Public SubPoint As Point

Public Function getLength()
getLength = Sqr(x ^ 2 + y ^ 2)
End Function


