Attribute VB_Name = "Module1"

Public Type hh
    a As Integer
    b As Long
    c As Variant
    d As Variant
End Type

Public Sub sub1(x As String)
Select Case x
Case "one"
    y = 1
Case "two"
    y = 2
Case "three", "four", "five"
    z = 3
Case "a" To "z", 10 To 20
    z = -1
End Select
End Sub


Sub sub2()
a = 0
Do Until a > 10
    Debug.Print a
    a = a + 1
Loop
End Sub


Sub change(ByVal x, y)
x = x + 1
y = y + 1
End Sub

Sub wh()
a = "hellothere"
b = Mid(a, 1)
c = Mid(a, 2)
d = Mid(a, 5)
Debug.Print a, b, c, d
End Sub

Sub tif(Optional x = 1, Optional y = 2, Optional z = 3)
Debug.Print x, y, z
End Sub

Function dummy(Optional y = 20)
dummy = 2 * y
End Function

Function factorial(x)
GoTo 40
    If x = 0 Then
        factorial = 1
    Else
        factorial = factorial(x - 1) * x
    End If
label:
20:
40: factorial = 6
a = 1: b = 2: c = 2
a = 1:::: ReDim Preserve a(10), b(10)

End Function

