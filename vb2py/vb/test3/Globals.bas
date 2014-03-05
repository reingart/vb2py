Attribute VB_Name = "Globals"
Public Sub Log(Text)
Debug.Print Time & " - " & Text
End Sub



Public Sub test()
Dim a() As String
b = UBound(a)
ReDim a(10)
a(1) = "hello"
a(10) = "bye"
c = a(1)
d = a(10)
End Sub

Public Function Factorial(X As Integer)
If X = 0 Then
    Factorial = 1
Else
    Factorial = X * Factorial(X - 1)
End If
End Function

Public Function Add(X As Single, Y As Single)
Add = X + Y
End Function


Public Sub EraseTest()
Dim a(10, 2) As Integer
For i = 1 To 10
    For j = 1 To 2
        a(i, j) = i + j
    Next j
Next i
Debug.Print GetArrayRepr(a)
Erase a
Debug.Print GetArrayRepr(a)
End Sub

Function GetArrayRepr(Arr)
total = 0
For i = 1 To UBound(Arr, 1)
    For j = 1 To UBound(Arr, 2)
       total = total + Arr(i, j)
    Next j
Next i
GetArrayRepr = total
End Function
