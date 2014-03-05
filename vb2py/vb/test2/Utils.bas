Attribute VB_Name = "Utils"
Public Function SplitOnWord(Text, Letter)
posn = 1
Dim Words()
ReDim Words(0)
For i = 1 To Len(Text)
    If Mid(Text, i, 1) = Letter Or i = Len(Text) Then
        ReDim Preserve Words(UBound(Words) + 1)
        If i = Len(Text) Then
            ThisWord = Mid(Text, posn, i - posn + 1)
        Else
            ThisWord = Mid(Text, posn, i - posn)
        End If
        Words(UBound(Words) - 1) = ThisWord
        posn = i + 1
    End If
Next i
'
ReDim Preserve Words(UBound(Words) - 1)
SplitOnWord = Words
'
End Function


Public Function Factorial(n)
If n = 0 Then
    Factorial = 1
Else
    Factorial = n * Factorial(n - 1)
End If
End Function

