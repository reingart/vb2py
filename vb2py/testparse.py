# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

txt = """
Dim myvar As Integer
Dim x(10) As String, b(10, MAXINT) As Single
Public secret

Public Sub DoIt(x, y, z As String)
	' Do some work
	Dim locVar As Integer
	Dim i, j, k
	'
	locVar = 0
	If x < 10 Then
		locVar = y
	ElseIf x > 20 Then
		locVar = y/2.0
	ElseIf z = 43 Then
		locVar = Nothing
		anotherVar = 21
		If Time = "Out" Then
			a = Failed
			b = Quit
		Else
			If Now = "Today" Then
				c = "Yes"
			End If
		End If
	Else
		locVar = z
	End If
	doACall(locVar)
	locVar = locVar*2/24
	locString = (locVar+2) & (locVar-(2/(10-3)))
	'
End Sub

"""
