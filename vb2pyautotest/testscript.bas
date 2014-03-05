Dim Results()
ReDim Results(0)
for each x in Array(0, 1, 2, 5, 10)
	Redim Preserve Results(UBound(Results)+1)
	Answer = Factorial(x)
	Results(Ubound(Results)) = Str(Answer) & "," & Str(x)
Next x
Chn = NextFile
Open 'test_Globals_Factorial_vb.txt' For Output As #Chn
Print #Chn, "# vb2Py Autotest results"
For Each X In Results
    Print #Chn, X
Next X
Close #Chn
Dim Results()
ReDim Results(0)
for each x in Array(-10.0, -5.5, -1.5, 0.0, 1.5, 5.5, 10.0)
	for each y in Array(-10.0, -5.5, -1.5, 0.0, 1.5, 5.5, 10.0)
		Redim Preserve Results(UBound(Results)+1)
		Answer = Add(x,y)
		Results(Ubound(Results)) = Str(Answer) & "," & Str(x) & "," & Str(y)
Next x
Next y
Chn = NextFile
Open 'test_Globals_Add_vb.txt' For Output As #Chn
Print #Chn, "# vb2Py Autotest results"
For Each X In Results
    Print #Chn, X
Next X
Close #Chn