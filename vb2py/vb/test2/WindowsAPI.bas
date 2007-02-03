Attribute VB_Name = "WindowsAPI"
Declare Function WNetGetUser Lib "mpr.dll" _
      Alias "WNetGetUserA" (ByVal lpName As String, _
      ByVal lpUserName As String, lpnLength As Long) As Long

   Const NoError = 0       'The Function call was successful

   Function GetUserName()

      ' Buffer size for the return string.
      Const lpnLength As Integer = 255

      ' Get return buffer space.
      Dim status As Integer

      ' For getting user information.
      Dim lpName, lpUserName As String

      ' Assign the buffer size constant to lpUserName.
      lpUserName = Space$(lpnLength + 1)

      ' Get the log-on name of the person using product.
      status = WNetGetUser(lpName, lpUserName, lpnLength)

      ' See whether error occurred.
      If status = NoError Then
         ' This line removes the null character. Strings in C are null-
         ' terminated. Strings in Visual Basic are not null-terminated.
         ' The null character must be removed from the C strings to be used
         ' cleanly in Visual Basic.
         lpUserName = Left$(lpUserName, InStr(lpUserName, Chr(0)) - 1)
      Else

         ' An error occurred.
         MsgBox "Unable to get the name."
         End
      End If

      ' Display the name of the person logged on to the machine.
      GetUserName = lpUserName

   End Function


