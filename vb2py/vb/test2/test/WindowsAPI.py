from vb2py.vbfunctions import *


NoError = 0

def GetUserName():
    _ret = None
    # Buffer size for the return string.
    # Get return buffer space.
    # For getting user information.
    # Assign the buffer size constant to lpUserName.
    lpUserName = Space(lpnLength + 1)
    # Get the log-on name of the person using product.
    status = WNetGetUser(lpName, lpUserName, lpnLength)
    # See whether error occurred.
    if status == NoError:
        # This line removes the null character. Strings in C are null-
        # terminated. Strings in Visual Basic are not null-terminated.
        # The null character must be removed from the C strings to be used
        # cleanly in Visual Basic.
        lpUserName = Left(lpUserName, InStr(lpUserName, Chr(0)) - 1)
    else:
        # An error occurred.
        MsgBox('Unable to get the name.')
        sys.exit(0)
    # Display the name of the person logged on to the machine.
    _ret = lpUserName
    return _ret

# VB2PY (UntranslatedCode) Attribute VB_Name = "WindowsAPI"
# VB2PY (UntranslatedCode) Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
