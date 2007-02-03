#
# Turn off logging in extensions (too loud!)
import vb2py.extensions
vb2py.extensions.disableLogging()

from unittest import *
from vb2py.vbparser import buildParseTree, VBParserError

#
# Set some config options which are appropriate for testing
import vb2py.config
Config = vb2py.config.VB2PYConfig()
Config.setLocalOveride("General", "ReportPartialConversion", "No")


tests = []

# << Parsing tests >> (1 of 60)
# Simple assignments
tests.append("""
a = 10
b = 20+30
c = "hello there"
oneVal = 10
twoVals = Array(10,20)
functioncall = myfunction.mymethod(10)
""")

# Set type assignments
tests.append("""
Set a = myobject
Set b = myobject.subobject
Set obj = function(10, 20, 30+40)
""")


# Set type assignments with "New" objects
tests.append("""
Set a = New myobject
Set b = New myobject.subobject
""")

# Assignments with tough parenthesis
tests.extend([
        "d=(((4*5)/2+10)-10)",
])

# Assignments with tough string quotes
tests.extend([
        'd="g""h""j"""',
])

# Assignments with tough strings in general
tests.extend([
        r'a="\"',  # The single slash is a killer
])
# << Parsing tests >> (2 of 60)
# Simple expressions
tests.extend([
'a = 10',
'a = 20+30',
'a = "hello there"',
'a = 10',
'a = Array(10,20)',
'a = myfunction.mymethod(10)',
'a = &HFF',
'a = &HFF&',
'a = #1/10/2000#',
'a = #1/10#',
'a = 10 Mod 2',
])


# Nested expressions
tests.extend(["a = 10+(10+(20+(30+40)))",
              "a = (10+20)+(30+40)",
              "a = ((10+20)+(30+40))",
])

# Conditional expressions
tests.extend(["a = a = 1",
              "a = a <> 10",
              "a = a > 10",
              "a = a < 10",
              "a = a <= 10",
              "a = a >= 10",
              "a = a = 1 And b = 2",
              "a = a = 1 Or b = 2",
              "a = a Or b",
              "a = a Or Not b",
              "a = Not a = 1",
              "a = Not a",
              "a = a Xor b",
              "a = b Is Nothing",
              "a = b \ 2",
              "a = b Like c",
              'a = "hello" Like "goodbye"',
])

# Things that failed
tests.extend([
            "a = -(x*x)",
            "a = -x*10",
            "a = 10 Mod 6",
            "Set NewEnum = mCol.[_NewEnum]",
])

# Functions
tests.extend([
            "a = myfunction",
            "a = myfunction()",
            "a = myfunction(1,2,3,4)",
            "a = myfunction(1,2,3,z:=4)",
            "a = myfunction(x:=1,y:=2,z:=4)",
            "a = myfunction(b(10))",
            "a = myfunction(b _\n(10))",
])

# String Functions
tests.extend([
            'a = Trim$("hello")',
            'a = Left$("hello", 4)',
])			

# Things that failed
tests.extend([
            "a = -(x*x)",
            "a = -x*10",
            "a = 10 Mod 6",
])

# Address of
tests.extend([
        "a = fn(AddressOf fn)",
        "a = fn(a, b, c, AddressOf fn)",
        "a = fn(a, AddressOf b, AddressOf c, AddressOf fn)",
        "a = fn(a, AddressOf b.m.m, AddressOf c.k.l, AddressOf fn)",
])

# Type of
tests.extend([
        "a = fn(TypeOf fn)",
        "a = fn(a, b, c, TypeOf fn)",
        "a = fn(a, TypeOf b, TypeOf c, TypeOf fn)",
        "a = fn(a, TypeOf b.m.m, TypeOf c.k.l, TypeOf fn)",
        "a = TypeOf Control Is This",
        "a = TypeOf Control Is This Or TypeOf Control Is That",])
# << Parsing tests >> (3 of 60)
# Using ByVal and ByRef in a call or expression
tests.extend([
'a = fn(ByVal b)',
'a = fn(x, y, z, ByVal b)',
'a = fn(x, y, z, ByVal b, 10, 20, 30)',
'a = fn(ByVal a, ByVal b, ByVal c)',
'a = fn(ByRef b)',
'a = fn(x, y, z, ByRef b)',
'a = fn(x, y, z, ByRef b, 10, 20, 30)',
'a = fn(ByRef a, ByRef b, ByRef c)',
'fn ByVal b',
'fn x, y, z, ByVal b',
'fn x, y, z, ByVal b, 10, 20, 30',
'fn ByVal a, ByVal b, ByVal c',
'fn ByRef b',
'fn x, y, z, ByRef b',
'fn x, y, z, ByRef b, 10, 20, 30',
'fn ByRef a, ByRef b, ByRef c',

])
# << Parsing tests >> (4 of 60)
# One line comments
tests.append("""
a = 10
' b = 20+30
' c = "hello there"
' oneVal = 10
twoVals = Array(10,20)
' functioncall = myfunction.mymethod(10)
""")

# One line comments with Rem
tests.append("""
a = 10
Rem b = 20+30
Rem not needed c = "hello there"
Rem opps oneVal = 10
twoVals = Array(10,20)
Rem dont do this anymore functioncall = myfunction.mymethod(10)
""")

# In-line comments
tests.append("""
a = 10
b = 20+30 ' comment
c = "hello there" ' another comment
oneVal = 10 ' yet another comment
twoVals = Array(10,20)
functioncall = myfunction.mymethod(10)
""")

# In-line comments with Rem
tests.append("""
a = 10
b = 20+30 Rem comment
c = "hello there" Rem another comment
oneVal = 10 Rem yet another comment
twoVals = Array(10,20)
functioncall = myfunction.mymethod(10)
""")

# Things which aren't comments
tests.append("""
a = "hello, this might ' look like ' a comment ' "
b = "wow there are a lot of '''''''' these here"
""")

# tough inline comments
tests.extend([
    "Public a As Integer ' and a comment"
])

# comments in awkward places
tests.extend([
"""
If a =0 Then ' nasty comment
    b=1
End If ' other nasty comment
""",

"""
While a<0 ' nasty comment
    b=1
Wend ' other nasty comment
""",

"""
Select Case a ' nasty comment
Case 10 ' oops
    b=1
Case Else ' other nasty comment
    b = 2
End Select ' gotcha
""",

"""
For i = 0 To 100 ' nasty comment
    b=1
Next i ' other nasty comment
""",

"""
Sub a() ' nasty comment
    b=1
End Sub ' other nasty comment
""",

"""
Function f() ' nasty comment
    b=1
End Function ' other nasty comment
""",

])
# << Parsing tests >> (5 of 60)
# Directives
tests.extend([
    "' VB2PY-Set General.Blah = Yes",
    "' VB2PY-Set General.Blah = ___",
    "' VB2PY-Unset General.Blah",
    "' VB2PY-Add: General.Option = 10",
])
# << Parsing tests >> (6 of 60)
# Two line continuations
tests.append("""
a = _
10 + 20 + 30
b = 10/ _
25
c = (one + _
     two + three)
""")

# Milti-line continuations
tests.append("""
a = _
      10 + 20 + 30 _
    * 10/ _
      25
c = (one + _
     two + three) * _
     four.five()
""")

tests.extend(["""
Private Declare Function GetTempPathA Lib "kernel32" _
 (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
""",
"""
Function GetTempPathA _
(ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
End Function
""",

])
# << Parsing tests >> (7 of 60)
# Simple dims
tests.extend([
        "Dim A",
        "Dim B As String",
        "Dim variable As Object.OtherObj",
        "Dim Var As Variant",
        "Dim A As String * 100",
])

# Dims with New
tests.extend([
        "Dim A As New Object",
        "Dim B As New Collection",
])

# Multiple dims on one line
tests.extend([
        "Dim A, B, C, D, E, F",
        "Dim B As String, B As Long, B As Integer, B As String, B As String",
        "Dim variable As Object.OtherObj, B, C, D, E",
        "Dim Var As Variant",
        "Dim A, B, C As New Collection",
        "Dim E As New Collection, F As New Object, F, G",
        "Dim H As New Object, G As New Object",
])

# Array type dims
tests.extend([
        "Dim A()",
        "Dim B(10, 20, 30) As String",
        "Dim variable() As Object.OtherObj",
        "Dim Var(mysize) As Variant",
])

# Scoped dims
tests.extend([
        "Public A",
        "Private B As String",
        "Private A, B, C, D, E, F",
        "Private B As String, B As Long, B As Integer, B As String, B As String",
        "Private variable As Object.OtherObj, B, C, D, E",
        "Public Var As Variant",
])

# Static dims
tests.extend([
        "Static A",
        "Static B As String",
        "Static A, B, C, D, E, F",
        "Static B As String, B As Long, B As Integer, B As String, B As String",
        "Static variable As Object.OtherObj, B, C, D, E",
        "Static Var As Variant",
])
# << Parsing tests >> (8 of 60)
# Arrays
tests.extend([
    "Dim a(10)",
    "Dim a(0)",
    "Dim a(0), b(20), c(30)",
    "Dim a(10+20)",
    "Dim a(10+20, 1+3)",
    "Dim a(1 To 10)",
    "Dim a(1 To 10, 5 To 20)",
])

# Redims
tests.extend([
    "ReDim a(10)",
    "ReDim a(0)",
    "ReDim Preserve a(20)",
    "ReDim a(0), b(20), c(30)",
    "ReDim Preserve a(20), b(20)",
    "ReDim a(10+20)",
    "ReDim a(10+20, 1+3)",
    "ReDim a(1 To 10)",
    "ReDim a(1 To 10, 5 To 20)",
    "ReDim a(10).b(10)",
])


# Complex examples
tests.extend([
"""
With Obj
    ReDim .Child(10)
End With
""",
])
# << Parsing tests >> (9 of 60)
# Constants with different types
tests.extend([
    "Const a = 10",
    'Const a = "Hello"',
    "Const a = &HA1",
    "Const a = 1#",
    "Const a = 1%",
    "Const a = 1&",
    "Public Const a = 10",
    'Public Const a = "Hello"',
    "Public Const a = &HA1",
    "Public Const a = 1#",
    "Public Const a = 1%",
    "Public Const a = 1&",
    "Private Const a = 10",
    'Private Const a = "Hello"',
    "Private Const a = &HA1",
    "Private Const a = 1#",
    "Private Const a = 1%",
    "Private Const a = 1&",
])

# Constants
tests.extend([
        "Const A = 20",
        'Const B = "one"',
        "Private Const A = 1234.5 + 20",
        "Const a=10, b=20, c=30",
        "Private Const a=10, b=20, d=12345",
])

# Typed Constants
tests.extend([
        "Const A As Long = 20",
        'Const B As String = "one"',
        "Private Const A As Single = 1234.5 + 20",
        'Const a As Integer = 10, b As String = "hello", c As String * 10 = 43',
        'Private Const a As Integer = 10, b As String = "hello", c As String * 10 = 43',
])
# << Parsing tests >> (10 of 60)
# Odds and ends
tests.extend([
"Private WithEvents A As Button",
])
# << Parsing tests >> (11 of 60)
# Bare calls
tests.extend([
        "subr",
        "object.method",
        "object.method.method2.method",
])

# Explicit bare calls
tests.extend([
        "Call subr",
        "Call object.method",
        "Call object.method.method2.method",
])

# Bare calls with arguments
tests.extend([
        "subr 10, 20, 30",
        "object.method a, b, c+d, e",
        'object.method.method2.method 10, "hello", "goodbye" & name',
])

# Explicit calls with arguments
tests.extend([
        "Call subr(10, 20, 30)",
        "Call object.method(a, b, c+d, e)",
        'Call object.method.method2.method(10, "hello", "goodbye" & name)',
        "Call subr()",
])

# Bare calls with arguments and functions
tests.extend([
        "subr 10, 20, 30",
        "object(23).method a, b, c+d, e",
        'object.method(5, 10, 20).method2.method 10, "hello", "goodbye" & name',
])

# Bare calls with named arguments and functions
tests.extend([
        "subr 10, 20, z:=30",
        "object(23).method one:=a, two:=b, three:=c+d, four:=e",
        'object.method(5, 10, 20).method2.method 10, "hello", two:="goodbye" & name',
])

# Bare calls with ommitted arguments
tests.extend([
        "subr 10, , 30",
        "subr ,,,,0",
        "subr 10, , , , 5",
])
# << Parsing tests >> (12 of 60)
# labels
tests.extend([
    "label:",
    "label20:",
    "20:",
    "label: a=1",
    "20: a=1",
    "101: doit",
    "101:\ndoit",
    "102: doit now",
    "103: doit now, for, ever",
])

# Goto's
tests.extend([
    "GoTo Label",
    "GoTo 20",
    "GoTo Label:",
    "GoTo 20:",
])

# Structures with labels
tests.extend([
"""
101: If a < 10 Then
102:		b=1
103: End If
""",

"""
101: While a < 0
102:		b=1
103: Wend
""",

"""
101: Select Case a
102:		Case 10
103:			b= 1
104:		Case Else
105:			b=2
103: End Select
""",

"""
101: For i = 0 To 100
102:		b=1
103: Next i
""",

"""
101: Sub a()
102:		b=1
103: End Sub
""",

])

# Numeric labels don't even need a ':' ... aarg!
tests.extend([
"""
101 If a < 10 Then
102		b=1
103 End If
""",

"""
101 While a < 0
102		b=1
103 Wend
""",

"""
101 Select Case a
102		Case 10
103			b= 1
104		Case Else
105			b=2
103 End Select
""",

"""
101 For i = 0 To 100
102		b=1
103 Next i
""",

"""
101 Sub a()
102		b=1
103 End Sub
""",

])
# << Parsing tests >> (13 of 60)
# simple multi-line statements
tests.extend([
    "a = 10: b = 20",
    "a = 10: b = 20: c=1: d=1: e=2",
    "a=10:",
    "a=10: b=20:",
])

# Blocks on a line
tests.extend([
    "For i = 0 To 10: b=b+i: Next i",
    "If a > b Then a = 10: b = 20"
])


# Bug #809979 - Line ending with a colon fails 
tests.extend([
    "a = 10:\nb = 20",
    "a = 10: b = 20:\nc=1: d=1: e=2",
    "a=10:\nb=20:\nc=1",
])
# << Parsing tests >> (14 of 60)
# open statements
tests.extend([
    "Open fn For Output As 12",
    "Open fn For Output As #12",
    "Open fn For Input As 12",
    "Open fn For Input As #12",
    "Open fn.gk.gl() For Input As #NxtChn()",
    "Open fn For Append Lock Write As 23",
    "Close 1",
    "Close #1",
    "Close channel",
    "Close #channel",
    "Close",
    "Close\na=1",
    "Closet = 10",
])


# Bug #810968 Close #1, #2 ' fails to parse 
tests.extend([
    "Close #1, #2, #3, #4",
    "Close 1, 2, 3, 4",
    "Close #1, 2, #3, 4",
    "Close #one, #two, #three, #four",
    "Close one, two, three, four",
    "Close #1,#2,#3,#4",
    "Close   #1   ,   #2   ,   #3   ,   #4   ",
])
# << Parsing tests >> (15 of 60)
# print# statements
tests.extend([
    "Print 10",
    "Print #1, 10",
    "Print 10, 20, 30;",
    "Print #1, 10, 20, 30;",
    "Print #1, 10; 20; 30;",
    "Print #1, 10; 20; 30; 40, 50, 60, 70; 80; 90",
    "Print 10, 20, 30,",
    "Print 10, 20, 30",
    "Print",
    "Print ;;;",
    "Print ,,,",
    "Print 1,,,2,,,3,,,;",
    "Print #3,",
    "Print #3,;;;",
    "Print #3,,,",
    "Print #3,1,,,2,,,3,,,;",
])

# get# statements
tests.extend([
    "Get #1, a, b",
    "Get #1, , b",
])

# input # statements
tests.extend([
    "Input #1, a, b",
    "Input #1, b",
    "a = Input(20, #3)",
    "a = Input(20, #id)",
])

# line input # statements
tests.extend([
    "Line Input #1, b",
])


# Seek
tests.extend([
    "Seek #filenum, value",
    "10: Seek #filenum, value",
    "10: Seek #filenum, value ' comment",
    "Seek #filenum, value ' comment",
])
# << Parsing tests >> (16 of 60)
tests.extend([
    'Private Declare Function FileTimeToSystemTime Lib "kernel32" (ftFileTime As FILETIME, stSystemTime As SYSTEMTIME) As Long',
    'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)',
    'Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long',
    'Private Declare Function GetFileAttributes Lib "kernel32" _ \n(ByVal lpFileName As String) As Long',
    'Private Declare Function GetFileAttributes Lib "kernel32" _ \n(ByVal lpFileName As String, A) As Long',
    'Private Declare Function GetFileAttributes Lib "kernel32" _ \n(ByVal lpFileName As String , A) As Long',
    'Private Declare Function GetFileAttributes Lib "kernel32" _ \n(ByVal lpFileName As String ) As Long',
])
# << Parsing tests >> (17 of 60)
# General on error goto
tests.extend([
    "On Error GoTo 100",
    "On Error GoTo ErrTrap",
    "On Error GoTo 100 ' comment",
    "On Error GoTo ErrTrap ' comment",
    "100: On Error GoTo 100",
    "label: On Error GoTo ErrTrap",
    "100: On Error GoTo 100 ' comment",
    "label: On Error GoTo ErrTrap ' comment",
])

# General on error resume next
tests.extend([
    "On Error Resume Next",
    "On Error Resume Next ' comment",
    "100: On Error Resume Next",
    "label: On Error Resume Next",
    "100: On Error Resume Next ' comment",
    "label: On Error Resume Next ' comment",
])

# General on error goto - 
tests.extend([
    "On Error GoTo 0",
    "On Error GoTo 0 ' comment",
    "100: On Error GoTo 0",
    "100: On Error GoTo 0 ' comment",
])


# On something goto list 
tests.extend([
    "On var GoTo 20",
    "On var GoTo 10,20,30,40",
])

# Resume
tests.extend([
    "label: Resume Next",
    "Resume Next",
    "label: Resume Next ' Comment",
    "label: Resume 7",
    "Resume 7",
    "label: Resume 7 ' Comment",
    "label: Resume",
    "Resume\na=1",
    "label: Resume' Comment",
])

# General on local error resume next
tests.extend([
    "On Local Error Resume Next",
    "On Local Error Resume Next ' comment",
    "100: On Local Error Resume Next",
    "label: On Local Error Resume Next",
    "100: On Local Error Resume Next ' comment",
    "label: On Local Error Resume Next ' comment",
])

# Bug #809979 - On Error with : after the label fails 
tests.extend([
    "On Error GoTo 0:\na=1",
    "On Error GoTo 0: ' comment",
    "100: On Error GoTo 0:\na=1",
    "100: On Error GoTo 0: ' comment",
    "On Error GoTo lbl:\na=1",
    "On Error GoTo lbl: ' comment",
    "100: On Error GoTo lbl:\na=1",
    "100: On Error GoTo lbl: ' comment",
])
# << Parsing tests >> (18 of 60)
# Lines
tests.extend([
        "Line (10,20)-(30,40), 10, 20",
        "obj.Pset (10, 20), RGB(1,2,2)",
])

# Move
tests.extend([
        "Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2",
])
# << Parsing tests >> (19 of 60)
# General name test (rename a file)
tests.extend([
        "Name file As otherfile",
        "Name file & extension As otherfile",
        "Name file & extension As otherfile & otherextension",
        'Name path & "\origname.txt" As path & "\knewname.txt"',
])
# << Parsing tests >> (20 of 60)
# Attributes at the head of a file
tests.extend([
    'Attribute VB_Name = "frmMain"',
    'Attribute VB_GlobalNameSpace = False',
    'Attribute VB_Creatable = False',
    'Attribute VB_PredeclaredId = True',
    'Attribute VB_Exposed = False',
    'Attribute Me.VB_Exposed = False',
    'Attribute Me.VB_Exposed = False, 1, 2, 3, 4',
    'Attribute Me.VB_Exposed = False, "1", "2, 3,", 4',
])
# << Parsing tests >> (21 of 60)
# Attributes at the head of a file
tests.extend([
"""
Enum thing
    _one = 1
    _two = 2
    _three = 3
    _four = 4
End Enum
""",
"""
Enum thing
    _one
    _two
    _three
    _four
End Enum
""",
])
# << Parsing tests >> (22 of 60)
# Types
tests.extend([
"""
Private Type ShellFileInfoType
 hIcon As Long
 iIcon As Long
 dwAttributes As Long
 szDisplayName As String * 260
 szTypeName As String * 80
End Type
"""
])
# << Parsing tests >> (23 of 60)
# The Option statement
tests.extend([
"Option Base 0",
"Option Base 1",
"Option Explicit",
"Option String Compare",
"Option String Compare Text",
])
# << Parsing tests >> (24 of 60)
# The End statement
tests.extend([
"10: End",
"End",
"End ' wow this is it",
"10: End ' this is the end",
])

# If with an 'End' in there
tests.append("""
If a = 10 Then
    End
End If
""")

# Sub with an 'End' in there
tests.append("""
Sub doit()
 End
End Sub
""")

# Fn with an 'End' in there
tests.append("""
Function doit()
 End
End Function
""")

# With with an 'End' in there
tests.append("""
With obj
 End
End With
""")
# << Parsing tests >> (25 of 60)
# The Event statement
tests.extend([
"Event doit()",
"Public Event doit()",
"Private Event doit()",
"Public Event doit(a, b, c, e)",
"Public Event doit(a As Integer, b As Long, c(), e As Command.Button)",
])
# << Parsing tests >> (26 of 60)
# The Debug.Print statement
tests.extend([
"Debug.Print",
"Debug.Print a",
"Debug.Print a,b",
"Debug.Print a;b",
"Debug.Print a+10;b+20",
"Debug.Print a+20, b-20",
"Debug.Print a;b;",
])
# << Parsing tests >> (27 of 60)
# Recordset notation
tests.extend([
"RS!diskID = DriveID",
"RS!diskID = DriveID+10",
'RS!diskID = "DriveID"',
])
# << Parsing tests >> (28 of 60)
# Simple If
tests.append("""
If a = 10 Then
    b = 20
End If
If c < 1 Then
    d = 15
End If
""")

# Empty If
tests.append("""
If a = 10 Then
End If
""")

# Empty If with comments
tests.append("""
If a = 10 Then ' comment here
End If
""")

# Simple If with And/Or
tests.append("""
If a = 10 And k = "test" Then
    b = 20
End If
If c < 1 Or d Then
    d = 15
End If
""")

# Simple If with compount And/Or expression
tests.append("""
If (a = 10 And k = "test") And (c Or b Or e = 43.23) Then
    b = 20
End If
If (c < 1) Or d And e = "hello" Or e < "wow" Then
    d = 15
End If
""")

#  If Not
tests.append("""
If Not a = 10 Then
    b=2
End If
""")

#  If With labels and comment
tests.append("""
10: If Not a = 10 Then 'heres a comment
20:   	b=2  ' antoher here
30: End If ' here too
""")
# << Parsing tests >> (29 of 60)
# Simple If/Else
tests.append("""
If a = 10 Then
    b = 20
Else
    b = 10
End If
If c < 1 Then
    d = 15
Else
    d = -12
End If
""")

# Empty If/Else
tests.append("""
If a = 10 Then
Else
End If
""")

# Simple If with And/Or
tests.append("""
If a = 10 And k = "test" Then
    b = 20
Else
    b = 1234
End If
If c < 1 Or d Then
    d = 15
Else
    e = "hello"
End If
""")

# Simple If with compount And/Or expression
tests.append("""
If (a = 10 And k = "test") And (c Or b Or e = 43.23) Then
    b = 20
Else
    g = 12
End If
If (c < 1) Or d And e = "hello" Or e < "wow" Then
    d = 15
Else
    h = 1234
End If
""")
# << Parsing tests >> (30 of 60)
# Simple If/Else
tests.append("""
If a = 10 Then
    b = 20
ElseIf a < 10 Then
    b = 10
End If
If c < 1 Then
    d = 15
ElseIf c = 1 Then
    d = -12
End If
""")


# Simple If with And/Or
tests.append("""
If a = 10 And k = "test" Then
    b = 20
ElseIf b = -102 Then
    b = 1234
End If
If c < 1 Or d Then
    d = 15
ElseIf e = Myfunction Then
    e = "hello"
End If
""")

# Simple If with compount And/Or expression
tests.append("""
If (a = 10 And k = "test") And (c Or b Or e = 43.23) Then
    b = 20
ElseIf (a = 10 And k = "test") And (c Or b Or e = 43.23) Then
    g = 12
End If
If (c < 1) Or d And e = "hello" Or e < "wow" Then
    d = 15
ElseIf k < 43 Then
    h = 1234
End If
""")
# << Parsing tests >> (31 of 60)
# Simple If/Else
tests.append("""
If a = 10 Then
    b = 20
ElseIf a < 10 Then
    b = 10
Else
    b = 1111
End If
If c < 1 Then
    d = 15
ElseIf c = 1 Then
    d = -12
Else
    d = "wow"
End If
""")


# Simple If with And/Or
tests.append("""
If a = 10 And k = "test" Then
    b = 20
ElseIf b = -102 Then
    b = 1234
Else
    b = 4321
End If
If c < 1 Or d Then
    d = 15
ElseIf e = Myfunction Then
    e = "hello"
Else
    g = 1
End If
""")

# Simple If with compount And/Or expression
tests.append("""
If (a = 10 And k = "test") And (c Or b Or e = 43.23) Then
    b = 20
ElseIf (a = 10 And k = "test") And (c Or b Or e = 43.23) Then
    g = 12
Else
    k = 3234
End If
If (c < 1) Or d And e = "hello" Or e < "wow" Then
    d = 15
ElseIf k < 43 Then
    h = 1234
Else
    doIt
End If
""")
# << Parsing tests >> (32 of 60)
# Simple Nested If
tests.append("""
If a = 10 Then
    b = 20
    If c < 1 Then
        d = 15
    End If
End If
""")


# Complex nested If
tests.append("""
If (a = 10 And k = "test") And (c Or b Or e = 43.23) Then
    b = 20
ElseIf (a = 10 And k = "test") And (c Or b Or e = 43.23) Then
    If (c < 1) Or d And e = "hello" Or e < "wow" Then
        d = 15
    ElseIf k < 43 Then
        h = 1234
    Else
        If (a = 10 And k = "test") And (c Or b Or e = 43.23) Then
            b = 20
        End If
        If (c < 1) Or d And e = "hello" Or e < "wow" Then
            d = 15
        End If
    End If
Else
    k = 3234
End If
""")
# << Parsing tests >> (33 of 60)
# Inline ifs
tests.extend([
        "If a = 10 Then b = 20",
        "If a = 20 And b = 5 Then d = 123",
        "If a = 12 Then d = 1 Else g = 5",
        "If a = 10 Then doit",
        "If a = 10 Then doit 10, 20, 30",
        "If a = 10 Then doit Else dont",
        "If a = 10 Then doit 10, 20, 30 Else dont",
        "If a = 10 Then doit 10, 20, 30 Else dont 5, 10, 15",
        "If a = 10 Then Exit Function",
        "If a = 10 Then Exit Function Else DoIt",
        "If a = 10 Then Exit Function Else DoIt=1",
        "If a = 10 Then Exit Function Else DoIt 1, 2, 3",
        "If a = 10 Then DoIt Else Exit Function",
        "If a = 10 Then DoIt=1 Else Exit Function",
        "If a = 10 Then DoIt 1,2,34 Else Exit Function",
])

# Weird inline if followed by assignment that failed once
tests.extend([
        "If a = 10 Then b a\nc=1",
])
# << Parsing tests >> (34 of 60)
# #If
tests.append("""
#If a = 10 Then
    b = 20
#Else
    c=2
#End If
#If c < 1 Then
    d = 15
#Else
    c=2
#End If
""")

# Empty #If
tests.append("""
#If a = 10 Then
#Else
    c=2
#End If
""")

# Empty #If with comments
tests.append("""
#If a = 10 Then ' comment here
#Else
    c=2
#End If
""")

# Simple #If with And/Or
tests.append("""
#If a = 10 And k = "test" Then
    b = 20
#Else
    c=2
#End If
#If c < 1 Or d Then
    d = 15
#Else
    c=2
#End If
""")

# Simple #If with compount And/Or expression
tests.append("""
#If (a = 10 And k = "test") And (c Or b Or e = 43.23) Then
    b = 20
#Else
    c=2
#End If
#If (c < 1) Or d And e = "hello" Or e < "wow" Then
    d = 15
#Else
    c=2
#End If
""")

#  #If Not
tests.append("""
#If Not a = 10 Then
    b=2
#Else
    c=2
#End If
""")
# << Parsing tests >> (35 of 60)
# simple sub
tests.append("""
Sub MySub()
a=10
n=20
c="hello"
End Sub
""")


# simple sub with exit
tests.append("""
Sub MySub()
a=10
n=20
Exit Sub
c="hello"
End Sub
""")


# simple sub with scope
tests.extend(["""
Private Sub MySub()
a=10
n=20
c="hello"
End Sub""",
"""
Public Sub MySub()
a=10
n=20
c="hello"
End Sub
""",
"""
Friend Sub MySub()
a=10
n=20
c="hello"
End Sub
""",
"""
Private Static Sub MySub()
a=10
n=20
c="hello"
End Sub
""",
])

# simple sub with gap in ()
tests.append("""
Sub MySub(   )
a=10
n=20
c="hello"
End Sub
""")
# << Parsing tests >> (36 of 60)
# simple sub
tests.append("""
Sub MySub(x, y, z, a, b, c)
a=10
n=20
c="hello"
End Sub
""")


# simple sub with exit
tests.append("""
Sub MySub(x, y, z, a, b, c)
a=10
n=20
Exit Sub
c="hello"
End Sub
""")


# simple sub with scope
tests.append("""
Private Sub MySub(x, y, z, a, b, c)
a=10
n=20
c="hello"
End Sub
Public Sub MySub(x, y, z, a, b, c)
a=10
n=20
c="hello"
End Sub
""")
# << Parsing tests >> (37 of 60)
# simple sub
tests.append("""
Sub MySub(x As Single, y, z As Boolean, a, b As Variant, c)
a=10
n=20
c="hello"
End Sub
""")


# simple sub with exit
tests.append("""
Sub MySub(x As Single, y, z As Object, a, b As MyThing.Object, c)
a=10
n=20
Exit Sub
c="hello"
End Sub
""")


# simple sub with scope
tests.append("""
Private Sub MySub(x, y As Variant, z, a As Boolena, b, c As Long)
a=10
n=20
c="hello"
End Sub
Public Sub MySub(x, y, z, a, b, c)
a=10
n=20
c="hello"
End Sub
""")
# << Parsing tests >> (38 of 60)
# simple sub
tests.append("""
Sub MySub(x As Single, y, z As Boolean, a, Optional b As Variant, c)
a=10
n=20
c="hello"
End Sub
""")


# simple sub with exit
tests.append("""
Sub MySub(x() As Single, y, z As Object, Optional a, b As MyThing.Object, Optional c)
a=10
n=20
Exit Sub
c="hello"
End Sub
""")


# simple sub with scope
tests.append("""
Private Sub MySub(x, Optional y As Variant, Optional z, a As Boolena, b, c As Long)
a=10
n=20
c="hello"
End Sub
Public Sub MySub(x, y, z, a, b, c)
a=10
n=20
c="hello"
End Sub
""")

# simple sub with optional arguments and defaults
tests.append("""
Sub MySub(x As Single, y, z As Boolean, a, Optional b = 10, Optional c="hello")
a=10
n=20
c="hello"
End Sub
""")

# simple sub with optional arguments and defaults
tests.append("""
Sub MySub(x As Single, y, z As Boolean, a, Optional b = 10, Optional c As String = "hello")
a=10
n=20
c="hello"
End Sub
""")
# << Parsing tests >> (39 of 60)
# ByVal, ByRef args
tests.append("""
Sub MySub(ByVal a, ByRef y)
a=10
n=20
c="hello"
End Sub
""")

tests.append("""
Sub MySub(a, ByRef y)
a=10
n=20
c="hello"
End Sub
""")

tests.append("""
Sub MySub(ByVal a, y)
a=10
n=20
c="hello"
End Sub
""")

tests.append("""
Sub MySub(ByVal a As Single, y)
a=10
n=20
c="hello"
End Sub
""")
# << Parsing tests >> (40 of 60)
# 852166 Sub X<spc>(a,b,c) fails to parse 
tests.append("""
Sub MySub (ByVal a, ByRef y)
a=10
n=20
c="hello"
End Sub
""")

# 880612 Continuation character inside call  
tests.append("""
Sub MySub _
(ByVal a, ByRef y)
a=10
n=20
c="hello"
End Sub
""")
# << Parsing tests >> (41 of 60)
# simple fn
tests.append("""
Function MyFn()
a=10
n=20
c="hello"
MyFn = 20
End Function
""")


# simple fn with exit
tests.append("""
Function MyFn()
a=10
n=20
MyFn = 20
Exit Function
c="hello"
End Function
""")


# simple sub with scope
tests.extend(["""
Private Function MyFn()
a=10
n=20
c="hello"
MyFn = 20
End Function""",
"""
Public Function MyFn()
a=10
n=20
c="hello"
MyFn = 20
End Function
""",
"""
Friend Function MyFn()
a=10
n=20
c="hello"
MyFn = 20
End Function
""",
])

# simple fn with gap in ()
tests.append("""
Function MyFn(  )
a=10
n=20
c="hello"
MyFn = 20
End Function
""")
# << Parsing tests >> (42 of 60)
# simple sub
tests.append("""
Function MySub(x, y, z, a, b, c)
a=10
n=20
c="hello"
End Function
""")


# simple sub with exit
tests.append("""
Function MySub(x, y, z, a, b, c)
a=10
n=20
Exit Sub
c="hello"
End Function
""")


# simple sub with scope
tests.append("""
Private Function MySub(x, y, z, a, b, c)
a=10
n=20
c="hello"
End Function
Public Function fn(x, y, z, a, b, c)
a=10
n=20
c="hello"
End Function
""")
# << Parsing tests >> (43 of 60)
# simple sub
tests.append("""
Function fn(x As Single, y, z As Boolean, a, b As Variant, c) As Single
a=10
n=20
c="hello"
End Function
""")


# simple sub with exit
tests.append("""
Function fc(x As Single, y, z As Object, a, b As MyThing.Object, c) As Object.Obj
a=10
n=20
Exit Function
c="hello"
End Function
""")


# simple sub with scope
tests.append("""
Private Function MySub(x, y As Variant, z, a As Boolena, b, c As Long) As Variant
a=10
n=20
c="hello"
End Function
Public Function MySub(x, y, z, a, b, c) As String
a=10
n=20
c="hello"
End Function
""")

# function returning an array
tests.append("""
Function fn(x As Single, y, z As Boolean, a, b As Variant, c) As Single()
a=10
n=20
c="hello"
End Function
""")
# << Parsing tests >> (44 of 60)
# simple sub
tests.append("""
Function fn(x As Single, y, z As Boolean, a, Optional b As Variant, c) As Single
a=10
n=20
c="hello"
End Function
""")


# simple sub with exit
tests.append("""
Function MySub(x() As Single, y, z As Object, Optional a, b As MyThing.Object, Optional c) As Integer
a=10
n=20
Exit Function
c="hello"
End Function
""")


# simple sub with scope
tests.append("""
Private Function MySub(x, Optional y As Variant, Optional z, a As Boolena, b, c As Long) As Long
a=10
n=20
c="hello"
End Function
Public Function MySub(x, y, z, a, b, c) As Control.Buttons.BigButtons.ThisOne
a=10
n=20
c="hello"
End Function
""")

# simple fn with optional arguments and defaults
tests.append("""
Function MySub(x As Single, y, z As Boolean, a, Optional b = 10, Optional c="hello")
a=10
n=20
c="hello"
End Function
""")

# simple fn with optional arguments and defaults
tests.append("""
Function MySub(x As Single, y, z As Boolean, a, Optional b = 10, Optional c As String = "hello")
a=10
n=20
c="hello"
End Function
""")
# << Parsing tests >> (45 of 60)
# ByVal, ByRef args
tests.append("""
Function MySub(ByVal a, ByRef y)
a=10
n=20
c="hello"
End Function
""")

tests.append("""
Function MySub(a, ByRef y)
a=10
n=20
c="hello"
End Function
""")

tests.append("""
Function MySub(ByVal a, y)
a=10
n=20
c="hello"
End Function
""")

tests.append("""
Function MySub(ByVal a As Single, y)
a=10
n=20
c="hello"
End Function
""")
# << Parsing tests >> (46 of 60)
# Simple property let/get/set
tests.extend(["""
Property Let MyProp(NewVal As String)
 a = NewVal
 Exit Property
End Property
""",
"""
Property Get MyProp() As Long
 MyProp = NewVal
 Exit Property
End Property
""",
"""
Property Set MyProp(NewObject As Object) 
 Set MyProp = NewVal
 Exit Property
End Property
"""
"""
Public Property Let MyProp(NewVal As String)
 a = NewVal
End Property
""",
"""
Public Property Get MyProp() As Long
 MyProp = NewVal
End Property
""",
"""
Public Property Set MyProp(NewObject As Object) 
 Set MyProp = NewVal
End Property
""",
"""
Public Property Get MyProp(   ) As Long
 MyProp = NewVal
End Property
""",
])

# Simple property let/get/set with labels
tests.extend(["""
1: Property Let MyProp(NewVal As String)
1:  a = NewVal
1: End Property
""",
"""
1: Property Get MyProp() As Long
1:  MyProp = NewVal
1: End Property
""",
"""
1: Property Set MyProp(NewObject As Object) 
1:  Set MyProp = NewVal
1: End Property
"""
])

# Simple property let/get/set with labels and comment
tests.extend(["""
1: Property Let MyProp(NewVal As String) ' comment
1:  a = NewVal  ' comment
1: End Property  ' comment
""",
"""
1: Property Get MyProp() As Long  ' comment
1:  MyProp = NewVal  ' comment
1: End Property  ' comment
""",
"""
1: Property Set MyProp(NewObject As Object)   ' comment
1:  Set MyProp = NewVal  ' comment
1: End Property  ' comment
"""
])
# << Parsing tests >> (47 of 60)
# Simple case
tests.append("""
Select Case x
Case "one"
    y = 1
Case "two"
    y = 2
Case "three"
    z = 3
End Select
""")

# Simple case with else
tests.append("""
Select Case x
Case "one"
    y = 1
Case "two"
    y = 2
Case "three"
    z = 3
Case Else
    z = -1
End Select
""")

# Simple case with else and trailing colons
tests.append("""
Select Case x
Case "one":
    y = 1
Case "two":
    y = 2
Case "three":
    z = 3
Case Else:
    z = -1
End Select
""")

# Multiple case with else
tests.append("""
Select Case x
Case "one"
    y = 1
Case "two"
    y = 2
Case "three", "four", "five"
    z = 3
Case Else
    z = -1
End Select
""")

# Single line case with else
tests.append("""
Select Case x
Case "one": y = 1
Case "two": y = 2
Case "three", "four", "five": z = 3
Case Else: z = -1
End Select
""")


# Range case 
tests.append("""
Select Case x
Case "a" To "m"
    z = 1
Case "n" To "z"
    z = 20
End Select
""")

# Range case with Is and Like
tests.append("""
Select Case x
Case Is < "?", "a" To "m"
    z = 1
Case "n" To "z", Is > 10, Is Like "*blah"
    z = 20
End Select
""")

# Multiple Range case 
tests.append("""
Select Case x
Case "a" To "m", "A" To "G", "K" To "P"
    z = 1
Case "n" To "z", 10 To this.that(10,20)
    z = 20
End Select
""")

# Empty case
tests.append("""
    Select Case a
    Case 10
    Case 20
    End Select
""")

# Case with comments
tests.append("""
Select Case x
' Here is a nasty comment

Case "one"
    y = 1
Case "two"
    y = 2
Case "three"
    z = 3
End Select
""")
# << Parsing tests >> (48 of 60)
# Simple for
tests.append("""
For i = 0 To 1000
  a = a + 1
Next i
""")

# Simple for
tests.append("""
For i=0 To 1000
  a = a + 1
Next i
""")

# Empty for
tests.append("""
For i = 0 To 1000
Next i
""")

# Simple for with unnamed Next
tests.append("""
For i = 0 To 1000
  a = a + 1
Next
""")

# For with step
tests.append("""
For i = 0 To 1000 Step 2
  a = a + 1
Next i
""")

# For with exit
tests.append("""
For i = 0 To 1000
  a = a + 1
  Exit For
Next i
""")

# Nested for
tests.append("""
For i = 0 To 1000
  a = a + 1
  For j = 1 To i
     b = b + j
  Next j
Next i
""")
# << Parsing tests >> (49 of 60)
# Simple for
tests.append("""
For Each i In coll
  a = a + 1
Next i
""")

# Empty for
tests.append("""
For Each i In coll
Next i
""")

# Simple for with unnamed Next
tests.append("""
For Each i In coll
  a = a + 1
Next
""")


# For with exit
tests.append("""
For Each i In coll
  a = a + 1
  Exit For
Next i
""")

# Nested for
tests.append("""
For Each i In coll
  a = a + 1
  For Each j In coll
     b = b + j
  Next j
Next i
""")
# << Parsing tests >> (50 of 60)
# Simple while wend
tests.append("""
        a = 0
        While a < 10
            g = 10
            a = a + 1
        Wend
""")

# Nested while wend
tests.append("""
        a = 0
        While a < 10
            g = 10
            a = a + 1
            While b < 40
                doit
            Wend
        Wend
""")

# Simple while wend with line numbers
tests.append("""
1:		a = 0
2:		While a < 10
3:			g = 10
4:			a = a + 1
5:		Wend
""")
# << Parsing tests >> (51 of 60)
# Simple do while loop
tests.append("""
        a = 0
        Do While a < 10
            g = 10
            a = a + 1
        Loop
""")

# Simple do while with exit
tests.append("""
        a = 0
        Do While a < 10
            g = 10
            a = a + 1
            Exit Do
        Loop
""")

# Nested do while loop
tests.append("""
        a = 0
        Do While a < 10
            g = 10
            a = a + 1
            Do While b < 40
                doit
            Loop
        Loop
""")
# << Parsing tests >> (52 of 60)
# Simple do  loop
tests.append("""
        a = 0
        Do  
            g = 10
            a = a + 1
        Loop
""")

# Simple do  with exit
tests.append("""
        a = 0
        Do 
            g = 10
            a = a + 1
            Exit Do
        Loop
""")

# Nested do  loop
tests.append("""
        a = 0
        Do 
            g = 10
            a = a + 1
            Do 
                doit
            Loop
        Loop
""")
# << Parsing tests >> (53 of 60)
# Simple do  loop
tests.append("""
        a = 0
        Do  
            g = 10
            a = a + 1
        Loop While a < 10
""")

# Simple do  with exit
tests.append("""
        a = 0
        Do 
            g = 10
            a = a + 1
            Exit Do
        Loop While a <10
""")

# Nested do  loop
tests.append("""
        a = 0
        Do 
            g = 10
            a = a + 1
            Do 
                doit
            Loop While a <10
        Loop While a< 10
""")
# << Parsing tests >> (54 of 60)
# Simple do  loop
tests.append("""
        a = 0
        Do  
            g = 10
            a = a + 1
        Loop Until a < 10
""")

# Simple do  with exit
tests.append("""
        a = 0
        Do 
            g = 10
            a = a + 1
            Exit Do
        Loop Until a <10
""")

# Nested do  loop
tests.append("""
        a = 0
        Do 
            g = 10
            a = a + 1
            Do 
                doit
            Loop While a <10
        Loop Until a< 10
""")
# << Parsing tests >> (55 of 60)
# Simple do  loop
tests.append("""
        a = 0
        Do Until a < 10
            g = 10
            a = a + 1
        Loop 
""")

# Simple do  with exit
tests.append("""
        a = 0
        Do Until a <10
            g = 10
            a = a + 1
            Exit Do
        Loop 
""")

# Nested do  loop
tests.append("""
        a = 0
        Do Until a< 10
            g = 10
            a = a + 1
            Do While a <10
                doit
            Loop 
        Loop 
""")
# << Parsing tests >> (56 of 60)
# simple type
tests.append("""
Type myType
    A As Integer
    B As String
    C As MyClass.MyType
End Type
""")

# simple type with scope
tests.append("""
Public Type myType
    A As Integer
    B As String
    C As MyClass.MyType
End Type
""")

tests.append("""
Private Type myType
    A As Integer
    B As String
    C As MyClass.MyType
End Type
""")

# With a comment inside
tests.append("""
Private Type myType
    A As Integer
    B As String
    ' Here is a comment
    C As MyClass.MyType
End Type
""")
# << Parsing tests >> (57 of 60)
# General with with just the structure
tests.append("""
With MyObject
    a = 10
End With
""")

# General with with some assignments
tests.append("""
With MyObject
    .value = 10
    .other = "Hello"
End With
""")

# General with with some assignments and expressions
tests.append("""
With MyObject
    .value = .other + 10
    .other = "Hello" & .name
End With
""")

# Nested With
tests.append("""
With MyObject
    a = 10
    With .OtherObject
        b = 20
    End With
End With
""")

# General with with just the structure and labels
tests.append("""
1: With MyObject
2: 	a = 10
3: End With
""")

# Empty with
tests.append("""
With MyObject
End With
""")
# << Parsing tests >> (58 of 60)
# Simple header found at the top of most class files
tests.append("""
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
""")
# << Parsing tests >> (59 of 60)
# Simple enumeration
tests.append("""
Enum MyEnum
    one
    two
    three
    four
    five
End Enum
""")


# Scoped enumeration
tests.append("""
Public Enum MyEnum
    one
    two
    three
    four
    five
End Enum
""")

tests.append("""
Private Enum MyEnum
    one
    two
    three
    four
    five
End Enum
""")

# Simple enumeration with comments
tests.append("""
Enum MyEnum ' yeah
    one ' this 
    two ' is 
    three
    four ' neat
    five
End Enum
""")
# << Parsing tests >> (60 of 60)
failures = [
        "If a = 10 Then d = 1 Else If k = 12 Then b = 12",
        "If a = 10 Then d = 1 Else If k = 12 Then b = 12 Else g=123",
]
# -- end -- << Parsing tests >>

class ParsingTest(TestCase):
    """Holder class which gets built into a whole test case"""


def getTestMethod(vb):
    """Create a test method"""
    def testMethod(self):
        try:
            buildParseTree(vb)
        except VBParserError:
            raise "Unable to parse ...\n%s" % vb
    return testMethod

#
# Add tests to main test class
for idx in range(len(tests)):
    setattr(ParsingTest, "test%d" % idx, getTestMethod(tests[idx]))


if __name__ == "__main__":
    main()
