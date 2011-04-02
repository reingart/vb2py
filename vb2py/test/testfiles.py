from testframework import *
import os 
import vb2py.utils
PATH = vb2py.utils.rootPath()

# << File tests >> (1 of 14)
# Open with Input
tests.append((r"""
Open "%s" For Input As #3
Input #3, a
Input #3, b
Input #3, c, d, e
Input #3, f, g
Close #3
""" % vb2py.utils.relativePath("test/testread.txt"), {'a' : 'Can you hear me now?',
      'b' : 'Can you still hear me now?',
      'c' : 10, 'd' : 20, 'e' : 30,
      'f' : 5, 'g' : "hello",
}))

# Open with Line Input
tests.append((r"""
Open "%s/test/testread.txt" For Input As #3
Line Input #3, a
Line Input #3, b
Line Input #3, c
Line Input #3, d
Close #3
""" % PATH, {'a' : 'Can you hear me now?',
      'b' : 'Can you still hear me now?',
      'c' : '10, 20, 30',
      'd' : '5, "hello"',
}))


# Open and using Input() to get numbers of characters
tests.append((r"""
Open "%s/test/testread.txt" For Input As #3
a = Input(3, #3)
b = Input(1, #3)
c = Input(3, #3)
Close #3
""" % PATH, {'a' : 'Can',
      'b' : ' ',
      'c' : 'you',
}))

# Bug #810964 Input with indexed variable fails 
tests.append((r"""
Open "%s/test/testread.txt" For Input As #3
Dim _a(3) As String
Input #3, _a(1), _a(2), _a(3)
Close #3
a = _a(1)
b = _a(2)
c = _a(3)
""" % PATH, {'a' : 'Can you hear me now?',
      'b' : 'Can you still hear me now?',
      'c' : 10,
}))

# Open with Random access
tests.append((r"""
Open "%s" For Random As #3 Len = 2
' !!!!Dont expect this to work!!!!
Input #3, a
Input #3, b
Input #3, c, d, e
Input #3, f, g
Close #3
""" % vb2py.utils.relativePath("test/testread.txt"), {'a' : 'This wont work!!!!'}))
# << File tests >> (2 of 14)
# Open with print
tests.append((r"""
Open "%s/test/testwrite.txt" For Output As #3
Print #3, 10
Print #3, 20, 30
Print #3, 40, 50
Print #3, "hello"
Close #3
Open "%s/test/testwrite.txt" For Input As #3
Input #3, a, b, c, d, f
Line Input #3, e
""" % (PATH, PATH), {'a' : 10, 'b' : 20,
      'c' : 30, 'd' : 40, 'e' : 'hello',
      'f' : 50,
}))


# Open with print but no cr
tests.append((r"""
Open "%s/test/testwrite.txt" For Output As #3
Print #3, 10;
Print #3, 20, 30;
Print #3, 40, "hello", 50;
Close #3
Open "%s/test/testwrite.txt" For Input As #3
Line Input #3, a
""" % (PATH, PATH), {'a' : "1020\t3040\thello\t50"}
))


# Bare print with no channel number (Bug #805866 - used to fail during render)
tests.append(("Print 10", {}))
# << File tests >> (3 of 14)
# Open with Input
tests.append((r"""
Close
_a = FreeFile
Open "%s" For Input As FreeFile
Open "%s" For Output As FreeFile
_b = FreeFile
Close
_c = FreeFile
a = _a = _b
b = _a = _c
c = _b = _c
d = CStr(_a) & CStr(_b) & CStr(_c) 
""" % (vb2py.utils.relativePath("test/testread.txt"), 
       vb2py.utils.relativePath("test/testwrite.txt")),
{'a':0, 'b':1, 'c':0, 'd': '131',
}))

# Using Reset instead of Close
tests.append((r"""
Reset
_a = FreeFile
Open "%s" For Input As FreeFile
Open "%s" For Output As FreeFile
_b = FreeFile
Reset
_c = FreeFile
a = _a = _b
b = _a = _c
c = _b = _c
d = CStr(_a) & CStr(_b) & CStr(_c) 
""" % (vb2py.utils.relativePath("test/testread.txt"), 
       vb2py.utils.relativePath("test/testwrite.txt")),
{'a':0, 'b':1, 'c':0, 'd': '131',
}))


# Bug #810968 Close #1, #2 ' fails to parse 
tests.append((r"""
Open "%s" For Input As #3
Open "%s" For Output As #4
Close #3, #4
Input #3, a
""" % (vb2py.utils.relativePath("test/testread.txt"), 
       vb2py.utils.relativePath("test/testwrite.txt")),
{'FAIL' : 'yes',
}))
# << File tests >> (4 of 14)
# Seek as a way of moving around in a file
tests.append((r"""
Open "%s" For Input As #3
Input #3, a
Seek #3, 1
Input #3, b
Seek #3, 5
Input #3, c
""" % vb2py.utils.relativePath("test/testread.txt"),
{
'a' : 'Can you hear me now?',
'b' : 'Can you hear me now?',
'c' : 'you hear me now?',
}))


# Seek as a property of the file
tests.append((r"""
Open "%s" For Input As #3
a = Seek(3)
Input #3, _a
b = Seek(3)
Seek #3, 5
c = Seek(3)
""" % vb2py.utils.relativePath("test/testread.txt"),
{
'a' : 1,
'b' : 23,
'c' : 5,
}))
# << File tests >> (5 of 14)
# Dir
tests.append((r"""
a = Dir("test/test*.txt")
b = Dir()
c = Dir()
""",
{
'a' : 'testread.txt',
'b' : 'testwrite.txt',
'c' : '',
}))


# Dir$
tests.append((r"""
a = Dir$("test/test*.txt")
b = Dir$()
c = Dir$()
""",
{
'a' : 'testread.txt',
'b' : 'testwrite.txt',
'c' : '',
}))

# Dir no parenthesis
tests.append((r"""
a = Dir("test/test*.txt")
b = Dir
c = Dir
""",
{
'a' : 'testread.txt',
'b' : 'testwrite.txt',
'c' : '',
}))

# Dir$ no parenthesis
tests.append((r"""
a = Dir$("test/test*.txt")
b = Dir$
c = Dir$
""",
{
'a' : 'testread.txt',
'b' : 'testwrite.txt',
'c' : '',
}))
# << File tests >> (6 of 14)
# Dir
tests.append((r"""
_a = FreeFile
Open "__f1.txt" For Output As #_a
_b = FreeFile
Open "__f2.txt" For Output As #_b
Close #_b
_c = FreeFile
Close #_a
_d = FreeFile
da = _b-_a
db = _c-_a
dd = _d-_a
""",
{
'da' : 1,
'db' : 1,
'dd' : 0,
}))
# << File tests >> (7 of 14)
# Dir
tests.append((r"""
ChDir "%s"
Open "_test1.txt" For Output As #3
Print #3, "in testdir"
Close #3
ChDir "%s"
Open "_test1.txt" For Output As #3
Print #3, "not in testdir"
Close #3
ChDir "%s"
Open "_test1.txt" For Input As #3
Input #3, a
Close #3
ChDir "%s"
Open "_test1.txt" For Input As #3
Input #3, b
Close #3
""" % (vb2py.utils.relativePath("test/testdir"),
       vb2py.utils.relativePath("test"),
       vb2py.utils.relativePath("test/testdir"),
       vb2py.utils.relativePath("test")),
{
'a' : 'in testdir',
'b' : 'not in testdir',
}))
# << File tests >> (8 of 14)
# Dir
tests.append((r"""
Open "_test1.txt" For Output As #3
Print #3, "made file"
Close #3
Kill "_test1.txt"
a = Dir("_test1.txt")
""",
{
'a' : '',
}))
# << File tests >> (9 of 14)
try:
    for name in os.listdir(vb2py.utils.relativePath("test/mytest2")):
        os.remove(os.path.join(vb2py.utils.relativePath("test/mytest2"), name))
except OSError:
    pass

try:
    os.rmdir(vb2py.utils.relativePath("test/mytest2"))
except OSError, err:
    pass

# Dir
tests.append((r"""
MkDir "%s"
Open "%s\test1.txt" For Output As #3
Print #3, "made file"
Close #3
a = 1
""" % (vb2py.utils.relativePath("test/mytest2"),
       vb2py.utils.relativePath("test/mytest2")),
{
'a' : 1,
}))
# << File tests >> (10 of 14)
try:
    for name in os.listdir(vb2py.utils.relativePath("test/mytestdir")):
        os.remove(os.path.join(vb2py.utils.relativePath("test/mytestdir"), name))
except OSError:
    pass

try:
    os.rmdir(vb2py.utils.relativePath("test/mytestdir"))
except OSError:
    pass

# Dir
tests.append((r"""
MkDir "%s"
RmDir "%s"
a = 0
""" % (vb2py.utils.relativePath("test/mytestdir"),
       vb2py.utils.relativePath("test/mytestdir")),
{
'a' : os.path.isdir(vb2py.utils.relativePath("test/mytestdir")),
}))
# << File tests >> (11 of 14)
try:
    os.remove(os.path.join(vb2py.utils.relativePath("test"), "knewname.txt"))
except OSError:
    pass

# Dir
tests.append((r"""
_path = "%s"
Open _path & "\origname.txt" For Output As #3
Close #3
a = Dir(_path & "\origname.txt")
Name _path & "\origname.txt" As _path & "\knewname.txt"
b = Dir(_path & "\origname.txt")
c = Dir(_path & "\knewname.txt")
""" % (vb2py.utils.relativePath("test")),
{
'a' : "origname.txt",
'b' : "",
'c' : "knewname.txt",
}))
# << File tests >> (12 of 14)
try:
    os.remove(os.path.join(vb2py.utils.relativePath("test"), "finalcopy.txt"))
except OSError:
    pass

# Dir
tests.append((r"""
_path = "%s"
Open _path & "\origcopy.txt" For Output As #3
Print #3, "original"
Close #3
a = Dir(_path & "\origcopy.txt")
b = Dir(_path & "\finalcopy.txt")
FileCopy _path & "\origcopy.txt", _path & "\finalcopy.txt"
c = Dir(_path & "\origcopy.txt")
d = Dir(_path & "\finalcopy.txt")
""" % (vb2py.utils.relativePath("test")),
{
'a' : "origcopy.txt",
'b' : "",
'c' : "origcopy.txt",
'd' : "finalcopy.txt",
}))
# << File tests >> (13 of 14)
# Input as a function to get a certain number of characters
tests.append((r"""
Open "%s" For Input As #3
a = Input(3, #3)
b = Input(4, #3)
Close #3

""" % vb2py.utils.relativePath("test/testread.txt"),
{
'a' : 'Can',
'b' : ' you',
}))
# << File tests >> (14 of 14)
# Input as a function to get a certain number of characters
tests.append((r"""
Open "%s" For Input As #3
While Not EOF(#3)
    Input #3, a
End While
Close #3

""" % vb2py.utils.relativePath("test/testread.txt"),
{
'a' : 'hello',
}))
# -- end -- << File tests >>
#< < Small File tests >>


import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)


if __name__ == "__main__":
    main()
