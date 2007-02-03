from unittest import *
from testframework import *

# << Intrinsic tests >> (1 of 6)
tests.append("""
' VB2PY-Test: s2 = ["", "abc", "def", "g", "a", "fff", "abcdefg"]

s1 = "abcdefg"
a = InStrRev(s1, s2)
""")

tests.append("""
' VB2PY-Test: s2 = ["", "abc", "def", "g", "a", "fff"]

s1 = ""
a = InStrRev(s1, s2)
""")

tests.append("""
' VB2PY-Test: s2 = ["", "abc", "def", "g", "a", "fff", "ab"]

s1 = "ab"
a = InStrRev(s1, s2)
""")

tests.append("""
' VB2PY-Test: s2 = ["", "abcd", "def", "g", "a", "fff", "abcdefg"]

s1 = "abcdabcd"
a = InStrRev(s1, s2)
""")
# << Intrinsic tests >> (2 of 6)
tests.append("""

a = int(Timer/60.0)

""")
# << Intrinsic tests >> (3 of 6)
tests.append("""

a = RGB(0,0,0)
b = RGB(255,255,255)
c = RGB(255,0,0)
d = RGB(0,255,0)
e = RGB(0,0,255)
f = RGB(10,20,30)
g = RGB(1000,1000,1000)

""")

tests.append("""

a = RGB(-10,-10,-10)

""")
# << Intrinsic tests >> (4 of 6)
tests.append("""
' VB2PY-Test: r = [0, 1, 2, 3, 4, 5, 10]

a = Round(0, r)
b = Round(1, r)
c = Round(1234, r)
d = Round(1.234, r)
e = Round(1.2346, r)
f = Round(1234.5678, r)
g = Round(123.456e5, r)
h = Round(-0, r)
i = Round(-1, r)
j = Round(-1234, r)
k = Round(-1.234, r)
l = Round(-1.2345, r)
m = Round(-1234.5678, r)
n = Round(-123.456e5, r)
h = Round(.000123, r)
i = Round(.0012345, r)
j = Round(.1234, r)
k = Round(-.01234, r)
l = Round(1.2345e-5, r)
m = Round(1234.5678e-10, r)
n = Round(123.456e-15, r)

""")
# << Intrinsic tests >> (5 of 6)
tests.append("""

a = Replace("hello", "ll", "xx")
b = Replace("hello", "", "xx")
c = Replace("hello", "hello", "xx")
d = Replace("", "ll", "xx")

e = Replace("hellohello", "ll", "xx")
f = Replace("hellohello", "", "xx")
g = Replace("hellohello", "hello", "xx")
h = Replace("", "ll", "xx")

i = Replace("hellohello", "ll", "xx", 4)
j = Replace("hellohello", "", "xx", 4)
k = Replace("hellohello", "hello", "xx", 4)
l = Replace("", "ll", "xx", 4)

i = Replace("hellohello", "ll", "xx", 1, 1)
j = Replace("hellohello", "", "xx", 1, 1)
k = Replace("hellohello", "hello", "xx", 1, 1)
l = Replace("", "ll", "xx", 1, 1)

""")
# << Intrinsic tests >> (6 of 6)
tests.append("""

a_ = Array("one", "two", "three", "four", "five", "six")

b = Filter(a_, "o")(0)
c = Filter(a_, "t")(0)
d = Filter(a_, "t")(1)
e = Filter(a_, "f")(0)
f = Filter(a_, "f")(1)

g = Filter(a_, "")(0)
h = Filter(a_, "one")(0)
i = Filter(a_, "notthereatall")(0)

""")

tests.append("""

a_ = Array("one", "two", "three", "four", "five", "six")

b = Filter(a_, "o", True)(0)
c = Filter(a_, "t", True)(0)
d = Filter(a_, "t", True)(1)
e = Filter(a_, "f", True)(0)
f = Filter(a_, "f", True)(1)

g = Filter(a_, "", True)(0)
h = Filter(a_, "one", True)(0)
i = Filter(a_, "notthereatall", True)(0)

""")

tests.append("""

a_ = Array("one", "two", "three", "four", "five", "six")

b = Filter(a_, "o", False)(0)
c = Filter(a_, "t", False)(0)
d = Filter(a_, "t", False)(1)
e = Filter(a_, "f", False)(0)
f = Filter(a_, "f", False)(1)


h = Filter(a_, "one", False)(0)
i = Filter(a_, "notthereatall", False)(0)

""")
# -- end -- << Intrinsic tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addScriptTestsTo(BasicTest, tests)

if __name__ == "__main__":
    main()
