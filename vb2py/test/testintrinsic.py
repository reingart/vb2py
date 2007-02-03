# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

from unittest import *
from testframework import *

# << Intrinsic tests >> (1 of 10)
# Concatenation
tests.append(("""
a = "hello"
b = "there"
c = a & b
""", {"a" : "hello", "b" : "there", "c" : "hellothere"}))

# Lots of Concatenation
tests.append(("""
a = "hello"
b = "there"
c = a & a & a & b & b & b
""", {"a" : "hello", "b" : "there", "c" : "hellohellohellotheretherethere"}))

# Left
tests.append(("""
a = "hello"
b = Left(a, 1)
c = Left(a, 2)
d = Left(a, 10)
""", {"a" : "hello", "b" : "h", "c" : "he", "d" : "hello"}))

# Right
tests.append(("""
a = "hello"
b = Right(a, 1)
c = Right(a, 2)
d = Right(a, 10)
""", {"a" : "hello", "b" : "o", "c" : "lo", "d" : "hello"}))

# Mid with one parameter
tests.append(("""
a = "hellothere"
b = Mid(a, 1)
c = Mid(a, 2)
d = Mid(a, 5)
""", {"a" : "hellothere", "b" : "hellothere", "c" : "ellothere", "d" : "othere"}))

# Mid with two parameters
tests.append(("""
a = "hellothere"
b = Mid(a, 1, 3)
c = Mid(a, 2, 4)
d = Mid(a, 5, 20)
""", {"a" : "hellothere", "b" : "hel", "c" : "ello", "d" : "othere"}))
# << Intrinsic tests >> (2 of 10)
tests.extend([
	('a = InStr("hello", "ll")', {"a" : 3}),
	('a = InStr("hello", "lll")', {"a" : 0}),
	('a = InStr(4, "hellollo", "ll")', {"a" : 6}),
	('a = InStr(4, "hellollo", "lll")', {"a" : 0}),

# InstrB ??

	('a = Len("hello")', {"a" : 5}),
	('a = Len("")', {"a" : 0}),

	('a = LCase("hello")', {"a" : "hello"}),
	('a = LCase("HELlo")', {"a" : "hello"}),
	('a = LCase("HELLO")', {"a" : "hello"}),

	('a = UCase("hello")', {"a" : "HELLO"}),
	('a = UCase("HELlo")', {"a" : "HELLO"}),
	('a = UCase("HELLO")', {"a" : "HELLO"}),

	('a = Space(4)', {"a" : "    "}),
	('a = Space("4")', {"a" : "    "}),
	('a = Space("0")', {"a" : ""}),

	('a = StrComp("one", "two")', {"a" : -1}),
	('a = StrComp("two", "two")', {"a" : 0}),
	('a = StrComp("two", "one")', {"a" : 1}),

	('a = String(4, "a")', {"a" : "aaaa"}),
	('a = String(0, "a")', {"a" : ""}),
	('a = String(4, "abc")', {"a" : "aaaa"}),

	('a = LTrim("  hello there   ")', {"a" : "hello there   "}),
	('a = LTrim("hello there   ")', {"a" : "hello there   "}),

	('a = RTrim("  hello there   ")', {"a" : "  hello there"}),
	('a = RTrim("  hello there")', {"a" : "  hello there"}),

	('a = Trim("  hello there   ")', {"a" : "hello there"}),
	('a = Trim("hello there")', {"a" : "hello there"}),

	('a = Trim(1234)', {"a" : "1234"}),
	('a = LTrim(1234)', {"a" : "1234"}),
	('a = RTrim(1234)', {"a" : "1234"}),

	('a = IsNumeric("nope")', {"a" : 0}),
	('a = IsNumeric("123nope")', {"a" : 0}),
	('a = IsNumeric("123.0nope")', {"a" : 0}),
	('a = IsNumeric("123.0ne10 ope")', {"a" : 0}),
	('a = IsNumeric("1")', {"a" : 1}),
	('a = IsNumeric("-1")', {"a" : 1}),
	('a = IsNumeric("-12.45")', {"a" : 1}),
	('a = IsNumeric("-12.45e5")', {"a" : 1}),
	('a = IsNumeric("-12.45e-5")', {"a" : 1}),
	('a = IsNumeric("12.45")', {"a" : 1}),
	('a = IsNumeric("12.45e5")', {"a" : 1}),
	('a = IsNumeric("12.45e-5")', {"a" : 1}),
	('a = IsNumeric("+12.45")', {"a" : 1}),
	('a = IsNumeric("+12.45e5")', {"a" : 1}),
	('a = IsNumeric("+12.45e-5")', {"a" : 1}),
])
# << Intrinsic tests >> (3 of 10)
tests.extend([
	('a = Asc("a")', {"a" : 97}),
	('a = AscB("a")', {"a" : 97}),
	('a = AscW("a")', {"a" : 97}),

	('a = Abs(101)', {"a" : 101}),
	('a = Abs(-101)', {"a" : 101}),

	('a = Chr(97)', {"a" : "a"}),
	('a = ChrB(97)', {"a" : "a"}),
	('a = ChrW(97)', {"a" : "a"}),

# CDATE?

])
# << Intrinsic tests >> (4 of 10)
tests.extend([
	('a = CBool(-1)', {"a" : 1}),
	('a = CBool(0)', {"a" : 0}),
	('a = CBool(-1)', {"a" : 1}),

	('a = CByte(-1)', {"FAIL" : 1}),
	('a = CByte(67.4)', {"a" : 67}),
	('a = CByte("123.8")', {"a" : 124}),
	('a = CByte("1023")', {"FAIL" : 1}),
	('a = CByte("1ggg023")', {"FAIL" : 1}),

	('a = CDbl(-1)', {"a" : -1}),
	('a = CDbl(67.3)', {"a" : 67.3}),
	('a = CDbl("123")', {"a" : 123}),
	('a = CDbl("1023.1")', {"a" : 1023.1}),
	('a = CDbl("1023.8")', {"a" : 1023.8}),
	('a = CDbl("1ggg023")', {"FAIL" : 1}),

	('a = CInt(-1)', {"a" : -1}),
	('a = CInt(67.1)', {"a" : 67}),
	('a = CInt("123.3")', {"a" : 123}),
	('a = CInt("1023.8")', {"a" : 1024}),
	('a = CInt("-331023")', {"FAIL" : 1}),
	('a = CInt("331023")', {"FAIL" : 1}),
	('a = CInt("1ggg023")', {"FAIL" : 1}),

	('a = CLng(-1)', {"a" : -1}),
	('a = CLng(67.2)', {"a" : 67}),
	('a = CLng("123")', {"a" : 123}),
	('a = CLng("1023.1")', {"a" : 1023}),
	('a = CLng("-331023")', {"a" : -331023}),
	('a = CLng("331023.8")', {"a" : 331024}),
	('a = CLng("1ggg023")', {"FAIL" : 1}),

	('a = CSng(-1)', {"a" : -1}),
	('a = CSng(67.3)', {"a" : 67.3}),
	('a = CSng("123")', {"a" : 123}),
	('a = CSng("1023.1")', {"a" : 1023.1}),
	('a = CSng("1023.8")', {"a" : 1023.8}),
	('a = CSng("1ggg023")', {"FAIL" : 1}),

	('a = CStr(-1)', {"a" : "-1"}),
	('a = CStr("hello")', {"a" : "hello"}),

])
# << Intrinsic tests >> (5 of 10)
tests.extend([
	('a = Hex(255)', {"a" : "FF"}),
	('a = Hex("255")', {"a" : "FF"}),
	('a = Hex(0)', {"a" : "0"}),
	('a = Hex(12345)', {"a" : "3039"}),

	('a = Oct(255)', {"a" : "377"}),
	('a = Oct("255")', {"a" : "377"}),
	('a = Oct(0)', {"a" : "0"}),
	('a = Oct(12345)', {"a" : "30071"}),

	('a = Fix(-1)', {"a" : -1}),
	('a = Fix(67.1)', {"a" : 67}),
	('a = Fix("123.3")', {"a" : 123}),
	('a = Fix("1023.8")', {"a" : 1023}),

	('a = Int(-1)', {"a" : -1}),
	('a = Int(67.1)', {"a" : 67}),
	('a = Int("123.3")', {"a" : 123}),
	('a = Int("1023.8")', {"a" : 1023}),

	('a = Sgn(-1)', {"a" : -1}),
	('a = Sgn(0)', {"a" : 0}),
	('a = Sgn(1)', {"a" : 1}),
	('a = Sgn("-10")', {"a" : -1}),
	('a = Sgn("10")', {"a" : 1}),
])
# << Intrinsic tests >> (6 of 10)
tests.extend([
	('a = Sin(0)', {"a" : 0}),
	('a = Sin("0")', {"a" : 0}),

	('a = Cos(0)', {"a" : 1}),
	('a = Cos("0")', {"a" : 1}),

	('a = Tan(0)', {"a" : 0}),
	('a = Tan("0")', {"a" : 0}),

	('a = Int(10*Atn(10))', {"a" : 14}),
	('a = Int(10*Atn("10"))', {"a" : 14}),

	('a = Exp(0)', {"a" : 1}),
	('a = Exp("0")', {"a" : 1}),

	('a = Int(Log(10))', {"a" : 2}),
	('a = Int(Log("10"))', {"a" : 2}),

	('a = Sqr(16)', {"a" : 4}),
	('a = Sqr("16")', {"a" : 4}),

	('a = Round(1.1)', {"a" : 1}),
	('a = Round(1.6)', {"a" : 2}),

	('a = Round(1.1, 0)', {"a" : 1}),
	('a = Round(1.6, 0)', {"a" : 2}),

	('a = Round(1.11, 1)', {"a" : 1.1}),
	('a = Round(1.16, 1)', {"a" : 1.2}),

	('a = Round(-1.1)', {"a" : -1}),
	('a = Round(-1.6)', {"a" : -2}),

	('a = Round(-1.1, 0)', {"a" : -1}),
	('a = Round(-1.6, 0)', {"a" : -2}),

	('a = Round(-1.11, 1)', {"a" : -1.1}),
	('a = Round(-1.16, 1)', {"a" : -1.2}),
])
# << Intrinsic tests >> (7 of 10)
tests.extend([
("""
Dim _A
b = UBound(_A)
""", {"FAIL" : 1}),

("""
Dim _A(10)
b = UBound(_A)
""", {"b" : 10}),

("""
Dim _A(10, 20)
b = UBound(_A)
c = UBound(_A, 1)
d = UBound(_A, 2)
""", {"b" : 10, "c" : 10, "d" : 20}),

("""
Dim _A(10, 20, 30)
b = UBound(_A)
c = UBound(_A, 1)
d = UBound(_A, 2)
e = UBound(_A, 3)
""", {"b" : 10, "c" : 10, "d" : 20, "e" : 30}),

("""
Dim _A(5 To 10, 10 To 20, 30)
b = UBound(_A)
c = UBound(_A, 1)
d = UBound(_A, 2)
e = UBound(_A, 3)
""", {"b" : 10, "c" : 10, "d" : 20, "e" : 30}),
])
# << Intrinsic tests >> (8 of 10)
tests.extend([
("""
Dim _A
b = LBound(_A)
""", {"FAIL" : 1}),

("""
Dim _A(10)
b = LBound(_A)
""", {"b" : 0}),

("""
Dim _A(10, 20)
b = LBound(_A)
c = LBound(_A, 1)
d = LBound(_A, 2)
""", {"b" : 0, "c" : 0, "d" : 0}),

("""
Dim _A(10, 20, 30)
b = LBound(_A)
c = LBound(_A, 1)
d = LBound(_A, 2)
e = LBound(_A, 3)
""", {"b" : 0, "c" : 0, "d" : 0, "e" : 0}),

("""
Dim _A(5 To 10, 10 To 20, 30)
b = LBound(_A)
c = LBound(_A, 1)
d = LBound(_A, 2)
e = LBound(_A, 3)
""", {"b" : 5, "c" : 5, "d" : 10, "e" : 0}),
])
# << Intrinsic tests >> (9 of 10)
tests.extend([
	('a = Val("12")', {"a" : 12}),
	('a = Val("12.")', {"a" : 12}),
	('a = Val("-12.")', {"a" : -12}),
	('a = Val("-12")', {"a" : -12}),
	('a = Val("-12.5")', {"a" : -12.5}),
	('a = Val("12.5")', {"a" : 12.5}),
	('a = Val("12.5e5")', {"a" : 12.5e5}),
	('a = Val("-12.5e5")', {"a" : -12.5e5}),
	('a = Val("-12.5e-5")', {"a" : -12.5e-5}),
	('a = Val("12.5e-5")', {"a" : 12.5e-5}),
	('a = Val("12.5e-5 mdmdmf")', {"a" : 12.5e-5}),
	('a = Val("12 mdmdmf")', {"a" : 12}),
	('a = Val("12+mdmdmf")', {"a" : 12}),
	('a = Val("12-mdmdmf")', {"a" : 12}),
	('a = Val(" 12-mdmdmf")', {"a" : 12}),
	('a = Val("ccc 12-mdmdmf")', {"a" : 0}),
])
# << Intrinsic tests >> (10 of 10)
tests.extend([
("""
_a = Array(1,2,3,4)
l = UBound(_a)
a1 = _a(0)
a3 = _a(3)
""", {
   "l" : 3,
   "a1" : 1,
   "a3" : 4,
}),

("""
_a = Array("1","2","3","4")
l = UBound(_a)
a1 = _a(0)
a3 = _a(3)
""", {
   "l" : 3,
   "a1" : "1",
   "a3" : "4",
}),

("""
_a = Array(Array(10,20,30),"2","3","4")
l = UBound(_a)
a1 = _a(0)(0)
a3 = _a(3)
""", {
   "l" : 3,
   "a1" : 10,
   "a3" : "4",
}),

("""
_a = Array()
l = UBound(_a)
""", {
   "l" : -1,
}),

])
# -- end -- << Intrinsic tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
	main()
