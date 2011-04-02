from testframework import *

# << Intrinsic tests >> (1 of 10)
tests.extend([
("""
_a = Array(1,2,3,4)
t = IsArray(_a)
f = IsArray(t)
""", {
   "t" : 1,
   "f" : 0,
}),
])
# << Intrinsic tests >> (2 of 10)
tests.extend([
("""
t = IIf(10<20, "true", "false")
f = IIf(10>20, "true", "false")
""", {
   "t" : "true",
   "f" : "false",
}),
])
# << Intrinsic tests >> (3 of 10)
tests.extend([
("""a = StrReverse("")""", { "a" : "",}),
("""a = StrReverse("1")""", { "a" : "1",}),
("""a = StrReverse("1234")""", { "a" : "4321",}),
("""a = StrReverse(1234)""", { "a" : "4321",}),
])
# << Intrinsic tests >> (4 of 10)
tests.extend([
("a = Choose(1)", { "a" : None,}),
("a = Choose(1, 10)", { "a" : 10,}),
("a = Choose(1, 10, 20, 30)", { "a" : 10,}),
("a = Choose(2, 10, 20, 30)", { "a" : 20,}),
("a = Choose(3, 10, 20, 30)", { "a" : 30,}),
("a = Choose(-1, 10, 20, 30)", { "a" : None}),
("a = Choose(20, 10, 20, 30)", { "a" : None}),
])
# << Intrinsic tests >> (5 of 10)
tests.extend([
('a = Join(Array(1,2,3))', { "a" : "1 2 3",}),
('a = Join(Array(1,2,3), ",")', { "a" : "1,2,3",}),
('a = Join(Array(1,2,3), "")', { "a" : "123",}),
('a = Join(Array("1","2","3"))', { "a" : "1 2 3",}),
('a = Join(Array("1","2","3"), ",")', { "a" : "1,2,3",}),
('a = Join(Array("1","2","3"), "")', { "a" : "123",}),
('a = Join(Array("1","2","3"), "  ")', { "a" : "1  2  3",}),
])
# << Intrinsic tests >> (6 of 10)
tests.extend([
('a = Switch()', { "a" : None,}),
('a = Switch(1,0)', { "a" : 0,}),
('a = Switch(0,0,1,10)', { "a" : 10,}),
('a = Switch(0,0,1,10,1,20)', { "a" : 10,}),
('a = Switch(0,0,0,10,1,20)', { "a" : 20,}),
('a = Switch(0,0,0,10,0,20)', { "a" : None,}),
('a = Switch(1>0,0,1>0,10,1>0,20)', { "a" : 0,}),
('a = Switch(1<0,0,1>0,10,1>0,20)', { "a" : 10,}),
('a = Switch(1<0,0,1<0,10,1>0,20)', { "a" : 20,}),
('a = Switch(1<0,0,1<0,10,1<0,20)', { "a" : None,}),
])
# << Intrinsic tests >> (7 of 10)
tests.extend([
("""
_a = Split("ab cd efg hijk")
l = UBound(_a)
a1 = _a(0)
a3 = _a(3)
""", { "l" : 3,
       "a1" : "ab",
       "a3" : "hijk"}),

("""
_a = Split("ab cd efg hijk", " ")
l = UBound(_a)
a1 = _a(0)
a3 = _a(3)
""", { "l" : 3,
       "a1" : "ab",
       "a3" : "hijk"}),

("""
_a = Split("ab cd efg hijk", ",")
l = UBound(_a)
a1 = _a(0)
""", { "l" : 0,
       "a1" : "ab cd efg hijk",}),

("""
_a = Split("ab,cd,efg,hijk", ",")
l = UBound(_a)
a1 = _a(0)
a3 = _a(3)
""", { "l" : 3,
       "a1" : "ab",
       "a3" : "hijk"}),

("""
_a = Split("ab,cd,efg,hijk", ",", 1)
l = UBound(_a)
a1 = _a(0)
""", { "l" : 0,
       "a1" : "ab,cd,efg,hijk",}),

("""
_a = Split("ab,cd,efg,hijk", ",",2)
l = UBound(_a)
a1 = _a(0)
a2 = _a(1)
""", { "l" : 1,
       "a1" : "ab",
       "a2" : "cd,efg,hijk"}),

("""
_a = Split("ab,cd,efg,hijk", ",", 10)
l = UBound(_a)
a1 = _a(0)
a3 = _a(3)
""", { "l" : 3,
       "a1" : "ab",
       "a3" : "hijk"}),

])
# << Intrinsic tests >> (8 of 10)
tests.extend([
("""
a = FileLen("/usr/local/lib/python2.6/dist-packages/vb2py/test/vbtestfile.txt")
""", { "a" : 164}),
])
# << Intrinsic tests >> (9 of 10)
tests.extend([
("""
Open "/usr/local/lib/python2.6/dist-packages/vb2py/test/vbtestfile.txt" For Input As #3
a = Lof(3)
""", { "a" : 164}),
])
# << Intrinsic tests >> (10 of 10)
tests.extend([
("""
_a = Environ("OS")
b = Environ(1000)

For _i = 1 To 100
   If Left$(Environ(_i), 3) = "OS=" Then c = Environ(_i)
Next _i

d = Environ("kskskskss")

""", { "a" : "Windows_NT",
       "b" : "",
       "c" : "OS=Windows_NT",
       "d" : "",
       }),
])
# -- end -- << Intrinsic tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
    main()
