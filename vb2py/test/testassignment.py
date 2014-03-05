from testframework import *

# << Assignment tests >> (1 of 8)
numeric = [
    ('a=10', 				{'a' : 10}),
    ('b=10.34', 				{'b' : 10.34}),
    ('c=10e5',      {'c' : 10e5}),
    ('d=-1',        {'d' :-1}),
    ('e=-10.765', 		{'e' : -10.765}),
    ('f=-12.4e4', 		{'f' : -12.4e4}),
    ('g=1.0e-45',   {'g' : 1.0e-45}),
    ('h=-1e-8',     {'h' :-1e-8}),
    ('i=&HFF',      {'i' :255}),
    ('j=&HFF&',     {'j' :255}),
]

strings = [
    ('a="a"',       {'a' : "a"}),	
    ('b="abcdef"',  {'b' : "abcdef"}),	
    ("""c="'" """,  {'c' : "'"}),	
    ('d="g\"\"h\"\"j"""', {'d' : 'g"h"j"'}),	
    (r'd="\"', {'d' : '\\'}),	# Trailing single \ is tough
    (r'd="hello\not"', {'d' : r'hello\not'}),
]

tests.extend(numeric)
tests.extend(strings)
# << Assignment tests >> (2 of 8)
numeric = [
    ('Let a=10', 				{'a' : 10}),
    ('Let b=10.34', 				{'b' : 10.34}),
    ('Let c=10e5',      {'c' : 10e5}),
    ('Let d=-1',        {'d' :-1}),
    ('Let e=-10.765', 		{'e' : -10.765}),
    ('Let f=-12.4e4', 		{'f' : -12.4e4}),
    ('Let g=1.0e-45',   {'g' : 1.0e-45}),
    ('Let h=-1e-8',     {'h' :-1e-8}),
    ('Let i=&HFF',      {'i' :255}),
    ('Let j=&HFF&',     {'j' :255}),
]

strings = [
    ('Let a="a"',       {'a' : "a"}),	
    ('Let b="abcdef"',  {'b' : "abcdef"}),	
    ("""Let c="'" """,  {'c' : "'"}),	
    ('Let d="g\"\"h\"\"j"""', {'d' : 'g"h"j"'}),	
    (r'Let d="\"', {'d' : '\\'}),	# Trailing single \ is tough
    (r'Let d="hello\not"', {'d' : r'hello\not'}),
]

tests.extend(numeric)
tests.extend(strings)
# << Assignment tests >> (3 of 8)
numeric_exp = [
    ('a=10+20', 				   {'a' : 30}),
    ('b=10.5+20.5',    {'b' : 31}),
    ('c=(10+20)/6+2', 	{'c' : 7}),
    ('d=(((4*5)/2+10)-10)', 				{'d' : 10}),
    ('e=-(10*10)',     {'e' : -100}),
    ('f=-10*10', 				  {'f' : -100}),
    ('g=&HFF', 				  {'g' : 255}),
    ('h=5^2', 				  {'h' : 25}),
    ('i=10 Mod 2', 			{'i' : 0}),
    ('i=10 Mod 3', 			{'i' : 1}),
    ('i=10 ^ - 1',    {'i' : 0.1}),
]

string_exp = [
    ('a="hello" & "world"', {'a' : "helloworld"}),	
]

tests.extend(numeric_exp)
tests.extend(string_exp)
# << Assignment tests >> (4 of 8)
# Tough to test this one - just create a collection and check it has no length
tests.append(("""
Set _a = New Collection
l = len(_a)
""", {"l" : 0}))

tests.append(("""
Set _a = New Collection
Set _b = _a
l1 = len(_a)
l2 = len(_b)
""", {"l1" : 0, "l2" : 0}))
# << Assignment tests >> (5 of 8)
tests.extend([
    ('a%=10', 				{'a' : 10}),
    ('a&=10', 				{'a' : 10}),
    ('a#=10', 				{'a' : 10}),
    ('a$="10"', 				{'a' : "10"}),
])

tests.extend([
    ('a=10%', 				{'a' : 10}),
    ('a=10#', 				{'a' : 10}),
    ('a=10&', 				{'a' : 10}),
])
# << Assignment tests >> (6 of 8)
tests.extend([
    ('a=0 Or 0', 				{'a' : 0}),
    ('a=1 Or 0', 				{'a' : 1}),
    ('a=0 Or 1', 				{'a' : 1}),
    ('a=1 Or 1', 				{'a' : 1}),
    ('a=0 And 0', 				{'a' : 0}),
    ('a=1 And 0', 				{'a' : 0}),
    ('a=0 And 1', 				{'a' : 0}),
    ('a=1 And 1', 				{'a' : 1}),

    ('a=0 Or Not 0', 				{'a' : 1}),
    ('a=1 Or Not 0', 				{'a' : 1}),
    ('a=0 Or Not 1', 				{'a' : 0}),
    ('a=1 Or Not 1', 				{'a' : 1}),
    ('a=0 And Not 0', 				{'a' : 0}),
    ('a=1 And Not 0', 				{'a' : 1}),
    ('a=0 And Not 1', 				{'a' : 0}),
    ('a=1 And Not 1', 				{'a' : 0}),

    ('a=Not 0 Or 0', 				{'a' : 1}),
    ('a=Not 1 Or 0', 				{'a' : 0}),
    ('a=Not 0 Or 1', 				{'a' : 1}),
    ('a=Not 1 Or 1', 				{'a' : 1}),
    ('a=Not 0 And 0', 				{'a' : 0}),
    ('a=Not 1 And 0', 				{'a' : 0}),
    ('a=Not 0 And 1', 				{'a' : 1}),
    ('a=Not 1 And 1', 				{'a' : 0}),

    ('a=Not 0 Or Not 0', 				{'a' : 1}),
    ('a=Not 1 Or Not 0', 				{'a' : 1}),
    ('a=Not 0 Or Not 1', 				{'a' : 1}),
    ('a=Not 1 Or Not 1', 				{'a' : 0}),
    ('a=Not 0 And Not 0', 				{'a' : 1}),
    ('a=Not 1 And Not 0', 				{'a' : 0}),
    ('a=Not 0 And Not 1', 				{'a' : 0}),
    ('a=Not 1 And Not 1', 				{'a' : 0}),
])

tests.extend([
("""
Dim _a As Object
If _a Is Nothing Then
    b = 1
Else
    b = 2
End If
""", {"b" : 1}),

("""
Dim _a As Object
Set _a = New Collection
If _a Is Nothing Then
    b = 1
Else
    b = 2
End If
""", {"b" : 2}),

("""
Dim _a As Object
Set _a = New Collection
Set _a = Nothing
If _a Is Nothing Then
    b = 1
Else
    b = 2
End If
""", {"b" : 1}),
])
# << Assignment tests >> (7 of 8)
# Lset tests
tests.append(("""
a = "1234"
LSet a = "abcdefgh"
""", {"a" : "abcd"}))

tests.append(("""
a = "1234"
LSet a = "ab"
""", {"a" : "ab  "}))

tests.append(("""
a = "1234"
LSet a = "abcd"
""", {"a" : "abcd"}))

tests.append(("""
a = ""
LSet a = "abcd"
""", {"a" : ""}))

tests.append(("""
a = "1234"
LSet a = ""
""", {"a" : "    "}))

tests.append(("""
a = ""
LSet a = ""
""", {"a" : ""}))
# << Assignment tests >> (8 of 8)
# Lset tests
tests.append(("""
a = "1234"
RSet a = "abcdefgh"
""", {"a" : "abcd"}))

tests.append(("""
a = "1234"
RSet a = "ab"
""", {"a" : "  ab"}))

tests.append(("""
a = "1234"
RSet a = "abcd"
""", {"a" : "abcd"}))

tests.append(("""
a = ""
RSet a = "abcd"
""", {"a" : ""}))

tests.append(("""
a = "1234"
RSet a = ""
""", {"a" : "    "}))

tests.append(("""
a = ""
RSet a = ""
""", {"a" : ""}))
# -- end -- << Assignment tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
    main()
