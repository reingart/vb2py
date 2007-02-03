from unittest import *
from testframework import *

# << While tests >> (1 of 5)
# Simple while / wend
tests.append(("""
a = 0
b = 0
While a<=10
    b = b + a
    a = a + 1
Wend
""", {"a" : 11, "b" : 55}))

# Nested While
tests.append(("""
a = 0
b = 0
While a<=10
    c = 0
    While c < 10
        c = c + 1
        b = b + 1
    Wend
    b = b + a
    a = a + 1
Wend
""", {"a" : 11, "b" : 165, "c" : 10}))
# << While tests >> (2 of 5)
# Simple do while loop
tests.append(("""
a = 0
b = 0
Do While a<=10
    b = b + a
    a = a + 1
Loop
""", {"a" : 11, "b" : 55}))

# Simple do while loop with exit
tests.append(("""
a = 1
b = 0
Do While a<=10
    b = b + a
    a = a + 1
    Exit Do
Loop
""", {"a" : 2, "b" : 1}))

# Nested Do While Loop
tests.append(("""
a = 0
b = 0
Do While a<=10
    c = 0
    Do While c < 10
        c = c + 1
        b = b + 1
    Loop
    b = b + a
    a = a + 1
Loop
""", {"a" : 11, "b" : 165, "c" : 10}))

# Nested Do While Loop With inner exit
tests.append(("""
a = 0
b = 0
Do While a<=10
    c = 0
    Do While c < 10
        c = c + 1
        b = b + 1
        Exit Do
    Loop
    b = b + a
    a = a + 1
Loop
""", {"a" : 11, "b" : 66, "c" : 1}))

# Nested Do While Loop With outer exit
tests.append(("""
a = 0
b = 0
Do While a<=10
    c = 0
    Do While c < 10
        c = c + 1
        b = b + 1
    Loop
    b = b + a
    a = a + 1
    Exit Do
Loop
""", {"a" : 1, "b" : 10, "c" : 10}))
# << While tests >> (3 of 5)
# Simple do while loop with exit
tests.append(("""
a = 1
b = 0
Do
    b = b + a
    a = a + 1
    Exit Do
Loop
""", {"a" : 2, "b" : 1}))


# Nested Do While Loop With inner exit
tests.append(("""
a = 0
b = 0
Do
    c = 0
    Do While c < 10
        c = c + 1
        b = b + 1
        Exit Do
    Loop
    b = b + a
    a = a + 1
    Exit Do
Loop
""", {"a" : 1, "b" : 1, "c" : 1}))

# Nested Do While Loop With outer exit
tests.append(("""
a = 0
b = 0
Do
    c = 0
    Do While c < 10
        c = c + 1
        b = b + 1
    Loop
    b = b + a
    a = a + 1
    Exit Do
Loop
""", {"a" : 1, "b" : 10, "c" : 10}))
# << While tests >> (4 of 5)
# Simple do until
tests.append(("""
a = 1
b = 0
Do
    b = b + a
    a = a + 1
Loop Until a > 10
""", {"a" : 11, "b" : 55}))


# Nested Do Until
tests.append(("""
a = 0
b = 0
Do
    c = 0
    Do 
        c = c + 1
        b = b + 1
    Loop Until c > 10
    b = b + a
    a = a + 1
Loop Until a > 10
""", {"a" : 11, "b" : 176, "c" : 11}))
# << While tests >> (5 of 5)
# Simple do until
tests.append(("""
a = 1
b = 0
Do Until a > 10
    b = b + a
    a = a + 1
Loop 
""", {"a" : 11, "b" : 55}))


# Nested Do Until
tests.append(("""
a = 0
b = 0
Do Until a > 10
    c = 0
    Do Until c > 10
        c = c + 1
        b = b + 1
    Loop 
    b = b + a
    a = a + 1
Loop 
""", {"a" : 11, "b" : 176, "c" : 11}))
# -- end -- << While tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
    main()
