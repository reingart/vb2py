from unittest import *
from complexframework import *

# << Property tests >> (1 of 3)
#
# Simple property Let and Get
tests.append((
        VBClassModule(),
        """
        Public my_a As String
        Public Property Let a(newval As Variant)
            my_a = newval & "_ending"
        End Property

        Public Property Get a() As Variant
            a = my_a
        End Property
        """,
        ("A = MyClass()\n"
         "A.a = 'hello'\n"
         "assert A.a == 'hello_ending', 'A.a was (%s)' % (A.a,)\n",)
))         

#
# Simple property Get and Set
tests.append((
        VBClassModule(),
        """
        Public my_a As String
        Public Property Set a(newval As Variant)
            my_a = newval & "_ending"
        End Property

        Public Property Get a() As Variant
            a = my_a
        End Property
        """,
        ("A = MyClass()\n"
         "A.a = 'hello'\n"
         "assert A.a == 'hello_ending', 'A.a was (%s)' % (A.a,)\n",)
))         

#
# Simple property Get
tests.append((
        VBClassModule(),
        """
        Public my_a As String
        Public Property Get a()
            a = "itworks"
        End Property
        """,
        ("A = MyClass()\n"
         "assert A.a == 'itworks', 'A.a was (%s)' % (A.a,)\n",)
))         

#
# Simple property Let
tests.append((
        VBClassModule(),
        """
        Public my_a As String
        Public Property Let a(newval)
            my_a = newval
        End Property
        """,
        ("A = MyClass()\n"
         "A.a = 'hello'\n"
         "assert A.my_a == 'hello', 'A.my_a was (%s)' % (A.my_a,)\n",)
))
# << Property tests >> (2 of 3)
#
# Multiple properties with internal access
tests.append((
        VBClassModule(),
        """
        Public my_a As String, my_b As String

        Public Property Let a(newval As Variant)
            my_a = newval & "_a"
        End Property

        Public Property Get a() As Variant
            a = my_a
        End Property


        Public Property Let b(newval As Variant)
            my_b = newval & a
        End Property

        Public Property Get b() As Variant
            b = my_b
        End Property
        """,
        ("A = MyClass()\n"
         "A.a = 'hello'\n"
         "A.b = 'there'\n"
         "assert A.a == 'hello_a', 'A.a was (%s)' % (A.a,)\n"
         "assert A.b == 'therehello_a', 'A.b was (%s)' % (A.b,)\n",)
))
# << Property tests >> (3 of 3)
#
# Property set with an exit (this failed once)
tests.append((
        VBClassModule(),
        """
        Public my_a As String
        Public Property Let a(newval As Variant)
            my_a = newval & "_ending"
            Exit Property
        End Property

        Public Property Get a() As Variant
            a = my_a
        End Property
        """,
        ("A = MyClass()\n"
         "A.a = 'hello'\n"
         "assert A.a == 'hello_ending', 'A.a was (%s)' % (A.a,)\n",)
))         

#
# Property get with an exit (this failed once)
tests.append((
        VBClassModule(),
        """
        Public my_a As String
        Public Property Set a(newval As Variant)
            my_a = newval & "_ending"
        End Property

        Public Property Get a() As Variant
            a = my_a
            Exit Property
        End Property
        """,
        ("A = MyClass()\n"
         "A.a = 'hello'\n"
         "assert A.a == 'hello_ending', 'A.a was (%s)' % (A.a,)\n",)
))
# -- end -- << Property tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
    main()
