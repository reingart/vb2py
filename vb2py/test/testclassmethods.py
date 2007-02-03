from unittest import *
from complexframework import *


# << ClassMethod tests >> (1 of 7)
#
# Simple public method
tests.append((
        VBClassModule(),
        """
        Public my_a As String

        Public Sub SetA(Value As Integer)
            my_a = Value
        End Sub
        """,
        ("A = MyClass()\n"
         "A.SetA('hello')\n"
         "assert A.my_a == 'hello', 'A.my_a was (%s)' % (A.my_a,)\n",)
))         

#
# Simple public method with a local variable shadowing a class variable
tests.append((
        VBClassModule(),
        """
        Public my_a As String
        Public my_b As String

        Public Sub SetA(Value As Integer)
            Dim my_b
            my_b = "other"
            my_a = Value + my_b
        End Sub

        Public Sub SetB(Value As Integer)
            my_b = Value
        End Sub
        """,
        ("A = MyClass()\n"
         "A.SetB('thisisb')\n"
         "A.SetA('thisisa')\n"
         "assert A.my_a == 'thisisaother', 'A.my_a was (%s)' % (A.my_a,)\n"
         "assert A.my_b == 'thisisb', 'A.my_b was (%s)' % (A.my_b,)\n",)
))         

#
# Simple public method calling another method
tests.append((
        VBClassModule(),
        """
        Public my_a As String
        Public my_b As String

        Public Sub SetA(Value As Integer)
            SetB Value
            my_a = my_b
        End Sub

        Public Sub SetB(Value As Integer)
            my_b = Value
        End Sub
        """,
        ("A = MyClass()\n"
         "A.SetA('thisisa')\n"
         "assert A.my_a == 'thisisa', 'A.my_a was (%s)' % (A.my_a,)\n"
         "assert A.my_b == 'thisisa', 'A.my_b was (%s)' % (A.my_b,)\n",)
))         


#
# Simple public method with a parameter shadowing a class variable
tests.append((
        VBClassModule(),
        """
        Public my_a As String
        Public my_b As String

        Public Sub SetA(my_b As Integer)
            my_a = my_b
        End Sub

        Public Sub SetB(Value As Integer)
            my_b = Value
        End Sub
        """,
        ("A = MyClass()\n"
         "A.SetB('thisisb')\n"
         "A.SetA('thisisa')\n"
         "assert A.my_a == 'thisisa', 'A.my_a was (%s)' % (A.my_a,)\n"
         "assert A.my_b == 'thisisb', 'A.my_b was (%s)' % (A.my_b,)\n",)
))
# << ClassMethod tests >> (2 of 7)
#
# Simple public function
tests.append((
        VBClassModule(),
        """
        Public lower_bound As Integer

        Public Sub setLowerBound(Value As Integer)
            lower_bound = Value
        End Sub

        Public Function Factorial(Value As Integer)
            If Value = lower_bound Then
                Factorial = 1
            Else
                Factorial = Value * Factorial(Value-1)
            End If
        End Function
        """,
        ("A = MyClass()\n"
         "A.setLowerBound(1)\n"
         "assert A.Factorial(6) == 720, 'A.Factorial(6) was (%s)' % (A.Factorial(6),)\n",)
))
# << ClassMethod tests >> (3 of 7)
#
# Simple public function
tests.append((
        VBClassModule(),
        """
        Public a1, a2, a3, a4
        Public Function add(Optional X=10, Optional Y=20, Optional Z=30)
            add = X + Y + Z
        End Function

        Public Sub set()
            a1 = add(1,2,3)
            a2 = add(,2,3)
            a3 = add(,,3)
            a4 = add()
        End Sub

        """,
        ("A = MyClass()\n",
         "A.set()\n",
         "assert A.a1 == 6\n",
         "assert A.a2 == 15\n",
         "assert A.a3 == 33\n",
         "assert A.a4 == 60\n",
         )
))
# << ClassMethod tests >> (4 of 7)
#
# Simple private method
tests.append((
        VBClassModule(),
        """
        Public my_a As String

        Private Sub SetA(Value As Integer)
            my_a = Value
        End Sub
        """,
        ("A = MyClass()\n"
         "try:\n"
         "  A.SetA('hello')\n"
         "except AttributeError:\n"
         "  pass\n"
         "else:\n"
         "  assert 0, 'Method should be private'\n",)
))
# << ClassMethod tests >> (5 of 7)
#
# Simple init method called automatically
tests.append((
        VBClassModule(),
        """
        Public my_a As String

        Public Sub Class_Initialize()
            my_a = "hello"
        End Sub

        Public Sub SetA(Value As Integer)
            my_a = Value
        End Sub
        """,
        ("A = MyClass()\n"
         "assert A.my_a == 'hello', 'A.my_a was (%s)' % (A.my_a,)\n"
         "A.SetA('bye')\n"
         "assert A.my_a == 'bye', 'A.my_a was (%s)' % (A.my_a,)\n",)
))         


#
# Explicitely calling the init method
tests.append((
        VBClassModule(),
        """
        Public my_a As String

        Public Sub Class_Initialize()
            my_a = "hello"
        End Sub

        Public Sub ReInit()
            Class_Initialize
        End Sub

        Public Sub SetA(Value As Integer)
            my_a = Value
        End Sub

        """,
        ("A = MyClass()\n"
         "A.SetA('bye')\n"
         "A.ReInit()\n"
         "assert A.my_a == 'hello', 'A.my_a was (%s)' % (A.my_a,)\n",)
))         

#
# Explicitely calling the terminate method
tests.append((
        VBClassModule(),
        """
        Public my_a As String

        Public Sub Class_Terminate()
            my_a = "hello"
        End Sub

        Public Sub Reset()
            Class_Terminate
        End Sub

        Public Sub SetA(Value As Integer)
            my_a = Value
        End Sub

        """,
        ("A = MyClass()\n"
         "A.SetA('bye')\n"
         "A.Reset()\n"
         "assert A.my_a == 'hello', 'A.my_a was (%s)' % (A.my_a,)\n",)
))         


#
# init method is private
tests.append((
        VBClassModule(),
        """
        Public my_a As String

        Sub Class_Initialize()
            my_a = "hello"
        End Sub

        Public Sub ReInit()
            Class_Initialize
        End Sub

        Public Sub SetA(Value As Integer)
            my_a = Value
        End Sub

        """,
        ("A = MyClass()\n"
         "A.SetA('bye')\n"
         "A.ReInit()\n"
         "assert A.my_a == 'hello', 'A.my_a was (%s)' % (A.my_a,)\n",)
))         


#
# Terminate method is private
tests.append((
        VBClassModule(),
        """
        Public my_a As String

        Sub Class_Terminate()
            my_a = "hello"
        End Sub

        Public Sub Reset()
            Class_Terminate
        End Sub

        Public Sub SetA(Value As Integer)
            my_a = Value
        End Sub

        """,
        ("A = MyClass()\n"
         "A.SetA('bye')\n"
         "A.Reset()\n"
         "assert A.my_a == 'hello', 'A.my_a was (%s)' % (A.my_a,)\n",)
))         

tests.append((
        VBClassModule(),
        """
        Public my_a As String

        Sub Class_Terminate()
            'my_a = 1/0
        End Sub

        Public Sub SetA(Value As Integer)
            my_a = Value
        End Sub

        """,
        ("$assert python.find('def __del__(self') <> -1, '__del__ method not created'", )
))
# << ClassMethod tests >> (6 of 7)
#
# Class properties
tests.append((
        VBClassModule(),
        """
        Public arr()

        Public Sub DoIt(Value As Integer)
            ReDim arr(Value)
        End Sub
        """,
        ("A = MyClass()\n"
         "B = MyClass()\n"
         "A.DoIt(10)\n"
         "B.DoIt(20)\n"
         "assert len(A.arr) == 11, 'len(A.arr) was (%s)' % (len(A.arr),)\n"
         "assert len(B.arr) == 21, 'len(B.arr) was (%s)' % (len(B.arr),)\n",)
))         

#
# Make sure class properties are not shared
tests.append((
        VBClassModule(),
        """
        Public arr(20)

        Public Sub DoIt(Value As Integer)
            arr(10) = Value
        End Sub
        """,
        ("A = MyClass()\n"
         "B = MyClass()\n"
         "A.DoIt(10)\n"
         "B.DoIt(20)\n"
         "assert A.arr[10] == 10, 'A.arr[10] was (%s)' % (A.arr[10],)\n"
         "assert B.arr[10] == 20, 'B.arr[10] was (%s)' % (B.arr[10],)\n",)
))
# << ClassMethod tests >> (7 of 7)
#
# Me in an expression
tests.append((
        VBClassModule(),
        """
        Public Val

        Public Sub DoIt(Value As Integer)
            Me.Val = Value
        End Sub
        """,
        ("A = MyClass()\n"
         "A.DoIt(10)\n"
         "assert A.Val==10, 'A.Val was (%s)' % (A.Val,)\n",)
))         


#
# Me in a call
tests.append((
        VBClassModule(),
        """
        Public Val

        Public Sub DoIt(Value As Integer)
            Val = Value
            Me.AddOne
        End Sub

        Public Sub AddOne()
            Val = Val + 1
        End Sub

        """,
        ("A = MyClass()\n"
         "A.DoIt(10)\n"
         "assert A.Val==11, 'A.Val was (%s)' % (A.Val,)\n",)
))
# -- end -- << ClassMethod tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
    main()
