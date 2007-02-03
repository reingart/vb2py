from unittest import *
from complexframework import *

# << Private tests >> (1 of 4)
#
# Simple private data is available from inside class
tests.append((
        VBClassModule(),
        """
        Private a As String

        Public Sub SetA(Value)
            a = Value
        End Sub

        Public Function GetA()
           GetA = a
        End Function

        """,
        ("A = MyClass()\n"
         "A.SetA('hello')\n"
         "assert A.GetA() == 'hello', 'A.a was (%s)' % (A.a,)\n",)
))         

#
# Simple private data is not available from outside class
tests.append((
        VBClassModule(),
        """
        Private a As String

        Public Sub SetA(Value)
            a = Value
        End Sub

        Public Function GetA()
           GetA = a
        End Function

        """,
        ("A = MyClass()\n"
         "A.SetA('hello')\n"
         "assert hasattr(A, 'a') == 0, 'Could see attribute a'\n",)
))         

#
# Simple private constant data
tests.append((
        VBClassModule(),
        """
        Private Const b = 10
        Private a

        Public Sub SetA()
            a = b
        End Sub

        Public Function GetA()
           GetA = a
        End Function

        """,
        ("A = MyClass()\n"
         "A.SetA()\n"
         "assert hasattr(A, 'a') == 0, 'Could see attribute a'\n"
         "assert A.GetA() == 10",)
))
# << Private tests >> (2 of 4)
#
# Private sub is available from inside class
tests.append((
        VBClassModule(),
        """
        Private a As String

        Public Sub SetA(Value)
            DoIt Value
        End Sub

        Private Sub DoIt(Value)
            a = Value
        End Sub

        Public Function GetA()
           GetA = a
        End Function

        """,
        ("A = MyClass()\n"
         "A.SetA('hello')\n"
         "assert A.GetA() == 'hello', 'A.a was (%s)' % (A.a,)\n",)
))         

#
# Simple private sub is not available from outside class
tests.append((
        VBClassModule(),
        """
        Private a As String

        Private Sub DoIt(Value)
            a = Value
        End Sub

        """,
        ("A = MyClass()\n"
         "assert hasattr(A, 'DoIt') == 0, 'Could see attribute DoIt'\n",)
))
# << Private tests >> (3 of 4)
#
# Private fn is available from inside class
tests.append((
        VBClassModule(),
        """
        Private a As String

        Public Sub SetA(Value)
            a = Value
        End Sub

        Private Function GetIt()
            GetIt = a
        End Function

        Public Function GetA()
           GetA = GetIt()
        End Function

        """,
        ("A = MyClass()\n"
         "A.SetA('hello')\n"
         "assert A.GetA() == 'hello', 'A.a was (%s)' % (A.a,)\n",)
))         

#
# Simple private fn is not available from outside class
tests.append((
        VBClassModule(),
        """
        Private a As String

        Private Function GetIt()
            GetIt = a
        End Function

        """,
        ("A = MyClass()\n"
         "assert hasattr(A, 'GetIt') == 0, 'Could see attribute GetIt'\n",)
))
# << Private tests >> (4 of 4)
#
# Private property is available from inside class
tests.append((
        VBClassModule(),
        """
        Private ma As String

        Private Property Let a(Value)
            ma = Value
        End Property

        Private Property Get a()
            a = ma
        End Property

        Public Function GetA()
           GetA = a
        End Function

        Public Sub SetA(Value)
          a = Value
        End Sub

        """,
        ("A = MyClass()\n"
         "A.SetA('hello')\n"
         "assert A.GetA() == 'hello', 'A.a was (%s)' % (A.a,)\n",)
))         

#
# Simple private fn is not available from outside class
tests.append((
        VBClassModule(),
        """
        Private ma As String

        Private Property Let a(Value)
            ma = Value
        End Property

        Private Property Get a()
            a = ma
        End Property

        """,
        ("A = MyClass()\n"
         "assert hasattr(A, 'a') == 0, 'Could see attribute a'\n",)
))
# -- end -- << Private tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
    main()
