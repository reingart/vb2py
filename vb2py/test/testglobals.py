import vb2py.vbparser
import unittest

class TestGlobals(unittest.TestCase):
    # << Globals tests >> (1 of 8)
    def setUp(self):
        """Set up our tests"""
        self.proj = vb2py.vbparser.VBProject()
        self.utils = vb2py.vbparser.VBCodeModule(modulename="utils")
        self.base = vb2py.vbparser.VBCodeModule(modulename="base")
        self.cls = vb2py.vbparser.VBClassModule(modulename="cls", classname="Cls")
        #
        self.utils.parent = self.base.parent = self.cls.parent = self.proj        
        #
        self.utils = vb2py.vbparser.parseVB("""
        Public x, y, z
        Private one, two, three

        Public Function Fact(x)
            Fact=1
        End Function
        """, container=self.utils)
        #
        self.base = vb2py.vbparser.parseVB("""
        Public a, b, c
        Private four, five, six

        Public Sub DoIt(a, b, x, y, h)
            x = 1
            y = 1
        End Sub
        """, container=self.base)
        #
    # << Globals tests >> (2 of 8)
    def testSimpleOneLookup(self):
        """testSimpleOneLookup: simple lookup of globals from one location"""
        py = vb2py.vbparser.parseVB("""
        Sub Run()
            m = x
            n = y
            o = z
        End Sub
        """, container=self.cls)
        #
        python = py.renderAsCode()
        #
        for converted in ("utils.x", "utils.y", "utils.z"):
            self.assertNotEqual(python.find(converted), -1, python)

    def testSimpleTwoLookups(self):
        """testSimpleTwoLookups: simple lookup of globals from two locations"""
        py = vb2py.vbparser.parseVB("""
        Sub Run()
            a = x
            b = y
            c = z
        End Sub
        """, container=self.cls)
        #
        python = py.renderAsCode()
        #
        for converted in ("utils.x", "utils.y", "utils.z",
                          "base.a", "base.b", "base.c"):
            self.assertNotEqual(python.find(converted), -1, python)
    # << Globals tests >> (3 of 8)
    def testSimpleNoLookup(self):
        """testSimpleNoLookup: don't lookup private variables"""
        py = vb2py.vbparser.parseVB("""
        Sub Run()
            one = two
            three = four
            five = six
        End Sub
        """, container=self.cls)
        #
        python = py.renderAsCode()
        #
        for converted in (".one", ".two", ".three", ".four", ".five", ".six"):
            self.assertEqual(python.find(converted), -1, python)
    # << Globals tests >> (4 of 8)
    def testLocalsShadowGlobals(self):
        """testLocalsShadowGlobals: locals will shadow globals"""
        py = vb2py.vbparser.parseVB("""
        Sub Run()
            Dim a, b, x, y
            a = x
            b = y
            c = z
        End Sub
        """, container=self.cls)
        #
        python = py.renderAsCode()
        #
        for converted in ("utils.z", "base.c"):
            self.assertNotEqual(python.find(converted), -1, python)
        for converted in ("utils.x", "utils.y", "base.a", "base.b"):
            self.assertEqual(python.find(converted), -1, python)

    def testModuleLocalsShadowGlobals(self):
        """testModuleLocalsShadowGlobals: module locals will shadow globals"""
        py = vb2py.vbparser.parseVB("""
        Dim a, b, x, y

        Sub Run()
            a = x
            b = y
            c = z
        End Sub
        """, container=self.cls)
        #
        python = py.renderAsCode()
        #
        for converted in ("utils.z", "base.c"):
            self.assertNotEqual(python.find(converted), -1, python)
        for converted in ("utils.x", "utils.y", "base.a", "base.b"):
            self.assertEqual(python.find(converted), -1, python)

    def testModuleDefsShadowGlobals(self):
        """testModuleDefsShadowGlobals: module definitions will shadow globals"""
        py = vb2py.vbparser.parseVB("""
        Sub Run()
            a = x
            b = y
            c = z
        End Sub

        Function a()
        End Function
        Function b()
        End Function
        Function x()
        End Function
        Function y()
        End Function

        """, container=self.cls)
        #
        python = py.renderAsCode()
        #
        for converted in ("utils.z", "base.c"):
            self.assertNotEqual(python.find(converted), -1, python)
        for converted in ("utils.x", "utils.y", "base.a", "base.b"):
            self.assertEqual(python.find(converted), -1, python)
    # << Globals tests >> (5 of 8)
    def testParametersShadowGlobals(self):
        """testParametersShadowGlobals: parameters will shadow globals"""
        py = vb2py.vbparser.parseVB("""
        Sub Run(a, b, x, y)
            a = x
            b = y
            c = z
        End Sub
        """, container=self.cls)
        #
        python = py.renderAsCode()
        #
        for converted in ("utils.z", "base.c"):
            self.assertNotEqual(python.find(converted), -1, python)
        for converted in ("utils.x", "utils.y", "base.a", "base.b"):
            self.assertEqual(python.find(converted), -1, python)
    # << Globals tests >> (6 of 8)
    def testPropertyShadowGlobals(self):
        """testPropertyShadowGlobals: properties will shadow globals"""
        py = vb2py.vbparser.parseVB("""
        Sub Run()
            a = x
            b = y
            c = z
        End Sub

        Property Get a()
        End Property
        Property Get b()
        End Property
        Property Get x()
        End Property
        Property Get y()
        End Property
        """, container=self.cls)
        #
        python = py.renderAsCode()
        #
        for converted in ("utils.z", "base.c"):
            self.assertNotEqual(python.find(converted), -1, python)
        for converted in ("utils.x", "utils.y", "base.a", "base.b"):
            self.assertEqual(python.find(converted), -1, python)
    # << Globals tests >> (7 of 8)
    def testSubGlobals(self):
        """testSubGlobals: lookup of a global subroutine"""
        py = vb2py.vbparser.parseVB("""
        Sub Run()
            DoIt
        End Sub
        """, container=self.cls)
        #
        python = py.renderAsCode()
        #
        for converted in ("base.DoIt()",):
            self.assertNotEqual(python.find(converted), -1, python)

    def testFnGlobals(self):
        """testFnGlobals: lookup of a global function"""
        py = vb2py.vbparser.parseVB("""
        Sub Run()
            a = Fact(10)
        End Sub
        """, container=self.cls)
        #
        python = py.renderAsCode()
        #
        for converted in ("utils.Fact(10)",):
            self.assertNotEqual(python.find(converted), -1, python)
    # << Globals tests >> (8 of 8)
    def testGlobalButLocalHere(self):
        """testGlobalButLocalHere: lookup of a global which happens to be local here"""
        mymod = vb2py.vbparser.VBCodeModule(modulename="this")
        mymod.parent = self.proj
        py = vb2py.vbparser.parseVB("""
        Public myval
        Sub Run()
            a=myval
        End Sub
        """, container=mymod)
        #
        python = py.renderAsCode()
        #
        for converted in ("this.myval",):
            self.assertEqual(python.find(converted), -1, python)

    def testGlobalButLocalHereNoNeedForGlobal(self):
        """testGlobalButLocalHereNoNeedForGlobal: lookup of a global which happens to be local here"""
        mymod = vb2py.vbparser.VBCodeModule(modulename="this")
        mymod.parent = self.proj
        py = vb2py.vbparser.parseVB("""
        Dim myval
        Public Function MyRun()
            MyRun=myval
        End Function
        """, container=mymod)
        #
        python = py.renderAsCode()
        #
        for converted in ("this.MyRun", "global MyRun"):
            self.assertEqual(python.find(converted), -1, python)		

    def testNeedAGlobalButOnlyUseOne(self):
        """testNeedAGlobalButOnlyUseOne: need to use a global statement but should just use one"""
        mymod = vb2py.vbparser.VBCodeModule(modulename="this")
        mymod.parent = self.proj
        py = vb2py.vbparser.parseVB("""
        Dim myval
        Public Function MyRun()
            myval = 10
            myval = 20
            myval = 30
        End Function
        """, container=mymod)
        #
        python = py.renderAsCode()
        #
        for converted in ("global myval, myval",):
            self.assertEqual(python.find(converted), -1, python)
    # -- end -- << Globals tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff

if __name__ == "__main__":
    unittest.main()
