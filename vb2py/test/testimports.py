import vb2py.vbparser
import unittest

class TestImports(unittest.TestCase):
    # << Imports tests >>
    def testImportClassToModule(self):
        """Import from class to module"""
        self.proj = vb2py.vbparser.VBProject()
        self.utils = vb2py.vbparser.VBCodeModule(modulename="utils")
        self.cls = vb2py.vbparser.VBClassModule(modulename="Cls", classname="Cls")
        #
        self.utils.assignParent(self.proj)
        self.cls.assignParent(self.proj)
        #
        utils = vb2py.vbparser.parseVB("""
        Public Function Fact(x)
            Dim c As New Cls
        End Function
        """, container=self.utils)
        #
        cls = vb2py.vbparser.parseVB("""
        Public A
        """, container=self.cls)
        #
        utils_code = self.utils.renderAsCode()
        self.assertNotEqual(utils_code.find("import Cls"), -1)
    # -- end -- << Imports tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff

if __name__ == "__main__":
    unittest.main()
