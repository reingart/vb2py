import unittest 
import vb2py.vbparser

# << Test Classes >>
class TestCustomImport(unittest.TestCase): 
    """Tests for the CustomImport class""" 
    # << CustomImport Tests >> (1 of 2)
    def setUp(self): 
        """Create the test fixture"""
    # << CustomImport Tests >> (2 of 2)
    def testComctlLib(self): 
        """testComctlLib: should be able to import ComctlLib"""
        proj = vb2py.vbparser.VBProject()
        module = vb2py.vbparser.VBCodeModule(modulename="test")
        #
        module.parent = proj        
        #
        vb = """
        ' VB2PY-GlobalAdd: CustomIncludes.ComctlLib = comctllib
        Public Function doit()
            Dim x As Node
            y = Node()
        End Function
        """
        #
        c = vb2py.vbparser.parseVB(vb, container=module)
        #
        py = module.renderAsCode()
        #
        self.assertNotEqual(py.find("x = vb2py.custom.comctllib.Node()"), -1,
                            "Failed on X:\n%s\n\n%s\n" % (vb, py))   
        self.assertNotEqual(py.find("y = vb2py.custom.comctllib.Node()"), -1,
                            "Failed on Y:\n%s\n\n%s\n" % (vb, py))
    # -- end -- << CustomImport Tests >>
# -- end -- << Test Classes >> 

if __name__ == "__main__": 
        unittest.main()
