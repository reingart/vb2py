#
# Turn off logging in extensions (too loud!)
import vb2py.extensions
import unittest
vb2py.extensions.disableLogging()
import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff

import vb2py.config
import ConfigParser
from vb2py.vbparser import convertVBtoPython
import re

class TestConfig(unittest.TestCase):

    def setUp(self):
        """Set up the test"""
        # << Setup info >>
        self.c = vb2py.config.VB2PYConfig(init=1)

        self.code1 =  """
            If a = 10 Then
                b = 1
            Else
                b = 2
            End If
            """

        self.code2 = """
            Function f()
            f = 10
            End Function
            """			

        self.code3 = """
            Select Case a
            Case 10
            Case 20
            End Select
            """
        # -- end -- << Setup info >>

    # << Config tests >> (1 of 14)
    def testGetconfig(self):
        """testGetConfig: should be able to get config items"""
        for section, name in (("General", "LoadUserPlugins"),
                              ("Functions", "ReturnVariableName"),
                              ("Labels", "IgnoreLabels"),
                             ):
            a = self.c[section, name]
    # << Config tests >> (2 of 14)
    def testGetConfigMissing(self):
        """testGetConfigMissing: should raise an error if no config items"""
        for section, name in (("Generalyy", "LoadUserPlugins"),
                              ("Functionsyy", "ReturnVariableName"),
                              ("Labelsyy", "IgnoreLabels"),
                             ):
            self.assertRaises(ConfigParser.NoSectionError, self.c.__getitem__, (section, name))
    # << Config tests >> (3 of 14)
    def testSetLocalOveride(self):
        """testSetLocalOveride: should be able to overide config items"""
        for section, name in (("General", "LoadUserPlugins"),
                              ("Functions", "ReturnVariableName"),
                              ("Labels", "IgnoreLabels"),
                             ):
            a = self.c[section, name]
            c = "%s_new" % a
            self.c.setLocalOveride(section, name, c)
            b = self.c[section, name]
            self.assertEqual(b, c)
            self.assertNotEqual(a, b)
    # << Config tests >> (4 of 14)
    def testSetLocalOverideDoesntExist(self):
        """testSetLocalOverideDoesntExist: should raise an error if overide non-existant value"""
        for section, name in (("Genekkral", "LoadUserPlugins"),
                              ("Functillons", "ReturnVariableName"),
                              ("Labelkks", "IgnoreLabels"),
                             ):
            self.assertRaises(ConfigParser.NoSectionError, self.c.setLocalOveride, section, name, "ok")

    def testRemoveLocalOverideDoesntExist(self):
        """testRemoveLocalOverideDoesntExist: should raise an error if remove overide non-existant value"""
        for section, name in (("Genekkral", "LoadUserPlugins"),
                              ("Functillons", "ReturnVariableName"),
                              ("Labelkks", "IgnoreLabels"),
                             ):
            self.assertRaises(ConfigParser.NoSectionError, self.c.removeLocalOveride, section, name)
    # << Config tests >> (5 of 14)
    def testRemoveLocalOveride(self):
        """testRemoveLocalOveride: should be able to remove overide of config items"""
        for section, name in (("General", "LoadUserPlugins"),
                              ("Functions", "ReturnVariableName"),
                              ("Labels", "IgnoreLabels"),
                             ):
            a = self.c[section, name]
            c = "%s_new" % a
            self.c.setLocalOveride(section, name, c)
            self.c.removeLocalOveride(section, name)
            b = self.c[section, name]
            self.assertEqual(a, b)
    # << Config tests >> (6 of 14)
    def testSpaceOrTab(self):
        """testSpaceOrTab: should be change between spaces and tabs"""
        self.c.setLocalOveride("General", "IndentAmount", 4)	
        self.c.setLocalOveride("General", "IndentCharacter", "Space")	
        c_space = convertVBtoPython(self.code1)
        self.c.setLocalOveride("General", "IndentAmount", 1)	
        self.c.setLocalOveride("General", "IndentCharacter", "Tab")	
        c_tabs = convertVBtoPython(self.code1)
        #
        # Should be different
        self.assertNotEqual(c_space, c_tabs)
        #
        # But only tabs and spaces
        self.assertEqual(c_space, c_tabs.replace("\t", "    "))
    # << Config tests >> (7 of 14)
    def testIndentAmount(self):
        """testSpaceOrTab: should be change between spaces and tabs"""
        self.c.setLocalOveride("General", "IndentAmount", 4)	
        self.c.setLocalOveride("General", "IndentCharacter", "Space")	
        c_four = convertVBtoPython(self.code1)
        self.c.setLocalOveride("General", "IndentAmount", 8)	
        c_eight = convertVBtoPython(self.code1)
        #
        # Should be different
        self.assertNotEqual(c_four, c_eight)
        #
        # But only by number of spaces
        self.assertEqual(c_four, c_eight.replace("        ", "    "))
    # << Config tests >> (8 of 14)
    def testRespectPrivateStatus(self):
        """testRespectPrivateStatus: should be able to turn off data hiding"""
        self.c.setLocalOveride("General", "RespectPrivateStatus", "Yes")	
        c_on = convertVBtoPython(self.code2, container=vb2py.vbparser.VBClassModule())
        self.c.setLocalOveride("General", "RespectPrivateStatus", "No")	
        c_off = convertVBtoPython(self.code2, container=vb2py.vbparser.VBClassModule())
        #
        # Should be different
        self.assertNotEqual(c_on, c_off, "Option made no difference: '%s'" % c_on)
        #
        # On should have __, off should not
        self.assertNotEqual(-1, c_on.find("__f"), "Yes didn't have __: '%s'" % c_on)
        self.assertEqual(-1, c_off.find("__f"), "No had __: '%s'" % c_off)
    # << Config tests >> (9 of 14)
    def testPrivateDataPrefix(self):
        """testPrivateDataPrefix: should be able to data hiding prefix"""
        self.c.setLocalOveride("General", "RespectPrivateStatus", "Yes")	
        self.c.setLocalOveride("General", "PrivateDataPrefix", "__")	
        c_on = convertVBtoPython(self.code2, container=vb2py.vbparser.VBClassModule())
        self.c.setLocalOveride("General", "PrivateDataPrefix", "m_")	
        c_off = convertVBtoPython(self.code2, container=vb2py.vbparser.VBClassModule())
        #
        # Should be different
        self.assertNotEqual(c_on, c_off, "Option made no difference: '%s'" % c_on)
        #
        # On should have __, off should have m_
        self.assertNotEqual(-1, c_on.find("__f"), "Yes didn't have __: '%s'" % c_on)
        self.assertNotEqual(-1, c_off.find("m_f"), "No didn't have m_: '%s'" % c_off)
    # << Config tests >> (10 of 14)
    def testFunctionVariable(self):
        """testFunctionVariable: should be able to change function variable"""
        self.c.setLocalOveride("Functions", "ReturnVariableName", "_ret")	
        c_ret = convertVBtoPython(self.code2)
        self.c.setLocalOveride("Functions", "ReturnVariableName", "_other")	
        c_other = convertVBtoPython(self.code2)
        #
        # Should be different
        self.assertNotEqual(c_ret, c_other)
        #
        # But only tabs and spaces
        self.assertEqual(c_ret, c_other.replace("_other", "_ret"))
    # << Config tests >> (11 of 14)
    def testPreInitVariable(self):
        """testPreInitVariable: should be able to change if variable is pre initialized"""
        self.c.setLocalOveride("Functions", "ReturnVariableName", "_ret")	
        self.c.setLocalOveride("Functions", "PreInitializeReturnVariable", "Yes")	
        c_yes = convertVBtoPython(self.code2)
        self.c.setLocalOveride("Functions", "PreInitializeReturnVariable", "No")	
        c_no = convertVBtoPython(self.code2)
        #
        # Should be different
        self.assertNotEqual(c_yes, c_no)
        #
        # With init should have _ret = None, not without
        self.assertNotEqual(c_yes.find("_ret = None"), -1)
        self.assertEqual(c_no.find("_ret = None"), -1)
    # << Config tests >> (12 of 14)
    def testSelectVariable(self):
        """testSelectVariable: should be able to change select variable"""
        self.c.setLocalOveride("Select", "UseNumericIndex", "No")	
        self.c.setLocalOveride("Select", "SelectVariablePrefix", "_ret")	
        c_ret = convertVBtoPython(self.code3)
        self.c.setLocalOveride("Select", "SelectVariablePrefix", "_other")	
        c_other = convertVBtoPython(self.code3)
        #
        # Should be different
        self.assertNotEqual(c_ret, c_other)
        #
        # But only tabs and spaces
        self.assertEqual(c_ret, c_other.replace("_other", "_ret"))
    # << Config tests >> (13 of 14)
    def testSelectVariableIndex(self):
        """testSelectVariableIndex: should be able to turn off select variable index"""
        self.c.setLocalOveride("Select", "UseNumericIndex", "Yes")	
        self.c.setLocalOveride("Select", "SelectVariablePrefix", "_ret")	
        c_1 = convertVBtoPython(self.code3)
        self.c.setLocalOveride("Select", "UseNumericIndex", "No")	
        c_2 = convertVBtoPython(self.code3)
        #
        # Should be different
        self.assertNotEqual(c_1, c_2)
    # << Config tests >> (14 of 14)
    def testEvalVariable(self):
        """testEvalVariable: should be able to change whether variable is used once or more than once"""
        self.c.setLocalOveride("Select", "EvaluateVariable", "Once")	
        self.c.setLocalOveride("Select", "SelectVariablePrefix", "_ret")	
        c_1 = convertVBtoPython(self.code3)
        self.c.setLocalOveride("Select", "EvaluateVariable", "EachTime")	
        c_2 = convertVBtoPython(self.code3)
        #
        r = re.compile("_ret")
        self.assert_(len(r.findall(c_1)) > 1)	
        self.assertEqual(len(r.findall(c_2)), 0)
    # -- end -- << Config tests >>


if __name__ == "__main__":
    unittest.main()
