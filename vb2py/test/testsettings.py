# Created by Leo from: C:\Development\Python23\Lib\site-packages\vb2py\vb2py.leo

from unittest import *
from testframework import *

# << Settings tests >> (1 of 3)
# Simple test of get with default
tests.append(("""
a = GetSetting("vbtest", "main", "key", "<default>")
""", {"a":"<default>"}))
tests.append(("""
a = GetSetting("vbtestother", "main", "key", "<default>")
""", {"a":"<default>"}))

# Simple test of set then get 
tests.append(("""
SaveSetting "vbtest", "main", "real", 10.5
SaveSetting "vbtest", "main", "int", 1
SaveSetting "vbtest", "main", "string", "hello"
a = GetSetting("vbtest", "main", "real")
b = GetSetting("vbtest", "main", "int")
c = GetSetting("vbtest", "main", "string")
""", {"a":"10.5", "b":"1", "c":"hello"})) # always returned as a string

# Simple test of set then get with default
tests.append(("""
SaveSetting "vbtest", "main", "real", 10.5
SaveSetting "vbtest", "main", "int", 1
SaveSetting "vbtest", "main", "string", "hello"
a = GetSetting("vbtest", "main", "real", "mmm")
b = GetSetting("vbtest", "main", "int", "mmm")
c = GetSetting("vbtest", "main", "string", "mmm")
""", {"a":"10.5", "b":"1", "c":"hello"})) # always returned as a string
# << Settings tests >> (2 of 3)
# Simple test of set then getall 
tests.append(("""
SaveSetting "vbtest", "main", "real", 10.5
SaveSetting "vbtest", "main", "int", 1
SaveSetting "vbtest", "main", "string", "hello"
Dim _Setting, _Settings
a=""
_Settings = GetAllSettings("vbtest", "main")
For _Setting = 0 To UBound(_Settings)
	a = a & _Settings(_Setting, 0) & " = " & _Settings(_Setting, 1) & ":"
Next _Setting

""", {"a":"real = 10.5:string = hello:int = 1:"})) # always returned as a string
# << Settings tests >> (3 of 3)
# Simple test of delete
tests.append(("""
a = GetSetting("vbtest", "second", "key", "<default>")
SaveSetting "vbtest", "second", "key", "hello"
b = GetSetting("vbtest", "second", "key", "<default>")
DeleteSetting "vbtest", "second", "key"
c = GetSetting("vbtest", "second", "key", "<default>")
""", {"a":"<default>", "b":"hello", "c":"<default>"}))
# -- end -- << Settings tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
	main()
