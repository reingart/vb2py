test = """
<html>
<%

function factorial(x)
if x = 0 then
    factorial = 1
else
    factorial = x*factorial(x-1)
end if
end function

%>
</html>
"""

from vb2py.vbparser import parseVB, VBCodeModule
import re

def translateScript(match):
    """Translate VBScript fragment to Python"""
    block = parseVB(match.groups()[0], container=VBCodeModule())
    return "<%%\n%s\n%%>" % block.renderAsCode()

converter = re.compile(r"\<%(.*?)%\>", re.DOTALL + re.MULTILINE)

print converter.sub(translateScript, test)
