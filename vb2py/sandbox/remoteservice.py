"""Test of using vb2py as a remote service"""

import urllib
import re

SERVICE = "http://vb2py.sourceforge.net/cgi-bin/remote.py"
EXTRACT = re.compile(r"\<(\S+)\>(.*)\</\1\>", re.DOTALL)

code = """
Select Case a
  Case 10
    b=1
  Case 10,20
    b=2
  Case 30 To 40
    b=3
  Case Else
    b=4
End Select

"""

result = urllib.urlopen("%s?code=%s" % (SERVICE, urllib.quote(code)))
text = result.read()
parts = EXTRACT.match(text)

if parts:
    print "%s\n%s" % parts.groups()
else:
    print "Unable to decifer result! (%s)" % text
