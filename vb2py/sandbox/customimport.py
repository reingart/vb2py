# Created by Leo from: C:\Development\Python23\Lib\site-packages\vb2py\vb2py.leo

from vb2py.vbparser import *

p = VBProject()

m1 = VBCodeModule()
m1.parent = p

m2 = VBCodeModule()
m2.parent = p

c1 = parseVB("Sub a()\nb = This\nc=Node()\nDim d as Node\nEnd Sub", container=m1)
c2 = parseVB("' VB2PY-GlobalAdd: CustomIncludes.Comctllib = comctllib\nPublic Sub This()\nEnd Sub", container=m2)

print m1.renderAsCode()
print m2.renderAsCode()
