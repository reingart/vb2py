from vb2py.vbfunctions import *
from vb2py.vbdebug import *
import Point

class Point(Object):

    x = Single()
    y = Single()
    SubPoint = Point.Point()

    def getLength(self):
        _ret = None
        _ret = Sqr(self.x ** 2 + self.y ** 2)
        return _ret

    # VB2PY (UntranslatedCode) Attribute VB_Name = "Point"
    # VB2PY (UntranslatedCode) Attribute VB_GlobalNameSpace = False
    # VB2PY (UntranslatedCode) Attribute VB_Creatable = True
    # VB2PY (UntranslatedCode) Attribute VB_PredeclaredId = False
    # VB2PY (UntranslatedCode) Attribute VB_Exposed = False
