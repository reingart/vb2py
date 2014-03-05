# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

# << basesource declarations >>
"""
__version__ = "$Revision: 1.2 $"
__date__ = "$Date: 2003/08/04 00:09:15 $"
"""

from PythonCardPrototype import model

#
# VB constants
True = 1
False = 0
# -- end -- << basesource declarations >>
# << basesource methods >>
class MAINFORM(model.Background):
	# << class MAINFORM declarations >>
	"""The main form for the application"""

	    def cmdFromFile_Click(self, ):
        """Sub"""
            t = 0
            for i in vbForRange(0, 100):
                t = t + i

	# -- end -- << class MAINFORM declarations >>
# -- end -- << basesource methods >>

if __name__ == '__main__':
	app = model.PythonCardApp(MAINFORM)
	app.MainLoop()
