# Created by Leo from: C:\Development\Python23\Lib\site-packages\vb2py\vb2py.leo

import re	    # For text processing
import os     # For file processing
import pprint # For outputting dictionaries
import sys    # For getting Exec prefix
import getopt # For command line arguments

# TODO: refactor out this ugliness

from vb2py.converter import BaseResource
from vb2py import vbparser
from controls import *


twips_per_pixel = 15

# << Event translation >>
#	This is the mapping between VB events and PythonCard events.

event_translator = {
		"Click" : "mouseClick",
}
# -- end -- << Event translation >>
# << Resources >>
class Resource(BaseResource):
	"""Represents a Python Card resource object"""

	# << PyCardResource declarations >>
	target_name = "pythoncard"
	name = "basePyCardResource"

	form_class_name = "MAINFORM"
	form_super_classes = ["Background"]
	allow_new_style_class = 0
	# -- end -- << PyCardResource declarations >>
	# << class PyCardResource methods >> (1 of 2)
	def __init__(self, *args, **kw):
		"""Initialize the PythonCard resource"""
		BaseResource.__init__(self, *args, **kw)
		self._rsc = eval(open("%s.txt" % self.basesourcefile, "r").read().replace("\r\n", "\n"))
		self._code = open("%s.py" % self.basesourcefile, "r").read()
	# << class PyCardResource methods >> (2 of 2)
	def writeToFile(self, basedir, write_code=0):
		"""Write ourselves out to a directory"""
		# << Resource file >>
		fle = open(os.path.join(basedir, self.name) + ".rsrc.py", "w")
		log.info("Writing '%s'" % os.path.join(basedir, self.name) + ".rsrc.py")
		pprint.pprint(self._rsc, fle)
		fle.close()
		# -- end -- << Resource file >>
		# << Code file >>
		fle = open(os.path.join(basedir, self.name) + ".py", "w")
		log.info("Writing '%s'" % os.path.join(basedir, self.name) + ".py")

		if write_code:
			added_code = vbparser.renderCodeStructure(self.code_structure)
		else:
			added_code = ""

		self._code = self._code.replace("# CODE_GOES_HERE", added_code)

		fle.write(self._code)


		fle.close()
		# -- end -- << Code file >>
	# -- end -- << class PyCardResource methods >>
# -- end -- << Resources >>
