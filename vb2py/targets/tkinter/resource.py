# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

import vb2py
import os
from targets.tkinter.controls import *

# << Event translation >>
#	This is the mapping between VB events and Tkinter events.

event_translator = {
		"Click" : "mouseClick",
}
# -- end -- << Event translation >>

class Resource(vb2py.BaseResource):
	"""Represents a Tkinter resource object"""

	target_name = "Tkinter"

	# << class Resource methods >> (1 of 2)
	def __init__(self, *args, **kw):
		"""Initialize the PythonCard resource"""
		print "Resource init"
		vb2py.BaseResource.__init__(self, *args, **kw)
		print "After supre"
		self._rsc = eval(open("%s.txt" % self.basesourcefile, "r").read())
		self._rsc["controls"] = []
		self._code = open("%s.py" % self.basesourcefile, "r").read()
	# << class Resource methods >> (2 of 2)
	def writeToFile(self, basedir, write_code=0):
		"""Write ourselves out to a directory"""

		fle = open(os.path.join(basedir, self.name) + ".py", "w")

		lines = []
		for control in self._rsc['stack']['backgrounds'][0]['components']:
			control

		if write_code:
			#
			# Assemble our code

			for block in self.subs + self.fns:
				lines.append('    def %s(self, %s):\n        """Sub"""' % (block.name, block.args))
				for code_line in block.code.splitlines():
					lines.append("        %s" % code_line)
				lines.append("")

			added_code = "\n".join(lines)
		else:
			added_code = ""

		fle.write(self._code.replace("# CODE_GOES_HERE", added_code))
		fle.close()
	# -- end -- << class Resource methods >>
