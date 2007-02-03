# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

#
# Turn off logging in extensions (too loud!)
import vb2py.extensions
vb2py.extensions.disableLogging()
import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff

from unittest import *

from vb2py.plugins.attributenames import TranslateAttributes

class TestAttributeNames(TestCase):

	def setUp(self):
		"""Setup the tests"""
		self.p = TranslateAttributes()

	# << Tests >>
	def testAll(self):
		"""Do some tests on the attribute"""
		names =(("Text", "text"),
				("Visible", "visible"),)
		for attribute, replaced in names:
			for pattern in ("a.%s=b", ".%s=b", "b=a.%s", "b=.%s",
							"a.%s.b=c", ".%s.c=b", "b=a.%s.c", "b=.%s.c",
							"a.%s.b+10=c", ".%s.c+10=b", "b=a.%s.c+10", "b=.%s.c+10",):
				test = pattern % attribute
				new = self.p.postProcessPythonText(test)
				self.assertEqual(new, pattern % replaced)
		for attribute, replaced in names:
			for pattern in ("a.%slkjlk=b", ".%slkjlk=b", "b=a.%slkjl", "b=.%slkjl",
							"a.%slkj.b=c", ".%slkj.c=b", "b=a.%slkj.c", "b=.%slkj.c",
							"a.%slkj.b+10=c", ".%slkj.c+10=b", "b=a.%slkj.c+10", "b=.%slkj.c+10",):
				test = pattern % attribute
				new = self.p.postProcessPythonText(test)
				self.assertNotEqual(new, pattern % replaced)
	# -- end -- << Tests >>

if __name__ == "__main__":
	main()
