# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

try:
	import vb2py.extensions as extensions
except ImportError:
	import extensions

class ReplaceNothingWithNone(extensions.RenderHookPlugin, extensions.SystemPlugin):
	"""Plugin to replace Nothing with None in objects"""

	hooked_class_name = "VBObject"

	def addMarkup(self, indent, text):
		"""Add markup to the rendered text"""
		return text.replace("Nothing", "None")
