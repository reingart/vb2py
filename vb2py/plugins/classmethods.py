# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

try:
	import vb2py.extensions as extensions
except ImportError:
	import extensions

class RemoveTypeMarkers(extensions.SystemPluginREPlugin):
	"""Plugin to replace Class Method names with their Python equivalents

	This could be done in the parser but is done here to expose the translation and
	allow it to be customized.

	"""

	post_process_patterns = (
			(r"Class_Initialize\(", "__init__("),
			(r"Class_Terminate\(", "__del__("),
	)
