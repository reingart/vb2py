# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

try:
	import vb2py.extensions as extensions
except ImportError:
	import extensions

class RemoveTypeMarkers(extensions.SystemPluginREPlugin):
	"""Plugin to remove the type identifiers from functions

	Some VB functions have $, %, #, & markers at the end of their names and we
	need to remove them. The proper place to do this is in the parsing but this
	is a quick fix!

	"""

	name = "Remove $ from functions"

	post_process_patterns = (
			(r"Left\$\(", "Left("),
			(r"Right\$\(", "Right("),
			(r"Mid\$\(", "Mid("),
			(r"Chr\$\(", "Chr("),	
			(r"Dir\$\(", "Dir("),	
			(r"Trim\$\(", "Dir("),	
	)
