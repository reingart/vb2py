# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

try:
	import vb2py.extensions as extensions
except ImportError:
	import extensions


class TestREPlugin(extensions.RETextMarkup):
	"""An example plugin"""    

	name = "REPlugin"

	pre_process_patterns = (
			("(?P<Object>.*)_Click", "%(Object)s_click"),
			("\sError\s", " _errfn "),
	)    


class NotAPlugIn:
	"""Something that isn't a plugin"""
