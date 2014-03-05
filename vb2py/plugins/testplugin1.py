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
