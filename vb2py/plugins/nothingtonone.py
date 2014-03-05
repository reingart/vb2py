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
