try:
    import vb2py.extensions as extensions
except ImportError:
    import extensions

class TranslateAttributes(extensions.SystemPluginREPlugin):
    """Plugin to convert attribute names from VB to Pythoncard

    There are attribute like 'Text' and 'Visible' which are in lower
    case in Pythoncard and others are simply different. We do the conversion
    here. 

    Note that this means we will convert these names even if they don't belong
    to controls - this is unfortunate but still safe as we do the conversion
    consistently.

    """

    name = "PlugInAttributeNames"
    __enabled = 0   # If false the plugin will not be called

    post_process_patterns = (
#			(r"\.Text\b", ".text"),
#			(r"\.Caption\b", ".text"),
#			(r"\.Visible\b", ".visible"),
#			(r"\.Enabled\b", ".enabled"),
#			(r"\.BackColor\b", ".backgroundColor"),
#			(r"\.ToolTipText\b", ".ToolTipText"),
#			(r"\.AddItem\b", ".append"),
    )
