try:
    import vb2py.extensions as extensions
except ImportError:
    import extensions


class LineContinuations(extensions.SystemPlugin):
    """Plugin to handle line continuations

    Line continuations are indicated by a '_' at the end of a line and imply that
    the current line and the one following should be joined together. We could
    parse this out in the grammar but it is just easier to handle it as a pre-processor
    text as we aren't going to use it in the Python conversion.

    """

    order = 10 # We would like to happen quite early

    def preProcessVBText(self, txt):
        """Convert continuation markers by joining adjacent lines"""

        txt_lines = txt.split("\n")
        txtout = "\n".join([lne.strip() for lne in txt_lines if lne.strip()])
        txtout = txtout.replace(" _\n", " ")
        txtout += "\n\n"
        self.log.info("Line continuation:\nConverted '%s'\nTo '%s'" % (txt, txtout))
        return txtout
