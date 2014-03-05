"""The main form for the application"""

from PythonCard import model

# Allow importing of our custom controls
import PythonCard.resource
PythonCard.resource.APP_COMPONENTS_PACKAGE = "vb2py.targets.pythoncard.vbcontrols"

class Background(model.Background):

    def __getattr__(self, name):
        """If a name was not found then look for it in components"""
        return getattr(self.components, name)


    def __init__(self, *args, **kw):
        """Initialize the form"""
        model.Background.__init__(self, *args, **kw)
        # Call the VB Form_Load
        # TODO: This is brittle - depends on how the private indicator is set
        if hasattr(self, "_Background__Form_Load"):
            self._Background__Form_Load()
        elif hasattr(self, "Form_Load"):
            self.Form_Load()


# CODE_GOES_HERE


if __name__ == '__main__':
    app = model.Application(Background)
    app.MainLoop()
