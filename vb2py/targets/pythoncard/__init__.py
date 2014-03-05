from PythonCard import registry
PythonCardRegistry = registry.Registry.getInstance()

def Register(control):
    """Register a control for PythonCard"""
    #
    #import pdb; pdb.set_trace()
    PythonCardRegistry.register(control)
