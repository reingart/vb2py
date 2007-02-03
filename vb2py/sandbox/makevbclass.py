"""Make a "shim" VB class to call a Python COM class

The sole purpose of the shim is to allow VB's intellisense to work. The real way to
do this is to generate a type library file - but that appears to be a bit tricky and
requires access to the MIDL compiler.

This approach is cheap and easy (cheasy?)

"""

# << Declarations >>
import inspect
import sys
# -- end -- << Declarations >>
# << Error classes >>
class MakeVBError(Exception): """Base class for all errors in this module"""
# -- end -- << Error classes >>
# << Functions >> (1 of 3)
def makeVBClass(cls, init=None):
    """Returns the text of a VB class file wrapping the specified class

    If provided, then the init parameter is used as the initializer for the class
    so that you can link to the real Python object

    """
    vb = [] # Build up as list and then .join it later
    #
    # Get the method names to deal with
    method_names = dir(cls)
    method_names.sort()
    #
    # Add initializer
    if init:
        vb.extend([
            "Private myObj as Object",
            "",
            "Sub Class_Initialize()",
            init,
            "End Sub",
            ""
            ])
    #
    # Build VB
    for name in method_names:
        if not name.startswith("_"):
            vb.extend(makeVBMethod(cls, name))
    #
    return "\n".join(vb)
# << Functions >> (2 of 3)
def makeVBClassFile(cls, filename, init=None): 
    """Make a VB Class file froma class"""
    text = makeVBClass(cls, init)
    f = file(filename, "w")
    f.write(text)
    f.close()
# << Functions >> (3 of 3)
def makeVBMethod(cls, method_name): 
    """Make some VB to support the specified method"""
    vb = []
    method = getattr(cls, method_name)
    try:
        args, vararg, varkw, defaults = inspect.getargspec(method)
    except TypeError:
        return [] # Not a method
    #
    # Get docstring, if any
    if method.__doc__:
        doc = "\n".join(["   ' %s" % line for line in method.__doc__.split("\n")])
    else:
        doc = "   ' %s method" % method_name
    #
    # Get the argument list (remember to skip argument 0, which is self)
    if not defaults:
        arg_list = args[1:]
    else:
        arg_list = args[1:-len(defaults)]
        for arg, value in zip(args[-len(defaults):], defaults):
            if value is None:
                arg_list.append("Optional %s=%s" % (arg, "Nothing"))
            else:
                arg_list.append("Optional %s=%s" % (arg, repr(value)))				
    #
    vb.extend([
            "",
            "Public Function %s(%s)" % (method_name, ", ".join(arg_list)),
            doc,
            "   %s = myObj.%s(%s)" % (method_name, method_name, ", ".join(args[1:])),
            "End Function",
            ])
    #
    return vb
# -- end -- << Functions >>

if __name__ == "__main__":
    # << In situ-testing >>
    if len(sys.argv) <> 3:
        print "Usage: makevbclass <class>, filename\n"
    else:
        class_name, filename = sys.argv[1:]
        imports = class_name.split(".")
        if len(imports) > 1:
            exec "from %s import %s" % (".".join(imports[:-1]), imports[-1])
        cls = eval(imports[-1])
        #
        print "Using %s" % cls
        makeVBClassFile(cls, filename)
    # -- end -- << In situ-testing >>
