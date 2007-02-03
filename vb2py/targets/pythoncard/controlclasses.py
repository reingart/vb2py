# Created by Leo from: C:\Development\Python23\Lib\site-packages\vb2py\vb2py.leo

"""Classes to support mapping of VB Control Properties to PythonCard

The classes fall into two types,
- a MetaClass VBWrapped to create wrapped classes
- a ProxyClass to sit between the PythonCard class and the Mimicked VB one

Although it may appear to be simpler to just subclass the PythonCard classes there
are two things that work against it. The main thing is that the PythonCard classes
must be old style classes. This means that you cannot use properties, which are
required to map things like .Top, .Width to .position[0] etc.

Since you can't use properties you have to use a __getattr__ hook. Unfortunately this
turns out to be really slow (probably because of an interaction between this hook
and a hook in a lower class which is doing a similar thing.

The solution is to use a proxy class (VBWidget) which hands off most references to 
the PythonCard class. The VBWidget is a new style class and can therefore use properties.

The metaclass (VBWrapped) is used to automatically generate properties for names which
are similar ('Text' -> 'text') and names which require mapping ('Left' -> 'position[1]').
A metaclass solution may not be absolutely necessary but it seems to speed things up
by doing the class manipulation once (at import) rather than for each control as it is
created.

Typical usage is as follows

class VBTextField(VBWidget): 
    __metaclass__ = VBWrapped 

    _translations = { 
            "Text" : "text", 
            "Enabled" : "enabled", 
            "Visible" : "visible", 
        } 

    _indexed_translations = { 
            "Left" : ("position", 0), 
            "Top" : ("position", 1), 
            "Width" : ("size", 0), 
            "Height" : ("size", 1), 
        } 

    _proxy_for = textfield.TextField


"""

import new 
from vb2py.vbclasses import VBArray
import vb2py.logger
log = vb2py.logger.getLogger("VBWidget")

# << classes >> (1 of 2)
class VBWidget(object): 

    _translations = {}
    _name_to_method_translations = {}
    _indexed_translations = {}
    _method_translations = {}

    def __init__(self, *args, **kw):
        self.__dict__["_proxy"] = self.__class__._proxy_for(*args, **kw)

    def __getattr__(self, name): 
        return getattr(self._proxy, name) 

    def __setattr__(self, name, value): 
        if name in self._setters: 
            self._setters[name](self, value) 
        else: 
            try:
                setattr(self._proxy, name, value) 
            except AttributeError:
                log.debug("Setting local attribute '%s' for obj %s" % (name, self.__class__.__name__))
                self.__dict__[name] = value

    def __getitem__(self, name): 
        return self._proxy[name] 

    def __setitem__(self, name, value): 
        self._proxy[name] = value
# << classes >> (2 of 2)
class VBWrapped(type): 
    """A meta class to wrap PythonCard classes so VB converted code can use them""" 

    # << VBWrapped methods >> (1 of 6)
    def __new__(cls, name, bases, dict): 
        obj = type.__new__(cls, name, bases, dict) 
        obj._setters = {} 
        # Ordinary properties 
        for prop_name in obj._translations: 
            get, set = cls.createProperties(prop_name) 
            setattr(obj, prop_name, property(get, set)) 
            obj._setters[prop_name] = set 
        # Indexed properties 
        for prop_name in obj._indexed_translations: 
            get, set = cls.createIndexedProperties(prop_name) 
            setattr(obj, prop_name, property(get, set)) 
            obj._setters[prop_name] = set 
        # Method names 
        for method_name in obj._method_translations: 
            setattr(obj, method_name, cls.createMethodLookup(method_name)) 
        # Attributes which are properties in PythonCard
        for attr_name in obj._name_to_method_translations: 
            set = cls.createAttributeSet(attr_name)
            setattr(obj, attr_name, property(fget=cls.createAttributeLookup(attr_name),
                                             fset=set))
            obj._setters[attr_name] = set 


        # Set the _spec for the item
        obj._spec = obj._proxy_for._spec
        obj._spec.name = obj.__name__
        # Create the object 
        return obj
    # << VBWrapped methods >> (2 of 6)
    def createProperties(prop_name): 
        def set(obj, v): 
            setattr(obj, obj._translations[prop_name], v) 
        def get(obj): 
            item = getattr(obj, obj._translations[prop_name]) 
            # Wrap up list types into something we can handle
            if type(item) == type([]):
                return VBArray.createFromData(item)
            else:
                return item
        return get, set        

    createProperties = staticmethod(createProperties)
    # << VBWrapped methods >> (3 of 6)
    def createIndexedProperties(prop_name): 
        def set(obj, v): 
            attr_name, index = obj._indexed_translations[prop_name] 
            lst = list(getattr(obj, attr_name)) 
            lst[index] = v 
            setattr(obj, attr_name, lst) 
        def get(obj): 
            attr_name, index = obj._indexed_translations[prop_name] 
            return getattr(obj, attr_name)[index] 
        return get, set        

    createIndexedProperties = staticmethod(createIndexedProperties)
    # << VBWrapped methods >> (4 of 6)
    def createMethodLookup(method_name):
        def callMethod(obj, *args, **kw):
            return getattr(obj, obj._method_translations[method_name])(*args, **kw)
        return callMethod

    createMethodLookup = staticmethod(createMethodLookup)
    # << VBWrapped methods >> (5 of 6)
    def createAttributeLookup(attr_name):
        def get(obj):
            return getattr(obj, obj._name_to_method_translations[attr_name][0])()
        return get

    createAttributeLookup = staticmethod(createAttributeLookup)
    # << VBWrapped methods >> (6 of 6)
    def createAttributeSet(attr_name):
        def set(obj, value):
            getattr(obj, obj._name_to_method_translations[attr_name][1])(value)
        return set

    createAttributeSet = staticmethod(createAttributeSet)
    # -- end -- << VBWrapped methods >>
# -- end -- << classes >>
