# Created by Leo from: C:\Development\Python23\Lib\site-packages\vb2py\vb2py.leo

from utils import rootPath
import ConfigParser
import os

class VB2PYConfigObject(dict):
    """A dictionary of configuration options

    Options are accessed via a tuple
    c[section, key]

    """

    def __init__(self, *args, **kw):
        """Initialize the dictionary"""
        self.initConfig(*args, **kw)

    def __getitem__(self, key):
        """Get an item"""
        try:
            return self._local_overide["%s.%s" % key]
        except KeyError:
            return self._config.get(*key)

    def initConfig(self, filename="vb2py.ini", path=None):
        """Read the values"""
        if path is None:
            path = rootPath()
        self._config = ConfigParser.ConfigParser()
        self._config.read(os.path.join(path, filename))
        self._local_overide = {}	

    def setLocalOveride(self, section, name, value):
        """Set a local overide for a value"""
        self.checkValue(section, name)
        self.addLocalOveride(section, name, value)

    def addLocalOveride(self, section, name, value):
        """Add a local overide for a value"""
        if self._config.has_section(section):
            self._local_overide["%s.%s" % (section, name)] = value
        else:
            raise ConfigParser.NoSectionError("No such section '%s'" % section)

    def removeLocalOveride(self, section, name,):
        """Remove a local overide"""
        self.checkValue(section, name)
        del(self._local_overide["%s.%s" % (section, name)])        

    def checkValue(self, section, name):
        """Make sure we have this section and name"""
        dummy = self[section, name]    

    def getItemNames(self, section):
        """Return the list of items in a section"""
        base = self._config.options(section) 
        local = []
        for name in self._local_overide:
            section_name, option = name.split(".")
            if section_name == section:
                local.append(option)
        return base + local
#
# We always want people to use the same one
_VB2PYConfig = VB2PYConfigObject()

def VB2PYConfig(init=0):
    ret = _VB2PYConfig
    if init:
        ret.initConfig()
    return ret
