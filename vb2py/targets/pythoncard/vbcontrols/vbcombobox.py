from vb2py.targets.pythoncard.controlclasses import VBWrapped, VBWidget
from vb2py.targets.pythoncard import Register
import vb2py.logger
log = vb2py.logger.getLogger("VBComboBox")

from PythonCard.components import combobox
import wx
import sys
from PythonCard import event, registry, widget


class VBComboBox(VBWidget): 
    __metaclass__ = VBWrapped 

    _translations = { 
            "Text" : "text", 
            "Enabled" : "enabled", 
            "Visible" : "visible", 
            "List" : "items",
    } 

    _name_to_method_translations = {
            "ListCount" : ("getNumber", None),
            "ListIndex" : ("getSelectionIndex", None),
    }

    _indexed_translations = { 
            "Left" : ("position", 0), 
            "Top" : ("position", 1), 
            "Width" : ("size", 0), 
            "Height" : ("size", 1), 
    } 

    _method_translations = {			
            "Clear" : "clear",
            "RemoveItem" : "delete",	
    }

    _proxy_for = combobox.ComboBox

    # << VBComboBox methods >> (1 of 4)
    def AddItem(self, item, position=None):
        """Add an item to the list

        We cannot just map this to a PythonCard control event because it only has
        an 'append' and an 'insertItems' method, which isn't exactly the same

        """
        if position is None:
            self.append(item)
        else:
            self.insertItems([item], position)
    # << VBComboBox methods >> (2 of 4)
    def getNumber(self):
        """Get the number of items in the Combo

        This doesn't appear to be in the PythonCard control

        """
        return len(self.items)
    # << VBComboBox methods >> (3 of 4)
    def getSelectionIndex(self):
        """Get the index of the currently selected item

        This doesn't appear to be in the PythonCard control

        """
        try:
            return self.items.index(self.selection)
        except ValueError:
            return -1
    # << VBComboBox methods >> (4 of 4)
    def delete(self, position):
        """Remove the specified item from the Combo

        This doesn't appear to be in the PythonCard control

        """
        del(self.items[position]) # TODO - this doesn't actually work
    # -- end -- << VBComboBox methods >>   


log.debug("Registering VBComboBox as '%s'" % sys.modules[__name__].VBComboBox)
Register(VBComboBox)
