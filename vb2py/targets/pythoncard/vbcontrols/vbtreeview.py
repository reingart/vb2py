# Created by Leo from: C:\Development\Python23\Lib\site-packages\vb2py\vb2py.leo

from wx import TreeItemData
from wxPython import wx, stc
import vb2py.custom.comctllib

from vb2py.targets.pythoncard.controlclasses import VBWrapped, VBWidget
import vb2py.logger
log = vb2py.logger.getLogger("VBTreeView")

from PythonCardPrototype.components import tree
from wxPython import wx
import sys
from PythonCardPrototype import binding, event, registry, widget
from vb2py.vbclasses import Collection




# << Classes >> (1 of 2)
class VBTreeView(VBWidget): 

    # << class VBTreeView declarations >>
    __metaclass__ = VBWrapped 

    _translations = { 
            "ListImages" : "items",
            "Enabled" : "enabled", 
            "Visible" : "visible", 
    } 

    _name_to_method_translations = {
            "ListCount" : "getNumber",
            "ListIndex" : "getSelectionIndex",
    }

    _indexed_translations = { 
            "Left" : ("position", 0), 
            "Top" : ("position", 1), 
            "Width" : ("size", 0), 
            "Height" : ("size", 1), 
    } 

    _method_translations = {			
    }

    _proxy_for = tree.Tree
    # -- end -- << class VBTreeView declarations >> 
    # << class VBTreeView methods >> (1 of 2)
    def __init__(self, *args, **kw):
        """Initialize the tree view"""
        super(VBTreeView, self).__init__(*args, **kw)
        self.Nodes = TreeNodeCollection(self)
    # << class VBTreeView methods >> (2 of 2)
    def _getSelectedItem(self): 
        """Getting the selected item"""
        return vb2py.custom.comctllib.Node(self.GetSelection(), self)

    def _setSelectedItem(self, item): 
        """Setting the selected item"""
        self.SelectItem(item._id)

    SelectedItem = property(fget=_getSelectedItem,
                            fset=_setSelectedItem)
    # -- end -- << class VBTreeView methods >>
# << Classes >> (2 of 2)
class TreeNodeCollection(Collection): 
    """Represents a collection of nodes in a tree view""" 

    # << class TreeNodeCollection declarations >>
    pass
    # -- end -- << class TreeNodeCollection declarations >> 
    # << class TreeNodeCollection methods >> (1 of 5)
    def __init__(self, parent): 
        """Initialise the TreeNodeCollection instance"""
        super(TreeNodeCollection, self).__init__()
        self._parent = parent
        self._nodes = {}
        self._initTree()
    # << class TreeNodeCollection methods >> (2 of 5)
    def _initTree(self): 
        """Initialize the tree"""
        self._nodes["<vbtreeroot>"] = self._parent.AddRoot("Root", data=TreeItemData("<vbtreeroot>"))    
        self._parent.SetPyData(self._nodes["<vbtreeroot>"], "<vbtreeroot>")
    # << class TreeNodeCollection methods >> (3 of 5)
    def Clear(self): 
        """Clear all the nodes"""
        self._parent.DeleteAllItems()
        self._initTree()
    # << class TreeNodeCollection methods >> (4 of 5)
    def Add(self, Relative=None, Relationship=vb2py.custom.comctllib.tvwChild, 
            Key="", Text="", Image=None, SelectedImage=None): 
        """Add a node to the tree"""
        if Relative is None:
            id = self._nodes["<vbtreeroot>"]
        elif Relationship == vb2py.custom.comctllib.tvwChild:
            id = self._nodes[Relative]
        else:
            raise NotImplementedError("Tree Add not implemented for relationships other than tvwChild")
        #
        self._nodes[Key] = self._parent.AppendItem(id, Text)
        self._parent.SetPyData(self._nodes[Key], Key)   
        self._parent.SetItemHasChildren(id, True)
    # << class TreeNodeCollection methods >> (5 of 5)
    def __iter__(self): 
        """Return an iterator over the nodes"""
        for node in self._nodes.values():
            yield vb2py.custom.comctllib.Node(node, self._parent)
    # -- end -- << class TreeNodeCollection methods >>
# -- end -- << Classes >>

log.debug("Registering VBTreeView as '%s'" % sys.modules[__name__].VBTreeView)
registry.getRegistry().register( sys.modules[__name__].VBTreeView )
